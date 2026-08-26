/**
 * The worksheet merge, as a part-level transplant.
 *
 * The engine this replaces rebuilt every worksheet through a spreadsheet
 * library's object model, so a merged workbook kept only what that model
 * happened to represent: cells, and almost nothing else. This one starts from
 * the first input's package and COPIES the later inputs' worksheet parts into
 * it, with the parts they depend on, rewriting only the four things a copy
 * genuinely invalidates:
 *
 * 1. part paths and relationship ids, which must not collide;
 * 2. shared-string indexes, which point into a per-workbook table;
 * 3. style and differential-format indexes, which do too; and
 * 4. the NAMES a formula qualifies a reference with, when a collision forced
 *    a worksheet or an Excel Table to be renamed.
 *
 * Everything else - conditional formatting, data validation, merged ranges,
 * hyperlinks, comments and their drawings, Excel Tables, defined names, row
 * heights, column widths, number formats - travels because it is never
 * touched. That is the whole design: preservation is the default, and the
 * exceptions are enumerated below and declared in `src/contract.ts`.
 *
 * NOT carried, deliberately:
 * - pivot tables and their caches: a cache is a private copy of source rows
 *   whose `cacheId` registry is workbook-scoped; stripping matches the split's
 *   Tier-1 policy (`src/tier1/pivot.ts`). Strip, with a warning.
 * - external links: an external reference is addressed by its index in the
 *   workbook's own `externalReferences` list, and formulas cite that index
 *   positionally, so two inputs' links cannot be interleaved. Strip, warn.
 * - the calculation chain: a derived index keyed by sheet id, and every
 *   transplanted sheet takes a new one. Dropping it is the fix - Excel
 *   rebuilds the chain on open and the merged workbook asks it to, via
 *   `calcPr fullCalcOnLoad`. No warning: nothing authored is lost.
 * - the macro project, unless exactly one input carries one AND the output is
 *   named as a macro-enabled workbook. See `resolveMacroProject`.
 */
import { ConsultChimpsError } from "@consultchimps/core";

import { XLSX_ERRORS } from "../errors.js";
import { readExcelTableDefinitionsFrom } from "../excel-tables.js";
import {
  addAttribute,
  decodeXmlText,
  editElements,
  findElement,
  findElements,
  getAttribute,
  setAttribute,
} from "../model/xml.js";
import {
  MACRO_WORKBOOK_MAIN_CONTENT_TYPE,
  packagePartDirectory,
  relationshipsPartPath,
  VBA_PROJECT_PART,
  WORKBOOK_MAIN_CONTENT_TYPE,
  WorkbookPackage,
  type PackageRelationship,
} from "../package/index.js";
import { convertWorkbookToValues } from "../values-only.js";
import {
  addContentTypeOverride,
  addRelationship,
  allocatePartPath,
  copyContentTypeDeclaration,
  escapeXmlAttribute,
  escapeXmlText,
  partExtension,
  partStem,
  relativeRelationshipTarget,
  setContentTypeOverride,
  writeRelationships,
  WORKBOOK_PART,
} from "./ooxml.js";
import { rewriteFormulaElements, type NameRewrites } from "./references.js";
import { MergedStyles, SourceStyles } from "./styles.js";
import { SharedStringTable } from "./strings.js";
import {
  rewriteDifferentialFormatIds,
  rewriteRelationshipIds,
  rewriteWorksheetIndexes,
  type IndexRemaps,
} from "./worksheet.js";

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NAMESPACE_MAIN =
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELATIONSHIP_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const STYLES_PART = "xl/styles.xml";
const SHARED_STRINGS_PART = "xl/sharedStrings.xml";
const CALC_CHAIN_PART = "xl/calcChain.xml";

const SHARED_STRINGS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const WORKSHEET_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

const MAXIMUM_SHEET_NAME_LENGTH = 31;
const SHEET_INDEX_NAME = "Sheet Index";
const SHEET_INDEX_HEADERS = [
  "Source file",
  "Original worksheet",
  "Final worksheet",
  "Source visibility",
] as const;

/** Part prefixes a merged workbook never carries. */
const DROPPED_PART_PREFIXES = [
  "xl/pivottables/",
  "xl/pivotcache/",
  "xl/externallinks/",
] as const;

export interface MergeWorkbooksBuildOptions {
  includeSheetIndex?: boolean | undefined;
  values?: boolean | undefined;
  /**
   * Whether the surface will name the output as a macro-enabled workbook. A
   * macro project only travels into an output that says so, because a package
   * whose content type contradicts its file name is one Excel warns about.
   */
  macroOutput?: boolean | undefined;
}

export interface MergedWorkbook {
  bytes: Uint8Array;
  hiddenSheets: number;
  outputSheets: number;
  warnings: string[];
  /** True when the output kept a macro project and must be named `.xlsm`. */
  macroEnabled: boolean;
}

interface OutputSheet {
  /** The name the sheet has in the merged workbook. */
  name: string;
  visibility: 0 | 1 | 2;
}

interface CarriedDefinedName {
  name: string;
  /** Index into the merged workbook's sheet list, for sheet-scoped names. */
  localSheetId: number | undefined;
  reference: string;
}

export interface MergeWorkbooksState {
  readonly options: MergeWorkbooksBuildOptions;
  /** The output package, seeded by the first input that loads. */
  output: WorkbookPackage | undefined;
  styles: MergedStyles | undefined;
  strings: SharedStringTable | undefined;
  /** Lower-cased worksheet names already claimed in the merged workbook. */
  readonly usedSheetNames: Set<string>;
  /** Lower-cased Excel Table names already claimed in the merged workbook. */
  readonly usedTableNames: Set<string>;
  /** Lower-cased workbook-scoped defined names already claimed. */
  readonly usedDefinedNames: Set<string>;
  readonly definedNames: CarriedDefinedName[];
  readonly sheets: OutputSheet[];
  readonly indexRows: string[][];
  hiddenSheets: number;
  outputSheets: number;
  nextSheetId: number;
  /** Inputs that carry a macro project, whether or not it is kept. */
  macroInputs: number;
  /** Whether the first input - the only one whose project can travel - has one. */
  seedHasMacroProject: boolean;
  removedPivotTables: number;
  removedExternalLinks: number;
  /** The next free Excel Table id; ids are unique across the workbook. */
  nextTableId: number;
  readonly renamedTables: Array<[string, string]>;
  readonly renamedDefinedNames: Array<[string, string]>;
}

export function createMergeState(
  options: MergeWorkbooksBuildOptions,
): MergeWorkbooksState {
  return {
    options,
    output: undefined,
    styles: undefined,
    strings: undefined,
    usedSheetNames: new Set<string>(
      options.includeSheetIndex === false
        ? []
        : [SHEET_INDEX_NAME.toLocaleLowerCase()],
    ),
    usedTableNames: new Set<string>(),
    usedDefinedNames: new Set<string>(),
    definedNames: [],
    sheets: [],
    indexRows: [[...SHEET_INDEX_HEADERS]],
    hiddenSheets: 0,
    outputSheets: 0,
    nextSheetId: 1,
    macroInputs: 0,
    seedHasMacroProject: false,
    removedPivotTables: 0,
    removedExternalLinks: 0,
    nextTableId: 1,
    renamedTables: [],
    renamedDefinedNames: [],
  };
}

// -- Reading the source ----------------------------------------------------

/** A `<sheet>` declaration, with the position `localSheetId` counts by. */
interface SheetDeclaration {
  index: number;
  name: string;
  state: string | undefined;
  relationshipId: string;
  partPath: string | undefined;
}

const WORKSHEET_RELATIONSHIP_SUFFIX = "/worksheet";

function readSheetDeclarations(
  workbookPackage: WorkbookPackage,
): SheetDeclaration[] {
  const xml = workbookPackage.readText(WORKBOOK_PART);
  if (xml === undefined) {
    return [];
  }
  const sheets = findElement(xml, "sheets");
  if (!sheets) {
    return [];
  }
  const relationships = new Map(
    workbookPackage
      .relationshipsOf(WORKBOOK_PART)
      .map((relationship) => [relationship.id, relationship] as const),
  );
  const inner = xml.slice(sheets.innerStart, sheets.innerEnd);

  return findElements(inner, "sheet").map((element, index) => {
    const openTag = inner.slice(element.start, element.end);
    const relationshipId = getAttribute(openTag, "id") ?? "";
    const relationship = relationships.get(relationshipId);
    const isWorksheet =
      relationship?.type.endsWith(WORKSHEET_RELATIONSHIP_SUFFIX) === true;
    return {
      index,
      name: decodeXmlText(getAttribute(openTag, "name") ?? ""),
      state: getAttribute(openTag, "state"),
      relationshipId,
      partPath:
        isWorksheet && relationship
          ? workbookPackage.resolvePart(WORKBOOK_PART, relationship.target)
          : undefined,
    };
  });
}

function visibilityOf(state: string | undefined): 0 | 1 | 2 {
  if (state === "veryHidden") {
    return 2;
  }
  return state === "hidden" ? 1 : 0;
}

function visibilityLabel(visibility: number): string {
  if (visibility === 2) {
    return "Very hidden";
  }
  return visibility === 1 ? "Hidden" : "Visible";
}

/** The workbook-unique name a source sheet takes in the merged workbook. */
function claimSheetName(state: MergeWorkbooksState, original: string): string {
  const base = original.slice(0, MAXIMUM_SHEET_NAME_LENGTH) || "Sheet";
  let final = base;
  let suffix = 2;
  while (state.usedSheetNames.has(final.toLocaleLowerCase())) {
    const suffixText = ` (${suffix})`;
    final = `${base.slice(
      0,
      MAXIMUM_SHEET_NAME_LENGTH - suffixText.length,
    )}${suffixText}`;
    suffix += 1;
  }
  state.usedSheetNames.add(final.toLocaleLowerCase());
  return final;
}

/** The workbook-unique name a source Excel Table takes. */
function claimTableName(state: MergeWorkbooksState, original: string): string {
  let final = original;
  let suffix = 2;
  while (state.usedTableNames.has(final.toLocaleLowerCase())) {
    final = `${original}${suffix}`;
    suffix += 1;
  }
  state.usedTableNames.add(final.toLocaleLowerCase());
  if (final !== original) {
    state.renamedTables.push([original, final]);
  }
  return final;
}

// -- Removing what a merged workbook cannot carry --------------------------

/** The part whose relationships a `.rels` part stores. */
function relationshipsOwner(relationshipsPart: string): string {
  const directory = packagePartDirectory(relationshipsPart);
  const parent = packagePartDirectory(directory);
  const name = relationshipsPart.slice(relationshipsPart.lastIndexOf("/") + 1);
  const owner = name.replace(/\.rels$/u, "");
  if (owner === "") {
    return "";
  }
  return parent === "" ? owner : `${parent}/${owner}`;
}

/**
 * Remove parts and every reference to them: the relationship entries that
 * point at them, wherever they are declared, and their content-type
 * overrides. Returns the removed part paths.
 */
function dropParts(
  workbookPackage: WorkbookPackage,
  matches: (partPath: string) => boolean,
): string[] {
  const removed = workbookPackage.partNames().filter(matches);
  if (removed.length === 0) {
    return [];
  }
  const removedSet = new Set(removed);
  for (const partPath of removed) {
    workbookPackage.remove(partPath);
    workbookPackage.removeContentTypeOverride(partPath);
    workbookPackage.remove(relationshipsPartPath(partPath));
  }

  for (const relationshipsPart of workbookPackage
    .partNames()
    .filter((partPath) => partPath.endsWith(".rels"))) {
    const owner = relationshipsOwner(relationshipsPart);
    const xml = workbookPackage.readText(relationshipsPart);
    if (xml === undefined) {
      continue;
    }
    const rewritten = editElements(xml, "Relationship", (element, text) => {
      if (getAttribute(element.openTag, "TargetMode") === "External") {
        return text;
      }
      const target = getAttribute(element.openTag, "Target");
      if (target === undefined) {
        return text;
      }
      try {
        return removedSet.has(workbookPackage.resolvePart(owner, target))
          ? undefined
          : text;
      } catch {
        return text;
      }
    });
    if (rewritten !== xml) {
      workbookPackage.writeText(relationshipsPart, rewritten);
    }
  }

  return removed;
}

/** Remove one element of the workbook part, if it declares one. */
function removeWorkbookElement(
  workbookPackage: WorkbookPackage,
  localName: string,
): void {
  const xml = workbookPackage.readText(WORKBOOK_PART);
  if (xml === undefined) {
    return;
  }
  const element = findElement(xml, localName);
  if (!element) {
    return;
  }
  workbookPackage.writeText(
    WORKBOOK_PART,
    `${xml.slice(0, element.start)}${xml.slice(element.end)}`,
  );
}

/** Strip the structures a merged workbook does not carry from one package. */
function stripUncarriedStructures(workbookPackage: WorkbookPackage): {
  pivotTables: number;
  externalLinks: number;
} {
  const pivotTables = dropParts(workbookPackage, (partPath) =>
    /^xl\/pivotTables\/[^/]+\.xml$/iu.test(partPath),
  ).length;
  const externalLinks = dropParts(workbookPackage, (partPath) =>
    /^xl\/externalLinks\/externalLink[^/]*\.xml$/iu.test(partPath),
  ).length;
  dropParts(workbookPackage, (partPath) =>
    DROPPED_PART_PREFIXES.some((prefix) =>
      partPath.toLocaleLowerCase().startsWith(prefix),
    ),
  );
  removeWorkbookElement(workbookPackage, "pivotCaches");
  removeWorkbookElement(workbookPackage, "externalReferences");
  workbookPackage.removePartAndReferences(CALC_CHAIN_PART, WORKBOOK_PART);

  return { pivotTables, externalLinks };
}

// -- Transplanting one part ------------------------------------------------

interface TransplantContext {
  readonly source: WorkbookPackage;
  readonly output: WorkbookPackage;
  readonly rewrites: NameRewrites;
  /** Source part path to the path it was copied to, so a shared part copies once. */
  readonly copied: Map<string, string>;
  /** Per-part rewriting for the kinds that carry indexes of their own. */
  readonly rewritePart: (sourcePart: string, xml: string) => string;
}

function isDroppedPart(partPath: string): boolean {
  const normalized = partPath.toLocaleLowerCase();
  return (
    DROPPED_PART_PREFIXES.some((prefix) => normalized.startsWith(prefix)) ||
    normalized === VBA_PROJECT_PART.toLocaleLowerCase() ||
    normalized === CALC_CHAIN_PART.toLocaleLowerCase()
  );
}

function isXmlPart(partPath: string): boolean {
  return partExtension(partPath) === "xml";
}

/** Copy a source part's relationships, copying the parts they point at. */
function copyRelationships(
  context: TransplantContext,
  sourcePart: string,
  targetPart: string,
): Map<string, string> {
  const idMap = new Map<string, string>();
  const carried: PackageRelationship[] = [];

  for (const relationship of context.source.relationshipsOf(sourcePart)) {
    const id = `rId${carried.length + 1}`;
    if (relationship.targetMode === "External") {
      carried.push({ ...relationship, id });
      idMap.set(relationship.id, id);
      continue;
    }
    let resolved: string;
    try {
      resolved = context.source.resolvePart(sourcePart, relationship.target);
    } catch {
      continue;
    }
    const copied = copyPart(context, resolved);
    if (copied === undefined) {
      continue;
    }
    carried.push({
      id,
      type: relationship.type,
      target: relativeRelationshipTarget(targetPart, copied),
    });
    idMap.set(relationship.id, id);
  }

  writeRelationships(context.output, targetPart, carried);
  return idMap;
}

/**
 * Copy one dependent part into the output package under a free path, together
 * with its content-type declaration and, recursively, the parts it depends on.
 */
function copyPart(
  context: TransplantContext,
  sourcePart: string,
): string | undefined {
  const already = context.copied.get(sourcePart);
  if (already !== undefined) {
    return already;
  }
  if (isDroppedPart(sourcePart)) {
    return undefined;
  }
  const bytes = context.source.readBytes(sourcePart);
  if (bytes === undefined) {
    return undefined;
  }

  const targetPart = allocatePartPath(
    context.output,
    packagePartDirectory(sourcePart),
    partStem(sourcePart),
    partExtension(sourcePart),
  );
  context.copied.set(sourcePart, targetPart);
  // Claim the path before recursing so a dependent cannot be handed the same one.
  context.output.writeBytes(targetPart, bytes);
  copyContentTypeDeclaration(
    context.source,
    sourcePart,
    context.output,
    targetPart,
  );

  const idMap = copyRelationships(context, sourcePart, targetPart);
  if (isXmlPart(sourcePart)) {
    const xml = context.source.readText(sourcePart) ?? "";
    context.output.writeText(
      targetPart,
      context.rewritePart(
        sourcePart,
        rewriteFormulaElements(
          rewriteRelationshipIds(xml, idMap),
          context.rewrites,
        ),
      ),
    );
  }
  return targetPart;
}

// -- Transplanting one workbook -------------------------------------------

/** Load one input as a package, reporting failure with the stable read error. */
async function loadInput(
  bytes: Uint8Array,
  sourceLabel: string,
): Promise<WorkbookPackage> {
  try {
    return await WorkbookPackage.load(bytes, { sourceLabel });
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${sourceLabel}`,
      { cause: error, details: { source: sourceLabel } },
    );
  }
}

function recordSheet(
  state: MergeWorkbooksState,
  sourceFile: string,
  originalName: string,
  finalName: string,
  visibility: 0 | 1 | 2,
): void {
  if (visibility !== 0) {
    state.hiddenSheets += 1;
  }
  state.outputSheets += 1;
  state.sheets.push({ name: finalName, visibility });
  state.indexRows.push([
    sourceFile,
    originalName,
    finalName,
    visibilityLabel(visibility),
  ]);
}

/** The renames one input's sheets and tables force on its formulas. */
function planRewrites(
  sheetRenames: ReadonlyMap<string, string>,
  tableRenames: ReadonlyMap<string, string>,
): NameRewrites {
  return { sheets: sheetRenames, tables: tableRenames };
}

/**
 * Fold one input workbook into the merged one. The first call seeds the output
 * package from that input, so the workbook the user listed first keeps its
 * theme, its styles and its untouched parts byte-for-byte.
 */
export async function appendWorkbookSheets(
  state: MergeWorkbooksState,
  sourceFile: string,
  workbookBytes: Uint8Array,
): Promise<void> {
  const source = await loadInput(workbookBytes, sourceFile);
  if (source.has(VBA_PROJECT_PART)) {
    state.macroInputs += 1;
  }

  if (!state.output) {
    seedOutput(state, source, sourceFile);
    return;
  }
  transplantWorkbook(state, source, sourceFile);
}

/** Adopt the first input as the output package and register what it claims. */
function seedOutput(
  state: MergeWorkbooksState,
  source: WorkbookPackage,
  sourceFile: string,
): void {
  const removed = stripUncarriedStructures(source);
  state.removedPivotTables += removed.pivotTables;
  state.removedExternalLinks += removed.externalLinks;
  state.output = source;
  state.seedHasMacroProject = source.has(VBA_PROJECT_PART);
  state.styles = new MergedStyles(source.readText(STYLES_PART));
  state.strings = SharedStringTable.from(source.readText(SHARED_STRINGS_PART));

  const declarations = readSheetDeclarations(source);
  const sheetRenames = new Map<string, string>();
  const outputIndexBySource = new Map<number, number>();

  for (const declaration of declarations) {
    if (declaration.partPath === undefined) {
      continue;
    }
    const finalName = claimSheetName(state, declaration.name);
    if (finalName !== declaration.name) {
      sheetRenames.set(declaration.name.toLocaleLowerCase(), finalName);
      renameSheetDeclaration(source, declaration.name, finalName);
    }
    outputIndexBySource.set(declaration.index, state.sheets.length);
    recordSheet(
      state,
      sourceFile,
      declaration.name,
      finalName,
      visibilityOf(declaration.state),
    );
  }
  state.nextSheetId = nextSheetId(source);

  const tableRenames = new Map<string, string>();
  for (const table of readExcelTableDefinitionsFrom(source)) {
    // Nothing is claimed yet, so the seed's own tables always keep their names.
    claimTableName(state, table.name);
    state.nextTableId = Math.max(
      state.nextTableId,
      tableIdOf(source, table.tablePart) + 1,
    );
  }

  const rewrites = planRewrites(sheetRenames, tableRenames);
  carryDefinedNames(state, source, rewrites, outputIndexBySource);
  if (sheetRenames.size > 0) {
    rewriteSeedFormulas(source, declarations, rewrites);
  }
}

/** The next free `sheetId` in a workbook part. */
function nextSheetId(workbookPackage: WorkbookPackage): number {
  const xml = workbookPackage.readText(WORKBOOK_PART) ?? "";
  const sheets = findElement(xml, "sheets");
  if (!sheets) {
    return 1;
  }
  const inner = xml.slice(sheets.innerStart, sheets.innerEnd);
  let highest = 0;
  for (const element of findElements(inner, "sheet")) {
    const id = Number(
      getAttribute(inner.slice(element.start, element.end), "sheetId") ?? "0",
    );
    if (Number.isInteger(id)) {
      highest = Math.max(highest, id);
    }
  }
  return highest + 1;
}

/** Rename one `<sheet>` declaration in a workbook part. */
function renameSheetDeclaration(
  workbookPackage: WorkbookPackage,
  original: string,
  final: string,
): void {
  const xml = workbookPackage.readText(WORKBOOK_PART);
  if (xml === undefined) {
    return;
  }
  const rewritten = editElements(xml, "sheet", (element, text) => {
    if (
      decodeXmlText(getAttribute(element.openTag, "name") ?? "") !== original
    ) {
      return text;
    }
    return `${setAttribute(
      element.openTag,
      "name",
      escapeXmlAttribute(final),
    )}${text.slice(element.openTag.length)}`;
  });
  if (rewritten !== xml) {
    workbookPackage.writeText(WORKBOOK_PART, rewritten);
  }
}

/** Follow a seed rename into the seed's own parts that name the sheet. */
function rewriteSeedFormulas(
  source: WorkbookPackage,
  declarations: readonly SheetDeclaration[],
  rewrites: NameRewrites,
): void {
  const parts = new Set<string>();
  for (const declaration of declarations) {
    if (declaration.partPath !== undefined) {
      parts.add(declaration.partPath);
    }
  }
  for (const table of readExcelTableDefinitionsFrom(source)) {
    parts.add(table.tablePart);
  }
  for (const partPath of parts) {
    const xml = source.readText(partPath);
    if (xml === undefined) {
      continue;
    }
    const rewritten = rewriteFormulaElements(xml, rewrites);
    if (rewritten !== xml) {
      source.writeText(partPath, rewritten);
    }
  }
}

/** Copy every worksheet of a later input into the output package. */
function transplantWorkbook(
  state: MergeWorkbooksState,
  source: WorkbookPackage,
  sourceFile: string,
): void {
  const output = state.output;
  const styles = state.styles;
  const strings = state.strings;
  if (!output || !styles || !strings) {
    return;
  }

  const declarations = readSheetDeclarations(source);
  const sheetRenames = new Map<string, string>();
  const finalNames = new Map<number, string>();
  const outputIndexBySource = new Map<number, number>();

  for (const declaration of declarations) {
    if (declaration.partPath === undefined) {
      continue;
    }
    const finalName = claimSheetName(state, declaration.name);
    finalNames.set(declaration.index, finalName);
    if (finalName !== declaration.name) {
      sheetRenames.set(declaration.name.toLocaleLowerCase(), finalName);
    }
    outputIndexBySource.set(declaration.index, state.sheets.length);
    recordSheet(
      state,
      sourceFile,
      declaration.name,
      finalName,
      visibilityOf(declaration.state),
    );
  }

  state.removedPivotTables += source
    .partNames()
    .filter((partPath) =>
      /^xl\/pivotTables\/[^/]+\.xml$/iu.test(partPath),
    ).length;
  state.removedExternalLinks += source
    .partNames()
    .filter((partPath) =>
      /^xl\/externalLinks\/externalLink[^/]*\.xml$/iu.test(partPath),
    ).length;

  const tableDefinitions = readExcelTableDefinitionsFrom(source);
  const tableRenames = new Map<string, string>();
  const renamedTableByPart = new Map<string, string>();
  for (const table of tableDefinitions) {
    const finalName = claimTableName(state, table.name);
    renamedTableByPart.set(table.tablePart, finalName);
    if (finalName !== table.name) {
      tableRenames.set(table.name.toLocaleLowerCase(), finalName);
    }
  }

  const rewrites = planRewrites(sheetRenames, tableRenames);
  const sourceStyles = new SourceStyles(source.readText(STYLES_PART));
  const stringRemap = strings.absorb(source.readText(SHARED_STRINGS_PART));
  const styleRemap = new Map<number, number>();
  const dxfRemap = new Map<number, number>();

  const remaps = {
    style: (index: number): number => {
      const cached = styleRemap.get(index);
      if (cached !== undefined) {
        return cached;
      }
      const mapped = styles.remapCellXf(sourceStyles, index);
      styleRemap.set(index, mapped);
      return mapped;
    },
    sharedString: (index: number): number => stringRemap[index] ?? index,
    differentialFormat: (index: number): number => {
      const cached = dxfRemap.get(index);
      if (cached !== undefined) {
        return cached;
      }
      const mapped = styles.remapDxf(sourceStyles, index);
      dxfRemap.set(index, mapped);
      return mapped;
    },
  };

  const context: TransplantContext = {
    source,
    output,
    rewrites,
    copied: new Map<string, string>(),
    rewritePart: (sourcePart, xml) => {
      const renamed = renamedTableByPart.get(sourcePart);
      if (renamed === undefined) {
        return xml;
      }
      const id = state.nextTableId;
      state.nextTableId += 1;
      return renameTablePart(
        rewriteDifferentialFormatIds(xml, remaps.differentialFormat),
        renamed,
        id,
      );
    },
  };

  for (const declaration of declarations) {
    if (declaration.partPath === undefined) {
      continue;
    }
    const finalName = finalNames.get(declaration.index) ?? declaration.name;
    const targetPart = transplantWorksheet(
      context,
      declaration.partPath,
      remaps,
    );
    const relationshipId = addRelationship(output, WORKBOOK_PART, {
      type: `${RELATIONSHIP_NAMESPACE}/worksheet`,
      target: relativeRelationshipTarget(WORKBOOK_PART, targetPart),
    });
    appendSheetDeclaration(output, {
      name: finalName,
      sheetId: state.nextSheetId,
      state: declaration.state,
      relationshipId,
    });
    state.nextSheetId += 1;
  }

  carryDefinedNames(state, source, rewrites, outputIndexBySource);
}

/** Copy one worksheet part with its dependents and re-index it. */
function transplantWorksheet(
  context: TransplantContext,
  sourcePart: string,
  remaps: IndexRemaps,
): string {
  const targetPart = allocatePartPath(
    context.output,
    "xl/worksheets",
    "sheet",
    "xml",
  );
  context.copied.set(sourcePart, targetPart);
  context.output.writeText(targetPart, "");
  copyContentTypeDeclaration(
    context.source,
    sourcePart,
    context.output,
    targetPart,
  );
  addContentTypeOverride(context.output, targetPart, WORKSHEET_CONTENT_TYPE);

  const idMap = copyRelationships(context, sourcePart, targetPart);
  const xml = context.source.requireText(sourcePart);
  context.output.writeText(
    targetPart,
    rewriteRelationshipIds(
      rewriteFormulaElements(
        rewriteWorksheetIndexes(xml, remaps),
        context.rewrites,
      ),
      idMap,
    ),
  );
  return targetPart;
}

/** Give a copied Excel Table part its new workbook-unique name and id. */
function renameTablePart(xml: string, name: string, id: number): string {
  return editElements(xml, "table", (element, text) => {
    let openTag = setAttribute(element.openTag, "id", String(id));
    openTag = setAttribute(openTag, "name", escapeXmlAttribute(name));
    openTag = setAttribute(openTag, "displayName", escapeXmlAttribute(name));
    return `${openTag}${text.slice(element.openTag.length)}`;
  });
}

/** Append a `<sheet>` declaration to a workbook part. */
function appendSheetDeclaration(
  workbookPackage: WorkbookPackage,
  sheet: {
    name: string;
    sheetId: number;
    state: string | undefined;
    relationshipId: string;
  },
): void {
  const xml = workbookPackage.readText(WORKBOOK_PART);
  if (xml === undefined) {
    return;
  }
  const sheets = findElement(xml, "sheets");
  if (!sheets) {
    return;
  }
  const declaration = `<sheet name="${escapeXmlAttribute(sheet.name)}" sheetId="${
    sheet.sheetId
  }"${
    sheet.state === undefined
      ? ""
      : ` state="${escapeXmlAttribute(sheet.state)}"`
  } r:id="${sheet.relationshipId}"/>`;
  workbookPackage.writeText(
    WORKBOOK_PART,
    `${xml.slice(0, sheets.innerEnd)}${declaration}${xml.slice(sheets.innerEnd)}`,
  );
}

// -- Defined names ---------------------------------------------------------

/**
 * Carry one input's defined names.
 *
 * Workbook scope is first-come: the first input to claim a name keeps it, and
 * a later one is renamed with a numeric suffix so its formulas can still be
 * pointed at it. Sheet scope follows its sheet, so it never collides - two
 * inputs' `Print_Area` names are scoped to different sheets by construction.
 */
function carryDefinedNames(
  state: MergeWorkbooksState,
  source: WorkbookPackage,
  rewrites: NameRewrites,
  outputIndexBySource: ReadonlyMap<number, number>,
): void {
  const xml = source.readText(WORKBOOK_PART);
  if (xml === undefined) {
    return;
  }
  const container = findElement(xml, "definedNames");
  if (!container || container.selfClosing) {
    return;
  }
  const inner = xml.slice(container.innerStart, container.innerEnd);

  for (const element of findElements(inner, "definedName")) {
    const original = decodeXmlText(getAttribute(element.openTag, "name") ?? "");
    if (original === "") {
      continue;
    }
    const declared = getAttribute(element.openTag, "localSheetId");
    const reference = element.selfClosing
      ? ""
      : rewriteReference(
          decodeXmlText(inner.slice(element.innerStart, element.innerEnd)),
          rewrites,
        );

    if (declared === undefined) {
      let name = original;
      let suffix = 2;
      while (state.usedDefinedNames.has(name.toLocaleLowerCase())) {
        name = `${original}_${suffix}`;
        suffix += 1;
      }
      state.usedDefinedNames.add(name.toLocaleLowerCase());
      if (name !== original) {
        state.renamedDefinedNames.push([original, name]);
      }
      state.definedNames.push({ name, localSheetId: undefined, reference });
      continue;
    }

    const localSheetId = outputIndexBySource.get(Number(declared));
    if (localSheetId === undefined) {
      continue;
    }
    state.definedNames.push({ name: original, localSheetId, reference });
  }
}

function rewriteReference(reference: string, rewrites: NameRewrites): string {
  return rewriteFormulaElements(`<f>${escapeXmlText(reference)}</f>`, rewrites)
    .replace(/^<f>/u, "")
    .replace(/<\/f>$/u, "");
}

/** Write the merged defined-name list into the output workbook part. */
function writeDefinedNames(state: MergeWorkbooksState): void {
  const output = state.output;
  if (!output) {
    return;
  }
  const xml = output.readText(WORKBOOK_PART);
  if (xml === undefined) {
    return;
  }
  const block =
    state.definedNames.length === 0
      ? ""
      : `<definedNames>${state.definedNames
          .map(
            (entry) =>
              `<definedName name="${escapeXmlAttribute(entry.name)}"${
                entry.localSheetId === undefined
                  ? ""
                  : ` localSheetId="${entry.localSheetId}"`
              }>${escapeXmlText(entry.reference)}</definedName>`,
          )
          .join("")}</definedNames>`;

  const existing = findElement(xml, "definedNames");
  if (existing) {
    output.writeText(
      WORKBOOK_PART,
      `${xml.slice(0, existing.start)}${block}${xml.slice(existing.end)}`,
    );
    return;
  }
  if (block === "") {
    return;
  }
  const sheets = findElement(xml, "sheets");
  const at = sheets ? sheets.end : xml.lastIndexOf("</workbook>");
  if (at < 0) {
    return;
  }
  output.writeText(
    WORKBOOK_PART,
    `${xml.slice(0, at)}${block}${xml.slice(at)}`,
  );
}

/**
 * Ask Excel to recalculate on open. The merged workbook has no calculation
 * chain - every transplanted sheet took a new sheet id, which is what a chain
 * entry is keyed by - so the workbook says so rather than leaving a reader to
 * trust cached values it cannot verify.
 */
function requestFullCalculation(output: WorkbookPackage): void {
  const xml = output.readText(WORKBOOK_PART);
  if (xml === undefined) {
    return;
  }
  const existing = findElement(xml, "calcPr");
  if (existing) {
    const openTag =
      getAttribute(existing.openTag, "fullCalcOnLoad") === undefined
        ? addAttribute(existing.openTag, "fullCalcOnLoad", "1")
        : setAttribute(existing.openTag, "fullCalcOnLoad", "1");
    output.writeText(
      WORKBOOK_PART,
      `${xml.slice(0, existing.start)}${openTag}${xml.slice(
        existing.start + existing.openTag.length,
      )}`,
    );
    return;
  }
  const definedNames = findElement(xml, "definedNames");
  const sheets = findElement(xml, "sheets");
  const at = definedNames?.end ?? sheets?.end ?? -1;
  if (at < 0) {
    return;
  }
  output.writeText(
    WORKBOOK_PART,
    `${xml.slice(0, at)}<calcPr calcId="0" fullCalcOnLoad="1"/>${xml.slice(at)}`,
  );
}

// -- The sheet index -------------------------------------------------------

/**
 * The provenance worksheet, written as inline strings so it never has to claim
 * a slot in the merged shared string table (and so its text is readable in the
 * part itself, which the corpus asserts).
 */
function buildSheetIndexWorksheet(rows: readonly string[][]): string {
  const columns = ["A", "B", "C", "D"] as const;
  const body = rows
    .map((row, rowIndex) => {
      const cells = row
        .map((value, columnIndex) => {
          const reference = `${columns[columnIndex] ?? "A"}${rowIndex + 1}`;
          return `<c r="${reference}" t="inlineStr"><is><t>${escapeXmlText(
            value,
          )}</t></is></c>`;
        })
        .join("");
      return `<row r="${rowIndex + 1}">${cells}</row>`;
    })
    .join("");

  return (
    `${XML_DECLARATION}<worksheet xmlns="${NAMESPACE_MAIN}">` +
    `<dimension ref="A1:D${Math.max(rows.length, 1)}"/>` +
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="15"/>' +
    '<cols><col min="1" max="4" width="28" customWidth="1"/></cols>' +
    `<sheetData>${body}</sheetData>` +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    "</worksheet>"
  );
}

function appendSheetIndex(state: MergeWorkbooksState): void {
  const output = state.output;
  if (!output) {
    return;
  }
  const partPath = allocatePartPath(output, "xl/worksheets", "sheet", "xml");
  output.writeText(partPath, buildSheetIndexWorksheet(state.indexRows));
  addContentTypeOverride(output, partPath, WORKSHEET_CONTENT_TYPE);
  const relationshipId = addRelationship(output, WORKBOOK_PART, {
    type: `${RELATIONSHIP_NAMESPACE}/worksheet`,
    target: relativeRelationshipTarget(WORKBOOK_PART, partPath),
  });
  appendSheetDeclaration(output, {
    name: SHEET_INDEX_NAME,
    sheetId: state.nextSheetId,
    state: undefined,
    relationshipId,
  });
  state.nextSheetId += 1;
}

// -- Finishing -------------------------------------------------------------

/**
 * Decide the macro project's fate.
 *
 * Two macro projects cannot be combined: `vbaProject.bin` is an opaque
 * compound file whose module names would collide. One can travel, but only
 * into an output that is named as a macro-enabled workbook - a package whose
 * content type says `.xlsm` while its name says `.xlsx` is one Excel opens
 * with a corruption warning, which is a worse outcome than a warned removal.
 */
function resolveMacroProject(state: MergeWorkbooksState): {
  enabled: boolean;
  warning: string | undefined;
} {
  const output = state.output;
  if (!output) {
    return { enabled: false, warning: undefined };
  }
  const keep =
    state.seedHasMacroProject &&
    state.macroInputs === 1 &&
    state.options.macroOutput === true;

  if (keep) {
    setContentTypeOverride(
      output,
      WORKBOOK_PART,
      MACRO_WORKBOOK_MAIN_CONTENT_TYPE,
    );
    return { enabled: true, warning: undefined };
  }

  if (state.macroInputs > 0) {
    dropParts(output, (partPath) => partPath === VBA_PROJECT_PART);
  }
  setContentTypeOverride(output, WORKBOOK_PART, WORKBOOK_MAIN_CONTENT_TYPE);

  if (state.macroInputs === 0) {
    return { enabled: false, warning: undefined };
  }
  if (state.macroInputs > 1) {
    return {
      enabled: false,
      warning: `Removed the macro project: ${state.macroInputs} inputs carried one and two VBA projects cannot be combined.`,
    };
  }
  return {
    enabled: false,
    warning: state.seedHasMacroProject
      ? "Removed the macro project: name the merged workbook with an .xlsm extension to keep it."
      : "Removed the macro project: only the first input's macros can travel into a merged workbook.",
  };
}

function pluralize(count: number, singular: string, plural: string): string {
  return count === 1 ? singular : plural;
}

/** Write the merged shared string and style parts, when a merge changed them. */
function flushWorkbookTables(state: MergeWorkbooksState): void {
  const output = state.output;
  const styles = state.styles;
  const strings = state.strings;
  if (!output || !styles || !strings) {
    return;
  }

  if (styles.changed) {
    output.writeText(STYLES_PART, styles.toXml());
  }
  if (
    strings.changed ||
    (strings.size > 0 && !output.has(SHARED_STRINGS_PART))
  ) {
    output.writeText(SHARED_STRINGS_PART, strings.toXml());
    addContentTypeOverride(
      output,
      SHARED_STRINGS_PART,
      SHARED_STRINGS_CONTENT_TYPE,
    );
    if (
      !output
        .relationshipsOf(WORKBOOK_PART)
        .some((relationship) => relationship.type.endsWith("/sharedStrings"))
    ) {
      addRelationship(output, WORKBOOK_PART, {
        type: `${RELATIONSHIP_NAMESPACE}/sharedStrings`,
        target: "sharedStrings.xml",
      });
    }
  }
}

export async function finishMergedWorkbook(
  state: MergeWorkbooksState,
  options: MergeWorkbooksBuildOptions,
): Promise<MergedWorkbook> {
  const output = state.output;
  if (!output || state.outputSheets === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_SHEETS,
      "No worksheets were found in the input workbooks.",
    );
  }

  flushWorkbookTables(state);
  writeDefinedNames(state);
  requestFullCalculation(output);
  const macro = resolveMacroProject(state);
  if (options.includeSheetIndex !== false) {
    appendSheetIndex(state);
  }

  let bytes = await output.save();
  if (options.values) {
    bytes = await convertWorkbookToValues(bytes);
  }

  const warnings: string[] = [];
  if (state.hiddenSheets > 0) {
    const plural = pluralize(state.hiddenSheets, " was", "s were");
    warnings.push(
      options.includeSheetIndex === false
        ? `${state.hiddenSheets} source worksheet${plural} hidden in the merged workbook.`
        : `${state.hiddenSheets} source worksheet${plural} hidden; see the visible "${SHEET_INDEX_NAME}" worksheet.`,
    );
  }
  if (state.removedPivotTables > 0) {
    warnings.push(
      `Removed ${state.removedPivotTables} pivot table${pluralize(
        state.removedPivotTables,
        "",
        "s",
      )}: a pivot cache is a private copy of its source rows and its registry is workbook-scoped, so it cannot be carried across a merge. Rebuild the pivot in Excel from the merged worksheets if it is required.`,
    );
  }
  if (state.removedExternalLinks > 0) {
    warnings.push(
      `Removed ${state.removedExternalLinks} external link${pluralize(
        state.removedExternalLinks,
        "",
        "s",
      )}: external references are addressed by their position in a single workbook's list, which a merge cannot renumber. Re-create the link in Excel if it is required.`,
    );
  }
  if (macro.warning !== undefined) {
    warnings.push(macro.warning);
  }
  if (state.renamedTables.length > 0) {
    warnings.push(
      `Renamed ${state.renamedTables.length} Excel Table${pluralize(
        state.renamedTables.length,
        "",
        "s",
      )} to keep table names unique in the merged workbook: ${state.renamedTables
        .map(([from, to]) => `"${from}" to "${to}"`)
        .join(", ")}.`,
    );
  }
  if (state.renamedDefinedNames.length > 0) {
    warnings.push(
      `Renamed ${state.renamedDefinedNames.length} defined name${pluralize(
        state.renamedDefinedNames.length,
        "",
        "s",
      )} an earlier workbook already claimed: ${state.renamedDefinedNames
        .map(([from, to]) => `"${from}" to "${to}"`)
        .join(", ")}.`,
    );
  }

  return {
    bytes,
    hiddenSheets: state.hiddenSheets,
    outputSheets: state.outputSheets,
    warnings,
    macroEnabled: macro.enabled,
  };
}

/** The `id` an Excel Table part declares, or 0 when it declares none. */
function tableIdOf(
  workbookPackage: WorkbookPackage,
  tablePart: string,
): number {
  const xml = workbookPackage.readText(tablePart);
  const element = xml === undefined ? undefined : findElement(xml, "table");
  const id = Number(
    element === undefined ? "0" : (getAttribute(element.openTag, "id") ?? "0"),
  );
  return Number.isInteger(id) ? id : 0;
}
