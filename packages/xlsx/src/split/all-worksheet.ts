/**
 * L3 - the all-worksheet split, as an operation both surfaces can run.
 *
 * The engine used to live inside the file surface, tangled with `node:fs`, so
 * the byte surface could not reach it and offered a weaker split instead. What
 * moved here is everything that is true of the operation regardless of where
 * its bytes come from or go: finding the regions, grouping the keys, filtering
 * each output, and the metrics and warnings that describe what happened.
 *
 * What deliberately did NOT move is the part each surface owns. A filesystem
 * split resolves paths, refuses unsafe destinations and commits through a
 * staging directory; a byte split names downloads and returns buffers. They
 * meet at `runAllWorksheetSplit`, which builds one output at a time and hands
 * it to the caller's `write`.
 *
 * Behaviour is deliberately unchanged from the file-only engine, quirks
 * included: a hidden worksheet carrying the split column is filtered like any
 * other, and the range binding treats every row below the header as data (so
 * it removes a totals row and a footer block that the table binding keeps).
 */
import {
  ConsultChimpsError,
  throwIfAborted,
  type AbortOutputContext,
} from "@consultchimps/core";

import { XLSX_ERRORS } from "../errors.js";
import { WorkbookModel } from "../model/index.js";
import {
  MACRO_WORKBOOK_MAIN_CONTENT_TYPE,
  WORKBOOK_MAIN_PART,
  WorkbookPackage,
} from "../package/index.js";
import type { RowNumber } from "../model/types.js";
import { isTableEditReport } from "../region/table-binding.js";
import type { DataRegion } from "../region/types.js";
import { resolveRegions } from "../region/resolve.js";
import {
  normalizeHeader,
  readRowValues,
  type NormalizedValue,
} from "../region/values.js";
import { stripPivotParts } from "../tier1/pivot.js";
import { blankStaleCachedFormulas } from "../tier1/stale-values.js";
import { convertWorkbookToValuesWithReport } from "../values-only.js";

export const SPLIT_OPERATION = "sheets.split-by-column";
const XLSX_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const XLSM_MEDIA_TYPE = "application/vnd.ms-excel.sheet.macroEnabled.12";

export type WorkbookExtension = ".xlsm" | ".xlsx";

export type AllWorksheetSplitMetric =
  | "calcChainEntriesRemoved"
  | "formulaCellsBlankedForRemovedRows"
  | "formulaCellsConverted"
  | "formulaCellsWithoutCachedValues"
  | "groups"
  | "pivotTablesRemoved"
  | "inputFiles"
  | "inputRows"
  | "outputFiles"
  | "outputRows"
  | "rowsDeleted"
  | "sheetsCopiedUnchanged"
  | "sheetsFiltered"
  | "skippedRows"
  | "valuesOnly";

export interface SplitSheetDetail {
  deletedRows: number;
  retainedRows: number;
  sheet: string;
}

export interface SplitOutputDetail {
  formulaCellsConverted: number;
  formulaCellsWithoutCachedValues: number;
  /** The output's name on the surface that produced it: a path, or a filename. */
  output: string;
  sheets: SplitSheetDetail[];
  value: string;
}

/**
 * The input, options and worksheet roles behind a completed split, in the terms
 * a person reading the result would use. The filesystem surface adds the output
 * directory it wrote into; a byte split has no directory to name.
 */
export interface AllWorksheetSplitSummary {
  column: string;
  copiedUnchangedSheets: string[];
  filteredSheets: string[];
  input: string;
  valuesOnly: boolean;
}

/** The split's own options, with no surface-specific input or output naming. */
export interface AllWorksheetSplitSelection {
  column: string;
  headerRow?: number | undefined;
  /** Compare split values without trimming, case folding, or numeric coercion. */
  strict?: boolean | undefined;
  values?: boolean | undefined;
}

/**
 * How a surface names the workbook it is splitting.
 *
 * Every error the operation raises quotes `label` and carries `details`, so a
 * path-based failure still reads as a path and an in-memory one still reads as
 * the upload's name. The operation itself never inspects either.
 */
export interface SplitSourceIdentity {
  details: Record<string, unknown>;
  label: string;
}

export type SplitGroup = NormalizedValue;

/** One filtered worksheet, as the analysis pass understood it. */
interface SheetAnalysis {
  /** True when the region is an Excel Table rather than a worksheet range. */
  isTable: boolean;
  name: string;
  /** Normalized key per body row; every body row has an entry. */
  rowValues: Map<RowNumber, NormalizedValue | undefined>;
  worksheetPart: string;
}

/**
 * Everything a split learned by reading the source once. Both the plan and the
 * run are computed from this, so a preview and the split it previews cannot
 * disagree.
 */
export interface AllWorksheetSplitAnalysis {
  extension: WorkbookExtension;
  groups: SplitGroup[];
  inputRows: number;
  mediaType: string;
  sheets: SheetAnalysis[];
  skippedRows: number;
  unchangedSheets: string[];
  workbookBytes: Uint8Array;
  /** Sheet name by worksheet part, for every sheet, filtered or not. */
  worksheetNames: ReadonlyMap<string, string>;
}

function isColumnNotFound(error: unknown): boolean {
  return (
    error instanceof ConsultChimpsError &&
    error.code === XLSX_ERRORS.XLSX_SPLIT_COLUMN_NOT_FOUND
  );
}

/**
 * The workbook type a name declares, refusing anything the split cannot open.
 *
 * The check is on the name rather than the bytes because the extension also
 * decides what the outputs are called: a macro workbook has to stay `.xlsm`,
 * or Excel opens the result with a corruption warning.
 */
export function workbookExtensionOf(
  name: string,
  identity: SplitSourceIdentity,
): WorkbookExtension {
  const match = /\.(xls[xm])$/iu.exec(name);
  const extension = match ? `.${match[1]!.toLowerCase()}` : "";
  if (extension !== ".xlsx" && extension !== ".xlsm") {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_UNSUPPORTED_FILE,
      `Unsupported Excel workbook type "${extension || "(none)"}". Choose an .xlsx or .xlsm workbook.`,
      { details: { ...identity.details } },
    );
  }
  return extension;
}

export function splitMediaType(extension: WorkbookExtension): string {
  return extension === ".xlsm" ? XLSM_MEDIA_TYPE : XLSX_MEDIA_TYPE;
}

/**
 * The region's column carrying the split key. The resolver already proved the
 * header text is there; this finds it again by name so the operation never has
 * to know whether the names came from a table part or from header cells.
 */
function splitColumnOf(
  region: DataRegion,
  columnText: string,
): { index: number } | undefined {
  const target = normalizeHeader(columnText);
  return region.columns.find(
    (candidate) => normalizeHeader(candidate.name) === target,
  );
}

/**
 * Every worksheet that carries the split column, as a region, plus the names
 * of the worksheets that do not and are therefore copied through untouched.
 *
 * The all-worksheet selector has no `headerRow`, so each sheet is resolved on
 * its own with the `{ sheet, headerRow }` selector and a missing column is
 * read as "this sheet is not part of the split" rather than as a failure. That
 * is exactly what the previous engine's per-sheet header search did.
 */
async function resolveSplitRegions(
  workbook: WorkbookModel,
  columnText: string,
  headerRow: number | undefined,
): Promise<{
  regions: readonly DataRegion[];
  unchangedSheets: string[];
}> {
  const regions: DataRegion[] = [];
  const unchangedSheets: string[] = [];

  for (const sheet of workbook.sheets) {
    let resolved: readonly DataRegion[];
    try {
      resolved = await resolveRegions(
        workbook,
        { headerRow, sheet: sheet.name },
        columnText,
      );
    } catch (error) {
      if (isColumnNotFound(error)) {
        unchangedSheets.push(sheet.name);
        continue;
      }
      throw error;
    }
    const region = resolved[0];
    if (!region || !splitColumnOf(region, columnText)) {
      unchangedSheets.push(sheet.name);
      continue;
    }
    regions.push(region);
  }

  return { regions, unchangedSheets };
}

/** The source rows one group's output drops from one analysed worksheet. */
function rowsRemovedFromSheet(
  sheet: SheetAnalysis,
  groupKey: string,
): Set<RowNumber> {
  const removed = new Set<RowNumber>();
  for (const [row, value] of sheet.rowValues) {
    if (value?.key !== groupKey) {
      removed.add(row);
    }
  }
  return removed;
}

async function loadWorkbookModel(
  workbookBytes: Uint8Array,
  identity: SplitSourceIdentity,
): Promise<WorkbookModel> {
  try {
    return await WorkbookModel.load(workbookBytes);
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not inspect workbook structure: ${identity.label}. The file may be corrupted, encrypted, or contain invalid workbook XML.`,
      { cause: error, details: { ...identity.details } },
    );
  }
}

/**
 * Refuse a workbook whose package contradicts the name it arrived under.
 *
 * The split preserves the source package, so the outputs inherit whatever the
 * source declares while taking their extension and media type from the input's
 * name. When the two disagree, every output would be mislabelled: `.xlsm` bytes
 * named `.xlsx` would carry a macro project the name denies, and `.xlsx` bytes
 * named `.xlsm` would be advertised as macro-enabled while the package says
 * otherwise - a contradiction Excel opens with a corruption warning.
 *
 * Neither reading can be repaired without deciding something the caller did not
 * ask for: stripping macros loses work, rewriting the declared type changes what
 * the file is. So this refuses before an output exists, and says which side to
 * correct.
 */
function refuseMislabelledPackage(
  workbook: WorkbookModel,
  extension: WorkbookExtension,
  identity: SplitSourceIdentity,
): void {
  refuseMislabelledWorkbookType(workbook.macroEnabled, extension, identity);
}

/**
 * The refusal itself, over the one fact it needs.
 *
 * Taking `macroEnabled` rather than a whole model is what lets the preserved
 * Excel Table split reuse it: that mode preserves the source package too, but
 * it never loads the workbook model, so asking the package alone is both
 * cheaper and enough.
 */
export function refuseMislabelledWorkbookType(
  macroEnabled: boolean,
  extension: WorkbookExtension,
  identity: SplitSourceIdentity,
): void {
  const declaredExtension: WorkbookExtension = macroEnabled ? ".xlsm" : ".xlsx";
  if (declaredExtension === extension) {
    return;
  }
  throw new ConsultChimpsError(
    XLSX_ERRORS.XLSX_SPLIT_PACKAGE_TYPE_MISMATCH,
    macroEnabled
      ? `The workbook "${identity.label}" is a macro-enabled workbook but is named "${extension}". Rename it with an .xlsm extension, or save it as an ordinary .xlsx workbook in Excel, and run the split again. Splitting it as it is would produce files whose contents and names disagree.`
      : `The workbook "${identity.label}" is named "${extension}" but is an ordinary Excel workbook with no macro project. Rename it with an .xlsx extension, or save it as a macro-enabled workbook in Excel, and run the split again. Splitting it as it is would produce files whose contents and names disagree.`,
    {
      details: {
        declaredExtension,
        macroEnabled,
        nameExtension: extension,
        ...identity.details,
      },
    },
  );
}

/**
 * The extension and media type a split that preserves the source package must
 * name its outputs with.
 *
 * A preserved split copies the source package into every output, so each one
 * inherits whatever that package declares while taking its name from the
 * source's name. That is the all-worksheet split's situation exactly, and it
 * gets the same answer and the same refusal: a macro-enabled package handed
 * back under an `.xlsx` name, or an ordinary one under `.xlsm`, is a file whose
 * contents and name disagree, which Excel opens with a corruption warning.
 *
 * A split that rebuilds instead writes a fresh ordinary package and is always
 * `.xlsx`, so this is asked only when the workbook is preserved.
 */
export async function preservedSplitExtension(
  workbookBytes: Uint8Array,
  name: string,
  identity: SplitSourceIdentity,
): Promise<WorkbookExtension> {
  const extension = workbookExtensionOf(name, identity);
  // The content-type lookup parses `[Content_Types].xml` on demand, so a
  // malformed or DOCTYPE-bearing declaration fails here rather than in the
  // load. Both belong inside the wrapper: a caller of this operation gets the
  // package's stable read failure, never a raw parser error.
  let declaresMacroWorkbook: boolean;
  try {
    const workbookPackage = await WorkbookPackage.load(workbookBytes);
    declaresMacroWorkbook =
      workbookPackage.contentTypeOverride(WORKBOOK_MAIN_PART)?.trim() ===
      MACRO_WORKBOOK_MAIN_CONTENT_TYPE;
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not inspect workbook structure: ${identity.label}. The file may be corrupted, encrypted, or contain invalid workbook XML.`,
      { cause: error, details: { ...identity.details } },
    );
  }
  refuseMislabelledWorkbookType(declaresMacroWorkbook, extension, identity);
  return extension;
}

/**
 * Read the source once: which worksheets carry the split column, which rows
 * belong to which group, and which worksheets are copied through untouched.
 */
export async function analyzeAllWorksheetSplit(
  workbookBytes: Uint8Array,
  extension: WorkbookExtension,
  selection: AllWorksheetSplitSelection,
  identity: SplitSourceIdentity,
): Promise<AllWorksheetSplitAnalysis> {
  if (selection.headerRow !== undefined && selection.headerRow < 1) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_HEADER_ROW,
      "The header row must be a positive whole number counted from 1.",
      { details: { headerRow: selection.headerRow } },
    );
  }

  const workbook = await loadWorkbookModel(workbookBytes, identity);
  refuseMislabelledPackage(workbook, extension, identity);
  const columnNotFound = (): ConsultChimpsError =>
    new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_COLUMN_NOT_FOUND,
      `Column "${selection.column}" was not found in any worksheet. Check the header text or provide the correct header row.`,
      {
        details: {
          availableWorksheets: workbook.sheets.map((sheet) => sheet.name),
          column: selection.column,
          headerRow: selection.headerRow,
          ...identity.details,
        },
      },
    );
  if (selection.column.trim() === "") {
    throw columnNotFound();
  }

  const { regions, unchangedSheets } = await resolveSplitRegions(
    workbook,
    selection.column,
    selection.headerRow,
  );
  if (regions.length === 0) {
    throw columnNotFound();
  }

  const matching = { strict: selection.strict === true };
  const groupsByKey = new Map<string, SplitGroup>();
  const sheets: SheetAnalysis[] = [];
  let inputRows = 0;
  let skippedRows = 0;

  for (const region of regions) {
    const column = splitColumnOf(region, selection.column)!;
    const rowValues = readRowValues(
      region.worksheet,
      region.body.start.row,
      region.body.end.row,
      column.index,
      matching,
    );
    for (const value of rowValues.values()) {
      inputRows += 1;
      if (!value) {
        skippedRows += 1;
        continue;
      }
      if (!groupsByKey.has(value.key)) {
        groupsByKey.set(value.key, value);
      }
    }
    sheets.push({
      isTable: region.origin.kind === "table",
      name: region.sheetName,
      rowValues,
      worksheetPart: region.worksheet.info.partPath,
    });
  }

  const groups = [...groupsByKey.values()];
  if (groups.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_NO_GROUPS,
      `Column "${selection.column}" does not contain any non-blank values. Add at least one value and try again.`,
      { details: { column: selection.column, ...identity.details } },
    );
  }

  return {
    extension,
    groups,
    inputRows,
    mediaType: splitMediaType(extension),
    sheets,
    skippedRows,
    unchangedSheets,
    workbookBytes,
    worksheetNames: new Map(
      workbook.sheets.map((sheet) => [sheet.partPath, sheet.name] as const),
    ),
  };
}

/** What building one group's workbook produced, beyond its bytes. */
interface GroupOutput {
  bytes: Uint8Array;
  calcChainEntriesRemoved: number;
  formulaCellsBlanked: number;
  formulaCellsConverted: number;
  formulaCellsWithoutCachedValues: number;
  pivotTablesRemoved: number;
  /** `Sheet!A1` of every cached result cleared as covering other groups. */
  staleAggregates: string[];
  /** Sheets whose Excel Table kept its original range, with the reason. */
  tableFallbackSheets: string[];
  /** `Sheet!A1` of every formula that lost its value in a values conversion. */
  uncachedFormulas: string[];
}

/**
 * Build one group's workbook: clear results computed over other groups' rows,
 * bake values when asked, filter every region, and remove the pivot caches
 * that would otherwise hand this recipient every other group's records.
 */
async function buildGroupWorkbook(
  analysis: AllWorksheetSplitAnalysis,
  group: SplitGroup,
  selection: AllWorksheetSplitSelection,
  identity: SplitSourceIdentity,
): Promise<GroupOutput> {
  const output: GroupOutput = {
    bytes: analysis.workbookBytes,
    calcChainEntriesRemoved: 0,
    formulaCellsBlanked: 0,
    formulaCellsConverted: 0,
    formulaCellsWithoutCachedValues: 0,
    pivotTablesRemoved: 0,
    staleAggregates: [],
    tableFallbackSheets: [],
    uncachedFormulas: [],
  };
  const worksheetNameByPart = analysis.worksheetNames;
  // A table region compacts its rows, so a values conversion has to run before
  // the filter for the cached results to line up with the rows they describe.
  // A pure worksheet split converts afterwards, over the rows that survive.
  const containsFilteredTable = analysis.sheets.some((sheet) => sheet.isTable);

  const convertToValues = async (): Promise<void> => {
    const conversion = await convertWorkbookToValuesWithReport(output.bytes);
    output.bytes = conversion.bytes;
    output.formulaCellsConverted = conversion.formulasConverted;
    output.formulaCellsWithoutCachedValues =
      conversion.formulasWithoutCachedValues.length;
    for (const missing of conversion.formulasWithoutCachedValues) {
      output.uncachedFormulas.push(
        `${worksheetNameByPart.get(missing.worksheetPart) ?? missing.worksheetPart}!${missing.cell}`,
      );
    }
  };

  if (selection.values) {
    // A values-only conversion bakes each formula's cached result into the
    // output, so any result computed over rows this group does not receive is
    // cleared first, while row numbers are still the source's.
    const staleValues = await blankStaleCachedFormulas(
      output.bytes,
      new Map(
        analysis.sheets.map(
          (sheet) =>
            [
              sheet.worksheetPart,
              rowsRemovedFromSheet(sheet, group.key),
            ] as const,
        ),
      ),
    );
    output.bytes = staleValues.bytes;
    output.formulaCellsBlanked = staleValues.blankedCells.length;
    for (const blanked of staleValues.blankedCells) {
      output.staleAggregates.push(`${blanked.sheet}!${blanked.cell}`);
    }
    if (containsFilteredTable) {
      await convertToValues();
    }
  }

  const workbook = await loadWorkbookModel(output.bytes, identity);
  const { regions } = await resolveSplitRegions(
    workbook,
    selection.column,
    selection.headerRow,
  );
  for (const region of regions) {
    const sheet = analysis.sheets.find(
      (candidate) => candidate.name === region.sheetName,
    );
    if (!sheet) {
      continue;
    }
    const report = region.filterRows(
      (row) => sheet.rowValues.get(row)?.key === group.key,
    );
    if (isTableEditReport(report) && !report.tableResized) {
      output.tableFallbackSheets.push(region.sheetName);
    }
  }
  output.bytes = await workbook.save();
  // The invariant pass maintained the chain as the rows moved; a values-only
  // output has no chain left to maintain, because the conversion removes it.
  output.calcChainEntriesRemoved = selection.values
    ? 0
    : workbook.calcChainEntriesRemoved;

  if (selection.values && !containsFilteredTable) {
    await convertToValues();
  }

  // A pivot cache is a private copy of every source row, so it would hand this
  // group's recipient every other group's data. It leaves with the rows it
  // cached, on every output.
  const strippedPivots = await stripPivotParts(output.bytes);
  output.bytes = strippedPivots.bytes;
  output.pivotTablesRemoved = strippedPivots.removedPivotTables;
  return output;
}

/** The metrics a plan reports, before a single output has been built. */
export function plannedAllWorksheetSplitMetrics(
  analysis: AllWorksheetSplitAnalysis,
  selection: AllWorksheetSplitSelection,
  outputFiles: number,
): Record<Exclude<AllWorksheetSplitMetric, "outputRows">, number> {
  return {
    calcChainEntriesRemoved: 0,
    formulaCellsBlankedForRemovedRows: 0,
    formulaCellsConverted: 0,
    formulaCellsWithoutCachedValues: 0,
    groups: analysis.groups.length,
    inputFiles: 1,
    inputRows: analysis.inputRows,
    outputFiles,
    pivotTablesRemoved: 0,
    rowsDeleted: 0,
    sheetsCopiedUnchanged: analysis.unchangedSheets.length,
    sheetsFiltered: analysis.sheets.length,
    skippedRows: analysis.skippedRows,
    valuesOnly: selection.values === true ? 1 : 0,
  };
}

export interface AllWorksheetSplitRun {
  metrics: Record<AllWorksheetSplitMetric, number>;
  outputs: SplitOutputDetail[];
  summary: AllWorksheetSplitSummary;
  warnings: string[];
}

export interface AllWorksheetSplitRunOptions {
  analysis: AllWorksheetSplitAnalysis;
  identity: SplitSourceIdentity;
  /** Where a cancellation leaves the outputs already produced. */
  outputContext: AbortOutputContext;
  /** One name per group, in group order: a path, or a portable filename. */
  outputNames: readonly string[];
  selection: AllWorksheetSplitSelection;
  signal?: AbortSignal | undefined;
  /**
   * Take delivery of one finished output. The surface decides what that means
   * - staging it in a transaction directory, or keeping the buffer - and
   * reports its own progress, because the two surfaces describe different work.
   */
  write: (
    index: number,
    bytes: Uint8Array,
    detail: SplitOutputDetail,
  ) => Promise<void> | void;
}

/**
 * Build every group's workbook in group order, handing each to `write`, and
 * report what the whole split did.
 *
 * Outputs are produced one at a time on purpose: a workbook per group, each a
 * copy of the source, is the largest thing this operation holds, and a surface
 * that can put one down (the filesystem's staging directory) should not be
 * forced to keep all of them.
 */
export async function runAllWorksheetSplit(
  options: AllWorksheetSplitRunOptions,
): Promise<AllWorksheetSplitRun> {
  const { analysis, identity, outputContext, outputNames, selection, signal } =
    options;
  const outputs: SplitOutputDetail[] = [];
  const missingFormulaLocations = new Set<string>();
  const staleAggregateLocations = new Set<string>();
  const tableFallbackSheets = new Set<string>();
  let calcChainEntriesRemoved = 0;
  let formulaCellsBlankedForRemovedRows = 0;
  let formulaCellsConverted = 0;
  let formulaCellsWithoutCachedValues = 0;
  let outputRows = 0;
  let pivotTablesRemoved = 0;
  let rowsDeleted = 0;

  for (const [index, group] of analysis.groups.entries()) {
    throwIfAborted(signal, SPLIT_OPERATION, outputContext);
    const built = await buildGroupWorkbook(
      analysis,
      group,
      selection,
      identity,
    );
    pivotTablesRemoved += built.pivotTablesRemoved;
    calcChainEntriesRemoved += built.calcChainEntriesRemoved;
    formulaCellsBlankedForRemovedRows += built.formulaCellsBlanked;
    formulaCellsConverted += built.formulaCellsConverted;
    formulaCellsWithoutCachedValues += built.formulaCellsWithoutCachedValues;
    built.staleAggregates.forEach((location) =>
      staleAggregateLocations.add(location),
    );
    built.uncachedFormulas.forEach((location) =>
      missingFormulaLocations.add(location),
    );
    built.tableFallbackSheets.forEach((sheet) =>
      tableFallbackSheets.add(sheet),
    );

    const sheets = analysis.sheets.map((sheet) => {
      const retainedRows = [...sheet.rowValues.values()].filter(
        (value) => value?.key === group.key,
      ).length;
      const deletedRows = sheet.rowValues.size - retainedRows;
      outputRows += retainedRows;
      rowsDeleted += deletedRows;
      return { deletedRows, retainedRows, sheet: sheet.name };
    });

    const detail: SplitOutputDetail = {
      formulaCellsConverted: built.formulaCellsConverted,
      formulaCellsWithoutCachedValues: built.formulaCellsWithoutCachedValues,
      output: outputNames[index]!,
      sheets,
      value: group.display,
    };
    await options.write(index, built.bytes, detail);
    outputs.push(detail);
  }

  // The last workbook was serialized asynchronously; honour a cancellation
  // that arrived while it was being built.
  throwIfAborted(signal, SPLIT_OPERATION, outputContext);

  const warnings: string[] = [];
  if (analysis.skippedRows > 0) {
    warnings.push(
      `Skipped ${analysis.skippedRows} row${analysis.skippedRows === 1 ? "" : "s"} with blank values in "${selection.column}"; no blank-value workbook was created.`,
    );
  }
  if (formulaCellsWithoutCachedValues > 0) {
    const locations = [...missingFormulaLocations];
    const shown = locations.slice(0, 20).join(", ");
    warnings.push(
      `${formulaCellsWithoutCachedValues} formula cell${formulaCellsWithoutCachedValues === 1 ? " had" : "s had"} no cached value and became formatted blank cells in values-only output. Affected source locations: ${shown}${locations.length > 20 ? `, and ${locations.length - 20} more` : ""}. Open and recalculate the source workbook in Excel, save it, and rerun the split if these values are required.`,
    );
  }
  if (pivotTablesRemoved > 0) {
    warnings.push(
      `Removed ${pivotTablesRemoved} pivot table${pivotTablesRemoved === 1 ? "" : "s"}: their caches contained rows from other groups, and a cache travels inside the workbook whether or not the pivot is opened. Rebuild the pivot in Excel from each output's own rows if it is required.`,
    );
  }
  if (formulaCellsBlankedForRemovedRows > 0) {
    const locations = [...staleAggregateLocations];
    const shown = locations.slice(0, 20).join(", ");
    warnings.push(
      `${formulaCellsBlankedForRemovedRows} cached formula result${formulaCellsBlankedForRemovedRows === 1 ? " covered rows" : "s covered rows"} that are not part of this group and ${formulaCellsBlankedForRemovedRows === 1 ? "was" : "were"} cleared, so the values-only output shows a blank cell instead of a total computed over every group's rows. Affected locations: ${shown}${locations.length > 20 ? `, and ${locations.length - 20} more` : ""}. Recalculate them in Excel against the delivered rows if these values are required.`,
    );
  }
  if (tableFallbackSheets.size > 0) {
    warnings.push(
      `Excel Table${tableFallbackSheets.size === 1 ? "" : "s"} on ${[...tableFallbackSheets].join(", ")} contained formulas tied to row positions, so unmatched rows were removed without compacting the table range. The table's formatting and formulas were preserved; review the range in Excel before delivery.`,
    );
  }

  return {
    metrics: {
      calcChainEntriesRemoved,
      formulaCellsBlankedForRemovedRows,
      formulaCellsConverted,
      formulaCellsWithoutCachedValues,
      groups: analysis.groups.length,
      inputFiles: 1,
      inputRows: analysis.inputRows,
      outputFiles: outputNames.length,
      outputRows,
      pivotTablesRemoved,
      rowsDeleted,
      sheetsCopiedUnchanged: analysis.unchangedSheets.length,
      sheetsFiltered: analysis.sheets.length,
      skippedRows: analysis.skippedRows,
      valuesOnly: selection.values === true ? 1 : 0,
    },
    outputs,
    summary: {
      column: selection.column,
      copiedUnchangedSheets: analysis.unchangedSheets,
      filteredSheets: analysis.sheets.map((sheet) => sheet.name),
      input: identity.label,
      valuesOnly: selection.values === true,
    },
    warnings,
  };
}
