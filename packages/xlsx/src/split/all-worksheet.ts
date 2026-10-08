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
  type RandomAccessSource,
} from "@consultchimps/core";

import { XLSX_ERRORS } from "../errors.js";
import { WorkbookModel } from "../model/index.js";
import {
  MACRO_WORKBOOK_MAIN_CONTENT_TYPE,
  WORKBOOK_MAIN_PART,
  WorkbookPackage,
  ZipReader,
} from "../package/index.js";
import { encodeCell } from "../model/references.js";
import type { NormalizedValue } from "../region/values.js";
import { yieldToEventLoop } from "../shared.js";
import {
  uncachedFormulaHint,
  uncachedFormulaWarnings,
} from "../uncached-formulas.js";
import {
  lightModel,
  openStreamedSplitPackage,
  readSplitColumn,
  SheetSummary,
  summaryRegion,
  writeSplitGroup,
  type StreamedSheetRegion,
  type StreamedSplitPackage,
} from "./streamed-package.js";

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

/**
 * Everything a split learned by reading the source: the light package, each
 * filtered worksheet's region with one group number per body row, and the
 * groups. Both the plan and the run are computed from this, so a preview and
 * the split it previews cannot disagree.
 */
export interface AllWorksheetSplitAnalysis {
  extension: WorkbookExtension;
  groups: SplitGroup[];
  inputRows: number;
  mediaType: string;
  /** The workbook, its worksheets held as stubs and read again as needed. */
  package: StreamedSplitPackage;
  /** The filtered worksheets, in workbook order. */
  regions: StreamedSheetRegion[];
  skippedRows: number;
  unchangedSheets: string[];
  /**
   * `Sheet!B4` of every split-column cell holding a formula with no cached
   * value: the split reads it as blank, so its row joins no group.
   */
  uncachedSplitCells: string[];
}

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
/**
 * `preservedSplitExtension` for a workbook read in pieces: only its content
 * types are read to learn whether it declares macros.
 */
export async function preservedSplitExtensionOf(
  source: RandomAccessSource,
  name: string,
  identity: SplitSourceIdentity,
): Promise<WorkbookExtension> {
  const extension = workbookExtensionOf(name, identity);
  let declaresMacroWorkbook: boolean;
  try {
    const zip = await ZipReader.open(source);
    const contentTypes = await zip.readBytes("[Content_Types].xml");
    const workbookPackage = WorkbookPackage.fromParts(
      contentTypes
        ? [
            {
              name: "[Content_Types].xml",
              bytes: contentTypes,
              date: new Date(0),
            },
          ]
        : [],
    );
    declaresMacroWorkbook =
      workbookPackage.contentTypeOverride(WORKBOOK_MAIN_PART)?.trim() ===
      MACRO_WORKBOOK_MAIN_CONTENT_TYPE;
  } catch (error) {
    throw workbookStructureFailure(identity, error);
  }
  refuseMislabelledWorkbookType(declaresMacroWorkbook, extension, identity);
  return extension;
}

function workbookStructureFailure(
  identity: SplitSourceIdentity,
  error: unknown,
): ConsultChimpsError {
  return new ConsultChimpsError(
    XLSX_ERRORS.XLSX_READ_FAILED,
    `Could not inspect workbook structure: ${identity.label}. The file may be corrupted, encrypted, or contain invalid workbook XML.`,
    { cause: error, details: { ...identity.details } },
  );
}

/**
 * Read the source for a split: each worksheet once to find its header and
 * columns, and each worksheet carrying the column once more for its split
 * values. Nothing proportional to the rows is kept but one group number per
 * body row (see `streamed-package.ts`).
 */
export async function analyzeAllWorksheetSplit(
  source: RandomAccessSource,
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

  const summaries = new Map<string, SheetSummary>();
  const uncachedByPart = new Map<string, string[]>();
  let splitPackage: StreamedSplitPackage;
  let workbook: WorkbookModel;
  try {
    const zip = await ZipReader.open(source);
    splitPackage = await openStreamedSplitPackage(zip, (listed) => {
      const host = WorkbookModel.fromPackage(listed);
      const names = new Map(
        host.sheets.map((sheet) => [sheet.partPath, sheet.name] as const),
      );
      return (part, prefix) => {
        const summary = new SheetSummary(
          host,
          selection.column,
          selection.headerRow,
          prefix,
          selection.values === true,
        );
        summaries.set(part, summary);
        const uncached: string[] = [];
        uncachedByPart.set(part, uncached);
        const sheetName = names.get(part) ?? part;
        return {
          row: (streamed) => {
            summary.row(streamed);
            for (const cell of streamed.parse().cells) {
              if (cell.formula !== undefined && !cell.hasCachedValue) {
                uncached.push(
                  `${sheetName}!${encodeCell(cell.column, cell.row)}`,
                );
              }
            }
          },
          end: (suffix) => {
            summary.end(suffix);
          },
        };
      };
    });
    workbook = lightModel(splitPackage);
  } catch (error) {
    if (error instanceof ConsultChimpsError) throw error;
    throw workbookStructureFailure(identity, error);
  }
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

  const tables = await workbook.tables();
  const found: Array<{
    sheet: { name: string; partPath: string };
    summary: SheetSummary;
    region: NonNullable<ReturnType<typeof summaryRegion>>;
  }> = [];
  const unchangedSheets: string[] = [];
  for (const sheet of workbook.sheets) {
    const summary = summaries.get(sheet.partPath);
    if (!summary) {
      // Not read as a stream: the model reads it, and refuses it, as before.
      workbook.worksheet(sheet.name);
      unchangedSheets.push(sheet.name);
      continue;
    }
    const region = summaryRegion(summary, sheet.name, selection.column, tables);
    if (!region) {
      unchangedSheets.push(sheet.name);
      continue;
    }
    found.push({ sheet, summary, region });
  }
  if (found.length === 0) {
    const uncached = workbook.sheets.flatMap(
      (sheet) => uncachedByPart.get(sheet.partPath) ?? [],
    );
    const error = columnNotFound();
    if (uncached.length === 0) throw error;
    throw new ConsultChimpsError(
      error.code,
      `${error.message}${uncachedFormulaHint(uncached)}`,
      { details: { ...error.details, uncachedFormulas: uncached } },
    );
  }

  const strict = selection.strict === true;
  const groups: SplitGroup[] = [];
  const groupIndex = new Map<string, number>();
  const regions: StreamedSheetRegion[] = [];
  const uncachedSplitCells: string[] = [];
  let inputRows = 0;
  let skippedRows = 0;
  for (const { sheet, summary, region } of found) {
    // A group is named by the first row, top to bottom, carrying its value.
    const local: Array<{ value: NormalizedValue; row: number }> = [];
    const localIndex = new Map<string, number>();
    const stub = splitPackage.stubs.get(sheet.partPath)!;
    const ids = await readSplitColumn(
      stub.feed(),
      workbook,
      region,
      sheet.name,
      strict,
      (value, row) => {
        let id = localIndex.get(value.key);
        if (id === undefined) {
          id = local.length;
          localIndex.set(value.key, id);
          local.push({ value, row });
        } else if (row < local[id]!.row) {
          local[id] = { value, row };
        }
        return id;
      },
      uncachedSplitCells,
    );
    for (let index = 0; index < ids.length; index += 1) {
      inputRows += 1;
      const id = ids[index]!;
      if (id < 0) {
        skippedRows += 1;
        continue;
      }
      const value = local[id]!.value;
      let group = groupIndex.get(value.key);
      if (group === undefined) {
        group = groups.length;
        groupIndex.set(value.key, group);
        groups.push(value);
      }
      ids[index] = group;
    }
    regions.push({
      ...region,
      name: sheet.name,
      worksheetPart: sheet.partPath,
      declared: selection.headerRow !== undefined,
      lastRow: summary.lastRow,
      guard: summary.guard ?? { canRenumber: true },
      guardAfterValues: summary.guardAfterValues ?? { canRenumber: true },
      groupOfRow: ids,
    });
  }

  if (groups.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_NO_GROUPS,
      `Column "${selection.column}" does not contain any non-blank values. Add at least one value and try again.${uncachedFormulaHint(uncachedSplitCells)}`,
      {
        details: {
          column: selection.column,
          uncachedFormulas: uncachedSplitCells,
          ...identity.details,
        },
      },
    );
  }

  return {
    extension,
    groups,
    inputRows,
    mediaType: splitMediaType(extension),
    package: splitPackage,
    regions,
    skippedRows,
    unchangedSheets,
    uncachedSplitCells,
  };
}

export function plannedAllWorksheetSplitWarnings(
  analysis: AllWorksheetSplitAnalysis,
): string[] {
  return uncachedFormulaWarnings(
    analysis.uncachedSplitCells,
    "they read as blank, so a row whose split value is one joins no group",
  );
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
    // What the plan read: the split-column cells. A values-only run also
    // counts what its conversion loses, which only the run can see.
    formulaCellsWithoutCachedValues: analysis.uncachedSplitCells.length,
    groups: analysis.groups.length,
    inputFiles: 1,
    inputRows: analysis.inputRows,
    outputFiles,
    pivotTablesRemoved: 0,
    rowsDeleted: 0,
    sheetsCopiedUnchanged: analysis.unchangedSheets.length,
    sheetsFiltered: analysis.regions.length,
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

/** Where one output's bytes go, opened by the surface for each group. */
export interface SplitOutputTarget {
  /** Receives the output's bytes in order. */
  write: (chunk: Uint8Array) => void | Promise<void>;
  /** Called once every byte is written, with what the output holds. */
  close: (detail: SplitOutputDetail) => Promise<void> | void;
  /** Called instead of `close` when the output cannot be finished. */
  abort?: () => Promise<void> | void;
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
   * Open the output for one group. The surface decides what that means -
   * a staging file in a transaction directory, or chunks in memory - and
   * reports its own progress, because the two surfaces describe different work.
   */
  open: (index: number) => Promise<SplitOutputTarget> | SplitOutputTarget;
}

/**
 * Write every group's workbook in group order, each to the target `open`
 * gives, and report what the whole split did.
 *
 * Outputs are produced one at a time and written as they are produced: the
 * largest thing held is one compressed part of one output.
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
    const target = await options.open(index);
    let built: Awaited<ReturnType<typeof writeSplitGroup>>;
    try {
      built = await writeSplitGroup(
        analysis.package,
        analysis.regions,
        index,
        selection.values === true,
        target.write,
        async () => {
          await yieldToEventLoop();
          throwIfAborted(signal, SPLIT_OPERATION, outputContext);
        },
      );
    } catch (error) {
      await Promise.resolve(target.abort?.()).catch(() => undefined);
      throw error;
    }
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

    const sheets = analysis.regions.map((region) => {
      let retainedRows = 0;
      for (const id of region.groupOfRow) {
        if (id === index) retainedRows += 1;
      }
      const deletedRows = region.groupOfRow.length - retainedRows;
      outputRows += retainedRows;
      rowsDeleted += deletedRows;
      return { deletedRows, retainedRows, sheet: region.name };
    });

    const detail: SplitOutputDetail = {
      formulaCellsConverted: built.formulaCellsConverted,
      formulaCellsWithoutCachedValues: built.formulaCellsWithoutCachedValues,
      output: outputNames[index]!,
      sheets,
      value: group.display,
    };
    await target.close(detail);
    outputs.push(detail);
  }

  throwIfAborted(signal, SPLIT_OPERATION, outputContext);

  const warnings: string[] = [];
  if (analysis.skippedRows > 0) {
    warnings.push(
      `Skipped ${analysis.skippedRows} row${analysis.skippedRows === 1 ? "" : "s"} with blank values in "${selection.column}"; no blank-value workbook was created.`,
    );
  }
  // Like the other per-output counts, a values-only loss counts once per
  // output it happens in. A split-column cell read as blank is added once,
  // unless an output already lost it.
  const splitCellsOnly = analysis.uncachedSplitCells.filter(
    (location) => !missingFormulaLocations.has(location),
  );
  formulaCellsWithoutCachedValues += splitCellsOnly.length;
  for (const warning of uncachedFormulaWarnings(
    [...splitCellsOnly, ...missingFormulaLocations],
    selection.values === true
      ? "a row whose split value is one joins no group, and any other became a blank cell in the values-only outputs"
      : "they read as blank, so a row whose split value is one joins no group",
    formulaCellsWithoutCachedValues,
  )) {
    warnings.push(warning);
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
      sheetsFiltered: analysis.regions.length,
      skippedRows: analysis.skippedRows,
      valuesOnly: selection.values === true ? 1 : 0,
    },
    outputs,
    summary: {
      column: selection.column,
      copiedUnchangedSheets: analysis.unchangedSheets,
      filteredSheets: analysis.regions.map((region) => region.name),
      input: identity.label,
      valuesOnly: selection.values === true,
    },
    warnings,
  };
}
