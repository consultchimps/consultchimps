/**
 * L3/L5 - the all-worksheet split, composed over the layered engine.
 *
 * This operation used to be its own miniature spreadsheet library: it opened
 * the ZIP, matched `<row>` elements with a regular expression, renumbered `r`
 * attributes by hand, and knew nothing about the merged ranges, sqrefs,
 * hyperlinks, comment anchors and formulas that described the rows it moved.
 * All of that now lives below it - `resolveRegions` finds the data, a
 * `DataRegion` decides which rows go, and L1's invariant pass moves everything
 * that pointed at them. What is left here is what an operation should be: file
 * handling, option mapping, staging, metrics and messages.
 *
 * Behaviour is deliberately unchanged where the corpus pins it, quirks
 * included: a hidden worksheet carrying the split column is filtered like any
 * other, the range binding treats every row below the header as data (so it
 * removes a totals row and a footer block that the table binding keeps), and
 * outputs are named after the group value with no prefix unless one is asked
 * for.
 */
import {
  mkdtemp,
  readFile,
  rename,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import path from "node:path";

import {
  ConsultChimpsError,
  throwIfAborted,
  type OperationControlOptions,
  type OperationPlan,
  type OperationResult,
} from "@consultchimps/core";
import {
  ensureDirectory,
  ensureOutputAvailable,
  refuseInputOverwrite,
} from "@consultchimps/files";

import { XLSX_ERRORS } from "./errors.js";
import { WorkbookModel } from "./model/index.js";
import type { RowNumber } from "./model/types.js";
import { isTableEditReport } from "./region/table-binding.js";
import type { DataRegion } from "./region/types.js";
import {
  normalizeHeader,
  readRowValues,
  type NormalizedValue,
} from "./region/values.js";
import { resolveRegions } from "./region/resolve.js";
import { safeFilenameSegment, splitOutputPaths } from "./split-filenames.js";
import { stripPivotParts } from "./tier1/pivot.js";
import { blankStaleCachedFormulas } from "./tier1/stale-values.js";
import { convertWorkbookToValuesWithReport } from "./values-only.js";

const SPLIT_OPERATION = "sheets.split-by-column";
const XLSX_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const XLSM_MEDIA_TYPE = "application/vnd.ms-excel.sheet.macroEnabled.12";

export type FullWorkbookSplitMetric =
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

export interface FullWorkbookSplitOptions extends OperationControlOptions {
  column: string;
  filenamePrefix?: string | undefined;
  headerRow?: number | undefined;
  input: string;
  outputDirectory: string;
  overwrite?: boolean | undefined;
  strict?: boolean | undefined;
  values?: boolean | undefined;
}

export interface SplitSheetDetail {
  deletedRows: number;
  retainedRows: number;
  sheet: string;
}

export interface SplitOutputDetail {
  formulaCellsConverted: number;
  formulaCellsWithoutCachedValues: number;
  output: string;
  sheets: SplitSheetDetail[];
  value: string;
}

export interface FullWorkbookSplitSummary {
  column: string;
  copiedUnchangedSheets: string[];
  filteredSheets: string[];
  input: string;
  outputDirectory: string;
  valuesOnly: boolean;
}

export interface FullWorkbookSplitResult extends OperationResult<FullWorkbookSplitMetric> {
  outputs: SplitOutputDetail[];
  summary: FullWorkbookSplitSummary;
}

type SplitGroup = NormalizedValue;

/** One filtered worksheet, as the analysis pass understood it. */
interface SheetAnalysis {
  /** True when the region is an Excel Table rather than a worksheet range. */
  isTable: boolean;
  name: string;
  /** Normalized key per body row; every body row has an entry. */
  rowValues: Map<RowNumber, NormalizedValue | undefined>;
  worksheetPart: string;
}

interface ResolvedFullWorkbookSplit {
  absoluteInput: string;
  absoluteOutputDirectory: string;
  existingOutputs: Set<string>;
  extension: ".xlsm" | ".xlsx";
  groups: SplitGroup[];
  inputRows: number;
  mediaType: string;
  outputPaths: string[];
  sheets: SheetAnalysis[];
  skippedRows: number;
  unchangedSheets: string[];
  workbookBytes: Buffer;
  /** Sheet name by worksheet part, for every sheet, filtered or not. */
  worksheetNames: ReadonlyMap<string, string>;
}

function isMissingPathError(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    error.code === "ENOENT"
  );
}

function isColumnNotFound(error: unknown): boolean {
  return (
    error instanceof ConsultChimpsError &&
    error.code === XLSX_ERRORS.XLSX_SPLIT_COLUMN_NOT_FOUND
  );
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
  absoluteInput: string,
): Promise<WorkbookModel> {
  try {
    return await WorkbookModel.load(workbookBytes);
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not inspect workbook structure: ${absoluteInput}. The file may be corrupted, encrypted, or contain invalid workbook XML.`,
      { cause: error, details: { inputPath: absoluteInput } },
    );
  }
}

async function resolveFullWorkbookSplit(
  options: FullWorkbookSplitOptions,
): Promise<ResolvedFullWorkbookSplit> {
  const absoluteInput = path.resolve(options.input);
  const extension = path.extname(absoluteInput).toLowerCase();
  if (extension !== ".xlsx" && extension !== ".xlsm") {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_UNSUPPORTED_FILE,
      `Unsupported Excel workbook type "${extension || "(none)"}". Choose an .xlsx or .xlsm workbook.`,
      { details: { inputPath: absoluteInput } },
    );
  }
  if (options.headerRow !== undefined && options.headerRow < 1) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_HEADER_ROW,
      "The header row must be a positive whole number counted from 1.",
      { details: { headerRow: options.headerRow } },
    );
  }

  let workbookBytes: Buffer;
  try {
    const inputStat = await stat(absoluteInput);
    if (!inputStat.isFile()) {
      throw new Error("The input path is not a file.");
    }
    workbookBytes = await readFile(absoluteInput);
  } catch (error) {
    if (isMissingPathError(error)) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_INPUT_NOT_FOUND,
        `Input workbook was not found: ${absoluteInput}. Check the path and try again.`,
        { cause: error, details: { inputPath: absoluteInput } },
      );
    }
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${absoluteInput}. Confirm that the file exists, is a valid unencrypted Excel workbook, and is not password protected.`,
      { cause: error, details: { inputPath: absoluteInput } },
    );
  }

  const workbook = await loadWorkbookModel(workbookBytes, absoluteInput);
  const columnNotFound = (): ConsultChimpsError =>
    new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_COLUMN_NOT_FOUND,
      `Column "${options.column}" was not found in any worksheet. Check the header text or provide the correct header row.`,
      {
        details: {
          availableWorksheets: workbook.sheets.map((sheet) => sheet.name),
          column: options.column,
          headerRow: options.headerRow,
          inputPath: absoluteInput,
        },
      },
    );
  if (options.column.trim() === "") {
    throw columnNotFound();
  }

  const { regions, unchangedSheets } = await resolveSplitRegions(
    workbook,
    options.column,
    options.headerRow,
  );
  if (regions.length === 0) {
    throw columnNotFound();
  }

  const matching = { strict: options.strict === true };
  const groupsByKey = new Map<string, SplitGroup>();
  const sheets: SheetAnalysis[] = [];
  let inputRows = 0;
  let skippedRows = 0;

  for (const region of regions) {
    const column = splitColumnOf(region, options.column)!;
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
      `Column "${options.column}" does not contain any non-blank values. Add at least one value and try again.`,
      { details: { column: options.column, inputPath: absoluteInput } },
    );
  }

  const absoluteOutputDirectory = path.resolve(options.outputDirectory);
  const prefix = options.filenamePrefix
    ? safeFilenameSegment(options.filenamePrefix, "split")
    : undefined;
  const outputPaths = splitOutputPaths(
    absoluteOutputDirectory,
    prefix,
    groups.map((group) => group.display),
    extension,
  );
  outputPaths.forEach((outputPath) =>
    refuseInputOverwrite(outputPath, [absoluteInput]),
  );
  const existingOutputs = new Set<string>();
  await Promise.all(
    outputPaths.map(async (outputPath) => {
      try {
        const outputStat = await stat(outputPath);
        if (!outputStat.isFile()) {
          throw new ConsultChimpsError(
            XLSX_ERRORS.XLSX_SPLIT_OUTPUT_NOT_FILE,
            `Output path exists but is not a file: ${outputPath}`,
            { details: { outputPath } },
          );
        }
        existingOutputs.add(outputPath);
      } catch (error) {
        if (!isMissingPathError(error)) {
          throw error;
        }
      }
    }),
  );

  return {
    absoluteInput,
    absoluteOutputDirectory,
    existingOutputs,
    extension,
    groups,
    inputRows,
    mediaType: extension === ".xlsm" ? XLSM_MEDIA_TYPE : XLSX_MEDIA_TYPE,
    outputPaths,
    sheets,
    skippedRows,
    unchangedSheets,
    workbookBytes,
    worksheetNames: new Map(
      workbook.sheets.map((sheet) => [sheet.partPath, sheet.name] as const),
    ),
  };
}

export async function planFullWorkbookSplit(
  options: FullWorkbookSplitOptions,
): Promise<OperationPlan<Exclude<FullWorkbookSplitMetric, "outputRows">>> {
  const resolved = await resolveFullWorkbookSplit(options);
  const warnings: string[] = [];
  if (resolved.existingOutputs.size > 0 && options.overwrite !== true) {
    warnings.push(
      `${resolved.existingOutputs.size} planned output file${resolved.existingOutputs.size === 1 ? " already exists" : "s already exist"}; executing without overwrite will fail.`,
    );
  }
  return {
    operation: SPLIT_OPERATION,
    inputs: [resolved.absoluteInput],
    outputs: resolved.outputPaths.map((outputPath) => ({
      exists: resolved.existingOutputs.has(outputPath),
      kind: "file",
      mediaType: resolved.mediaType,
      path: outputPath,
    })),
    warnings,
    metrics: {
      calcChainEntriesRemoved: 0,
      formulaCellsBlankedForRemovedRows: 0,
      formulaCellsConverted: 0,
      formulaCellsWithoutCachedValues: 0,
      groups: resolved.groups.length,
      inputFiles: 1,
      inputRows: resolved.inputRows,
      outputFiles: resolved.outputPaths.length,
      pivotTablesRemoved: 0,
      rowsDeleted: 0,
      sheetsCopiedUnchanged: resolved.unchangedSheets.length,
      sheetsFiltered: resolved.sheets.length,
      skippedRows: resolved.skippedRows,
      valuesOnly: options.values === true ? 1 : 0,
    },
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
  resolved: ResolvedFullWorkbookSplit,
  group: SplitGroup,
  options: FullWorkbookSplitOptions,
): Promise<GroupOutput> {
  const output: GroupOutput = {
    bytes: resolved.workbookBytes,
    calcChainEntriesRemoved: 0,
    formulaCellsBlanked: 0,
    formulaCellsConverted: 0,
    formulaCellsWithoutCachedValues: 0,
    pivotTablesRemoved: 0,
    staleAggregates: [],
    tableFallbackSheets: [],
    uncachedFormulas: [],
  };
  const worksheetNameByPart = resolved.worksheetNames;
  // A table region compacts its rows, so a values conversion has to run before
  // the filter for the cached results to line up with the rows they describe.
  // A pure worksheet split converts afterwards, over the rows that survive.
  const containsFilteredTable = resolved.sheets.some((sheet) => sheet.isTable);

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

  if (options.values) {
    // A values-only conversion bakes each formula's cached result into the
    // output, so any result computed over rows this group does not receive is
    // cleared first, while row numbers are still the source's.
    const staleValues = await blankStaleCachedFormulas(
      output.bytes,
      new Map(
        resolved.sheets.map(
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

  const workbook = await loadWorkbookModel(
    output.bytes,
    resolved.absoluteInput,
  );
  const { regions } = await resolveSplitRegions(
    workbook,
    options.column,
    options.headerRow,
  );
  for (const region of regions) {
    const analysis = resolved.sheets.find(
      (sheet) => sheet.name === region.sheetName,
    );
    if (!analysis) {
      continue;
    }
    const report = region.filterRows(
      (row) => analysis.rowValues.get(row)?.key === group.key,
    );
    if (isTableEditReport(report) && !report.tableResized) {
      output.tableFallbackSheets.push(region.sheetName);
    }
  }
  output.bytes = await workbook.save();
  // The invariant pass maintained the chain as the rows moved; a values-only
  // output has no chain left to maintain, because the conversion removes it.
  output.calcChainEntriesRemoved = options.values
    ? 0
    : workbook.calcChainEntriesRemoved;

  if (options.values && !containsFilteredTable) {
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

export async function splitFullWorkbookByColumn(
  options: FullWorkbookSplitOptions,
): Promise<FullWorkbookSplitResult> {
  throwIfAborted(options.signal, SPLIT_OPERATION);
  const resolved = await resolveFullWorkbookSplit(options);
  await Promise.all(
    resolved.outputPaths.map((outputPath) =>
      ensureOutputAvailable(outputPath, { overwrite: options.overwrite }),
    ),
  );
  await ensureDirectory(resolved.absoluteOutputDirectory);
  const transactionDirectory = await mkdtemp(
    path.join(resolved.absoluteOutputDirectory, ".consultchimps-split-"),
  );
  const stagedOutputs: string[] = [];
  const committedOutputs: string[] = [];
  const backups = new Map<string, string>();
  const outputDetails: SplitOutputDetail[] = [];
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

  try {
    for (const [index, group] of resolved.groups.entries()) {
      throwIfAborted(options.signal, SPLIT_OPERATION);
      const built = await buildGroupWorkbook(resolved, group, options);
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

      const sheets = resolved.sheets.map((sheet) => {
        const retainedRows = [...sheet.rowValues.values()].filter(
          (value) => value?.key === group.key,
        ).length;
        const deletedRows = sheet.rowValues.size - retainedRows;
        outputRows += retainedRows;
        rowsDeleted += deletedRows;
        return { deletedRows, retainedRows, sheet: sheet.name };
      });

      const stagedOutput = path.join(
        transactionDirectory,
        `output-${String(index + 1).padStart(6, "0")}${resolved.extension}`,
      );
      await writeFile(stagedOutput, built.bytes);
      stagedOutputs.push(stagedOutput);
      outputDetails.push({
        formulaCellsConverted: built.formulaCellsConverted,
        formulaCellsWithoutCachedValues: built.formulaCellsWithoutCachedValues,
        output: resolved.outputPaths[index]!,
        sheets,
        value: group.display,
      });
      options.onProgress?.({
        operation: SPLIT_OPERATION,
        stage: "staging-workbooks",
        completed: index + 1,
        total: resolved.groups.length,
        detail: `${path.basename(resolved.outputPaths[index]!)} (${sheets.map((sheet) => `${sheet.sheet}: kept ${sheet.retainedRows}, deleted ${sheet.deletedRows}`).join("; ")})`,
      });
    }

    for (const [index, outputPath] of resolved.outputPaths.entries()) {
      const stagedOutput = stagedOutputs[index];
      if (!stagedOutput) {
        continue;
      }
      if (resolved.existingOutputs.has(outputPath)) {
        const backupPath = path.join(
          transactionDirectory,
          `backup-${String(index + 1).padStart(6, "0")}${resolved.extension}`,
        );
        await rename(outputPath, backupPath);
        backups.set(outputPath, backupPath);
      }
      await rename(stagedOutput, outputPath);
      committedOutputs.push(outputPath);
      options.onProgress?.({
        operation: SPLIT_OPERATION,
        stage: "committing-outputs",
        completed: index + 1,
        total: resolved.outputPaths.length,
        detail: path.basename(outputPath),
      });
    }
  } catch (error) {
    const rollbackErrors: unknown[] = [];
    for (const outputPath of [...committedOutputs].reverse()) {
      try {
        await rm(outputPath, { force: true });
        const backupPath = backups.get(outputPath);
        if (backupPath) {
          await rename(backupPath, outputPath);
          backups.delete(outputPath);
        }
      } catch (rollbackError) {
        rollbackErrors.push(rollbackError);
      }
    }
    for (const [outputPath, backupPath] of backups) {
      try {
        await rename(backupPath, outputPath);
      } catch (rollbackError) {
        rollbackErrors.push(rollbackError);
      }
    }
    if (rollbackErrors.length > 0) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_SPLIT_ROLLBACK_FAILED,
        "The split failed and one or more output files could not be restored.",
        {
          cause: error,
          details: {
            outputPaths: resolved.outputPaths,
            rollbackErrors: rollbackErrors.map((rollbackError) =>
              rollbackError instanceof Error
                ? rollbackError.message
                : String(rollbackError),
            ),
          },
        },
      );
    }
    throw error;
  } finally {
    await rm(transactionDirectory, { force: true, recursive: true }).catch(
      () => undefined,
    );
  }

  const warnings: string[] = [];
  if (resolved.skippedRows > 0) {
    warnings.push(
      `Skipped ${resolved.skippedRows} row${resolved.skippedRows === 1 ? "" : "s"} with blank values in "${options.column}"; no blank-value workbook was created.`,
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
    operation: SPLIT_OPERATION,
    artifacts: resolved.outputPaths.map((outputPath) => ({
      kind: "file",
      mediaType: resolved.mediaType,
      path: outputPath,
    })),
    warnings,
    metrics: {
      calcChainEntriesRemoved,
      formulaCellsBlankedForRemovedRows,
      formulaCellsConverted,
      formulaCellsWithoutCachedValues,
      groups: resolved.groups.length,
      inputFiles: 1,
      inputRows: resolved.inputRows,
      outputFiles: resolved.outputPaths.length,
      outputRows,
      pivotTablesRemoved,
      rowsDeleted,
      sheetsCopiedUnchanged: resolved.unchangedSheets.length,
      sheetsFiltered: resolved.sheets.length,
      skippedRows: resolved.skippedRows,
      valuesOnly: options.values === true ? 1 : 0,
    },
    outputs: outputDetails,
    summary: {
      column: options.column,
      copiedUnchangedSheets: resolved.unchangedSheets,
      filteredSheets: resolved.sheets.map((sheet) => sheet.name),
      input: resolved.absoluteInput,
      outputDirectory: resolved.absoluteOutputDirectory,
      valuesOnly: options.values === true,
    },
  };
}
