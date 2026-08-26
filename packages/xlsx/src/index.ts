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
  safeNameFragment,
  throwIfAborted,
  type OperationControlOptions,
  type OperationPlan,
  type OperationResult,
  type PlannedOutput,
} from "@consultchimps/core";
import {
  ensureDirectory,
  ensureOutputAvailable,
  ensureParentDirectory,
  pathExists,
  refuseInputOverwrite,
} from "@consultchimps/files";
import type { Table } from "@consultchimps/tabular";
import type * as XLSX from "xlsx";

import { XLSX_ERRORS } from "./errors.js";
import {
  type FullWorkbookSplitMetric,
  type FullWorkbookSplitSummary,
  type SplitOutputDetail,
  planFullWorkbookSplit,
  splitFullWorkbookByColumn,
} from "./workbook-column-split.js";
import {
  appendWorkbookSheets,
  buildSplitGroupBytes,
  buildTableWorkbookBytes,
  CONSOLIDATE_OPERATION,
  CONSOLIDATED_SHEET_NAME,
  consolidateTables,
  createMergeState,
  finishMergedWorkbook,
  isMacroWorkbookName,
  MACRO_WORKBOOK_MEDIA_TYPE,
  MERGE_OPERATION,
  parseExcelTableDefinitions,
  parseWorkbookBytes,
  preservedSplitTemplateBytes,
  resolveSplitSource,
  skippedRowsWarning,
  splitOutputFileNames,
  SPLIT_OPERATION,
  WORKBOOK_MEDIA_TYPE,
  workbookExcelTables,
  workbookNamedRanges,
  workbookTables,
  workbookWorksheetRecords,
  type ResolvedSplitSource,
} from "./shared.js";

import type {
  ConsolidateWorkbooksMetric,
  ConsolidateWorkbooksPlanMetric,
  MergeWorkbooksMetric,
  ReadWorkbookExcelTablesOptions,
  ReadWorkbookNamedRangesOptions,
  ReadWorkbookOptions,
  ReadWorksheetRecordsOptions,
  WorkbookExcelTable,
  WorkbookNamedRange,
  WorksheetRecords,
} from "./shared.js";

export { XLSX_ERRORS, type XlsxErrorCode } from "./errors.js";
export type {
  ConsolidateWorkbooksMetric,
  ConsolidateWorkbooksPlanMetric,
  MergeWorkbooksMetric,
  ReadWorkbookExcelTablesOptions,
  ReadWorkbookNamedRangesOptions,
  ReadWorkbookOptions,
  ReadWorksheetRecordsOptions,
  WorkbookExcelTable,
  WorkbookNamedRange,
  WorksheetRecords,
};

export type SplitWorkbookByColumnMetric = FullWorkbookSplitMetric;
export type SplitWorkbookByColumnPlanMetric = Exclude<
  SplitWorkbookByColumnMetric,
  "outputRows"
>;

export interface ConsolidateWorkbooksOptions
  extends ReadWorkbookOptions, OperationControlOptions {
  inputs: string[];
  output: string;
  addSourceColumns?: boolean | undefined;
  /**
   * Match columns whose headers differ only in case, spacing, or punctuation
   * (for example "Failed Checks" and "Failed_Checks") instead of requiring
   * the exact same header in every worksheet.
   */
  normalizeHeaders?: boolean | undefined;
  outputSheetName?: string | undefined;
  overwrite?: boolean | undefined;
  values?: boolean | undefined;
}

export interface MergeWorkbooksOptions extends OperationControlOptions {
  includeSheetIndex?: boolean | undefined;
  overwrite?: boolean | undefined;
  values?: boolean | undefined;
}

export interface SplitWorkbookByColumnOptions extends OperationControlOptions {
  input: string;
  outputDirectory: string;
  column: string;
  filenamePrefix?: string | undefined;
  headerRow?: number | undefined;
  includeBlank?: boolean | undefined;
  includeHiddenSheets?: boolean | undefined;
  overwrite?: boolean | undefined;
  /**
   * Keep the complete source workbook. In the default all-worksheet mode this
   * is enabled unless explicitly set to false; table splits also default to
   * true. Named-range and selected-worksheet splits remain compact when this
   * option is false.
   */
  preserveWorkbook?: boolean | undefined;
  range?: string | undefined;
  sheet?: string | undefined;
  table?: string | undefined;
  /** Compare split values without trimming, case folding, or numeric coercion. */
  strict?: boolean | undefined;
  values?: boolean | undefined;
}

export interface SplitWorkbookByColumnResult extends OperationResult<SplitWorkbookByColumnMetric> {
  /** Per-output, per-worksheet filtering details for all-worksheet splits. */
  outputs?: SplitOutputDetail[] | undefined;
  /** Input, output, option, and worksheet summary for all-worksheet splits. */
  summary?: FullWorkbookSplitSummary | undefined;
}

export interface WriteTableOptions {
  overwrite?: boolean | undefined;
  sheetName?: string | undefined;
}

interface WorkbookFile {
  bytes: Buffer;
  workbook: XLSX.WorkBook;
}

function isMissingPathError(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    error.code === "ENOENT"
  );
}

/**
 * Read a workbook's bytes from disk, reporting a missing or unreadable file as
 * the stable read error the parsing readers also raise.
 */
async function readWorkbookBytes(absolutePath: string): Promise<Uint8Array> {
  try {
    return await readFile(absolutePath);
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${absolutePath}`,
      { cause: error, details: { filePath: absolutePath } },
    );
  }
}

/**
 * Read and parse a workbook from disk, reporting both a missing file and an
 * unreadable workbook as the same stable error.
 */
async function readWorkbookFile(
  absolutePath: string,
  options: { cellText?: boolean } = {},
): Promise<WorkbookFile> {
  const details = { filePath: absolutePath };
  let bytes: Buffer;

  try {
    bytes = await readFile(absolutePath);
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${absolutePath}`,
      { cause: error, details },
    );
  }

  return {
    bytes,
    workbook: parseWorkbookBytes(bytes, absolutePath, { ...options, details }),
  };
}

export async function readWorkbookTables(
  filePath: string,
  options: ReadWorkbookOptions = {},
): Promise<Table[]> {
  const absolutePath = path.resolve(filePath);
  const { workbook } = await readWorkbookFile(absolutePath);
  return workbookTables(workbook, path.basename(absolutePath), options);
}

export async function readWorksheetRecords(
  filePath: string,
  options: ReadWorksheetRecordsOptions,
): Promise<WorksheetRecords> {
  const absolutePath = path.resolve(filePath);
  const { workbook } = await readWorkbookFile(absolutePath, {
    cellText: true,
  });
  return workbookWorksheetRecords(workbook, options);
}

export async function readWorkbookExcelTables(
  filePath: string,
  options: ReadWorkbookExcelTablesOptions = {},
): Promise<WorkbookExcelTable[]> {
  const absolutePath = path.resolve(filePath);
  const { bytes, workbook } = await readWorkbookFile(absolutePath);
  const definitions = await parseExcelTableDefinitions(bytes, absolutePath, {
    filePath: absolutePath,
  });
  return workbookExcelTables(
    workbook,
    definitions,
    path.basename(absolutePath),
    options,
  );
}

export async function readWorkbookNamedRanges(
  filePath: string,
  options: ReadWorkbookNamedRangesOptions = {},
): Promise<WorkbookNamedRange[]> {
  const absolutePath = path.resolve(filePath);
  const { workbook } = await readWorkbookFile(absolutePath, {
    cellText: true,
  });
  return workbookNamedRanges(workbook, path.basename(absolutePath), options);
}

export async function writeTable(
  outputPath: string,
  table: Table,
  options: WriteTableOptions = {},
): Promise<string> {
  const bytes = buildTableWorkbookBytes(
    table,
    options.sheetName ?? CONSOLIDATED_SHEET_NAME,
  );
  const absoluteOutput = await ensureParentDirectory(outputPath);
  await ensureOutputAvailable(absoluteOutput, {
    overwrite: options.overwrite,
  });
  await writeFile(absoluteOutput, bytes);
  return absoluteOutput;
}

interface ResolvedConsolidate {
  absoluteInputs: string[];
  absoluteOutput: string;
}

function resolveConsolidateWorkbooks(
  options: ConsolidateWorkbooksOptions,
): ResolvedConsolidate {
  if (options.inputs.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_INPUTS,
      "At least one workbook is required.",
    );
  }

  const absoluteInputs = options.inputs.map((inputPath) =>
    path.resolve(inputPath),
  );
  const absoluteOutput = path.resolve(options.output);
  refuseInputOverwrite(absoluteOutput, absoluteInputs);
  return { absoluteInputs, absoluteOutput };
}

export async function planConsolidateWorkbooks(
  options: ConsolidateWorkbooksOptions,
): Promise<OperationPlan<ConsolidateWorkbooksPlanMetric>> {
  const { absoluteInputs, absoluteOutput } =
    resolveConsolidateWorkbooks(options);

  for (const absoluteInput of absoluteInputs) {
    if (!(await pathExists(absoluteInput))) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_INPUT_NOT_FOUND,
        `Workbook not found: ${absoluteInput}`,
        { details: { inputPath: absoluteInput } },
      );
    }
  }

  const exists = await pathExists(absoluteOutput);
  const warnings =
    exists && options.overwrite !== true
      ? [
          "The planned output workbook already exists; executing without overwrite will fail.",
        ]
      : [];

  return {
    operation: CONSOLIDATE_OPERATION,
    inputs: absoluteInputs,
    outputs: [
      {
        kind: "file",
        mediaType: WORKBOOK_MEDIA_TYPE,
        path: absoluteOutput,
        exists,
      },
    ],
    warnings,
    metrics: {
      inputFiles: absoluteInputs.length,
      outputFiles: 1,
    },
  };
}

export async function consolidateWorkbooks(
  options: ConsolidateWorkbooksOptions,
): Promise<OperationResult<ConsolidateWorkbooksMetric>> {
  throwIfAborted(options.signal, CONSOLIDATE_OPERATION);
  const { absoluteInputs, absoluteOutput } =
    resolveConsolidateWorkbooks(options);

  const tables: Table[] = [];
  for (const [index, absoluteInput] of absoluteInputs.entries()) {
    throwIfAborted(options.signal, CONSOLIDATE_OPERATION);
    tables.push(...(await readWorkbookTables(absoluteInput, options)));
    options.onProgress?.({
      operation: CONSOLIDATE_OPERATION,
      stage: "reading-workbooks",
      completed: index + 1,
      total: absoluteInputs.length,
      detail: path.basename(absoluteInput),
    });
  }

  const table = consolidateTables(tables, options);
  throwIfAborted(options.signal, CONSOLIDATE_OPERATION);
  const output = await writeTable(absoluteOutput, table, {
    overwrite: options.overwrite,
    sheetName: options.outputSheetName,
  });
  options.onProgress?.({
    operation: CONSOLIDATE_OPERATION,
    stage: "writing-output",
    completed: 1,
    total: 1,
    detail: path.basename(output),
  });

  return {
    operation: CONSOLIDATE_OPERATION,
    artifacts: [
      {
        kind: "file",
        mediaType: WORKBOOK_MEDIA_TYPE,
        path: output,
      },
    ],
    warnings: [],
    metrics: {
      inputFiles: absoluteInputs.length,
      inputTables: tables.length,
      outputColumns: table.columns.length,
      outputRows: table.rows.length,
    },
  };
}

export async function mergeWorkbooks(
  inputPaths: string[],
  outputPath: string,
  options: MergeWorkbooksOptions = {},
): Promise<OperationResult<MergeWorkbooksMetric>> {
  throwIfAborted(options.signal, MERGE_OPERATION);
  if (inputPaths.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_INPUTS,
      "At least one workbook is required.",
    );
  }
  const absoluteInputs = inputPaths.map((inputPath) => path.resolve(inputPath));
  const absoluteOutput = path.resolve(outputPath);
  refuseInputOverwrite(absoluteOutput, absoluteInputs);
  await ensureOutputAvailable(absoluteOutput, { overwrite: options.overwrite });

  const buildOptions = {
    ...options,
    // The user names the output, so the extension decides whether a macro
    // project may travel: a package must never claim a type its name denies.
    macroOutput: isMacroWorkbookName(absoluteOutput),
  };
  const state = createMergeState(buildOptions);
  for (const [index, inputPath] of absoluteInputs.entries()) {
    throwIfAborted(options.signal, MERGE_OPERATION);
    await appendWorkbookSheets(
      state,
      path.basename(inputPath),
      await readWorkbookBytes(inputPath),
    );
    options.onProgress?.({
      operation: MERGE_OPERATION,
      stage: "merging-inputs",
      completed: index + 1,
      total: absoluteInputs.length,
      detail: path.basename(inputPath),
    });
  }
  const merged = await finishMergedWorkbook(state, buildOptions);

  throwIfAborted(options.signal, MERGE_OPERATION);
  await ensureParentDirectory(absoluteOutput);
  await writeFile(absoluteOutput, merged.bytes);
  options.onProgress?.({
    operation: MERGE_OPERATION,
    stage: "writing-output",
    completed: 1,
    total: 1,
    detail: path.basename(absoluteOutput),
  });

  return {
    operation: MERGE_OPERATION,
    artifacts: [
      {
        kind: "file",
        mediaType: merged.macroEnabled
          ? MACRO_WORKBOOK_MEDIA_TYPE
          : WORKBOOK_MEDIA_TYPE,
        path: absoluteOutput,
      },
    ],
    warnings: merged.warnings,
    metrics: {
      inputFiles: inputPaths.length,
      outputSheets: merged.outputSheets,
      hiddenSheets: merged.hiddenSheets,
    },
  };
}

interface ResolvedSplit extends ResolvedSplitSource {
  absoluteInput: string;
  absoluteOutputDirectory: string;
  existingOutputs: Set<string>;
  outputPaths: string[];
  workbookBytes: Buffer;
}

async function resolveSplitWorkbookByColumn(
  options: SplitWorkbookByColumnOptions,
): Promise<ResolvedSplit> {
  const absoluteInput = path.resolve(options.input);
  const details = { inputPath: absoluteInput };
  let workbookBytes: Buffer;

  try {
    workbookBytes = await readFile(absoluteInput);
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${absoluteInput}`,
      { cause: error, details },
    );
  }

  const resolved = await resolveSplitSource(
    workbookBytes,
    {
      details,
      file: path.basename(absoluteInput),
      label: absoluteInput,
    },
    options,
  );

  const absoluteOutputDirectory = path.resolve(options.outputDirectory);
  const filenamePrefix = safeNameFragment(
    options.filenamePrefix ?? path.parse(absoluteInput).name,
    "split",
  );
  const outputPaths = splitOutputFileNames(
    filenamePrefix,
    resolved.grouped.groups.map((group) => group.value),
  ).map((filename) => path.join(absoluteOutputDirectory, filename));

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
    ...resolved,
    absoluteInput,
    absoluteOutputDirectory,
    existingOutputs,
    outputPaths,
    workbookBytes,
  };
}

export async function planSplitWorkbookByColumn(
  options: SplitWorkbookByColumnOptions,
): Promise<OperationPlan<SplitWorkbookByColumnPlanMetric>> {
  if (
    !options.table &&
    !options.range &&
    !options.sheet &&
    options.preserveWorkbook !== false
  ) {
    return planFullWorkbookSplit(options);
  }
  const resolved = await resolveSplitWorkbookByColumn(options);
  const outputs: PlannedOutput[] = resolved.outputPaths.map((outputPath) => ({
    kind: "file",
    mediaType: WORKBOOK_MEDIA_TYPE,
    path: outputPath,
    exists: resolved.existingOutputs.has(outputPath),
  }));

  const warnings: string[] = [];
  if (resolved.grouped.skippedRows > 0) {
    warnings.push(skippedRowsWarning(resolved.grouped));
  }
  const collisions = resolved.existingOutputs.size;
  if (collisions > 0 && options.overwrite !== true) {
    warnings.push(
      `${collisions} planned output file${
        collisions === 1 ? " already exists" : "s already exist"
      }; executing without overwrite will fail.`,
    );
  }

  return {
    operation: SPLIT_OPERATION,
    inputs: [resolved.absoluteInput],
    outputs,
    warnings,
    metrics: {
      calcChainEntriesRemoved: 0,
      formulaCellsBlankedForRemovedRows: 0,
      formulaCellsConverted: 0,
      formulaCellsWithoutCachedValues: 0,
      pivotTablesRemoved: 0,
      groups: resolved.grouped.groups.length,
      inputFiles: 1,
      inputRows: resolved.table.rows.length,
      outputFiles: resolved.outputPaths.length,
      rowsDeleted: 0,
      sheetsCopiedUnchanged: 0,
      sheetsFiltered: 1,
      skippedRows: resolved.grouped.skippedRows,
      valuesOnly: options.values === true ? 1 : 0,
    },
  };
}

/**
 * Split a workbook into one file per distinct value of a column.
 *
 * Two engines sit behind one signature. With no `table`, `range` or `sheet`
 * selection the split keeps the whole workbook and filters every worksheet
 * that carries the column: that path runs on the layered engine, so every
 * reference describing a moved row moves with it. Selecting a region instead
 * asks for one of the older, narrower modes - a preserved Excel Table rewrite,
 * or a compact single-worksheet rebuild - which `shared.ts` still owns.
 */
export async function splitWorkbookByColumn(
  options: SplitWorkbookByColumnOptions,
): Promise<SplitWorkbookByColumnResult> {
  if (
    !options.table &&
    !options.range &&
    !options.sheet &&
    options.preserveWorkbook !== false
  ) {
    return splitFullWorkbookByColumn(options);
  }
  throwIfAborted(options.signal, SPLIT_OPERATION);
  const {
    absoluteOutputDirectory,
    existingOutputs,
    grouped,
    outputPaths,
    preservedTableDefinition,
    preserveWorkbook,
    table,
    workbookBytes,
  } = await resolveSplitWorkbookByColumn(options);

  await Promise.all(
    outputPaths.map((outputPath) =>
      ensureOutputAvailable(outputPath, { overwrite: options.overwrite }),
    ),
  );
  await ensureDirectory(absoluteOutputDirectory);

  const transactionDirectory = await mkdtemp(
    path.join(absoluteOutputDirectory, ".consultchimps-split-"),
  );
  const stagedOutputs: string[] = [];
  const committedOutputs: string[] = [];
  const backups = new Map<string, string>();
  const templateBytes = preserveWorkbook
    ? await preservedSplitTemplateBytes(workbookBytes, options.values)
    : undefined;
  let pivotTablesRemoved = 0;

  try {
    for (const [index, group] of grouped.groups.entries()) {
      throwIfAborted(options.signal, SPLIT_OPERATION);
      const stagedOutput = path.join(
        transactionDirectory,
        `output-${String(index + 1).padStart(6, "0")}.xlsx`,
      );
      await writeFile(
        stagedOutput,
        await buildSplitGroupBytes(group, {
          onPivotTablesRemoved: (removed) => {
            pivotTablesRemoved += removed;
          },
          preservedTableDefinition,
          sheetName: table.source?.sheet ?? "Split",
          templateBytes,
        }),
      );
      stagedOutputs.push(stagedOutput);
      options.onProgress?.({
        operation: SPLIT_OPERATION,
        stage: "staging-workbooks",
        completed: index + 1,
        total: grouped.groups.length,
        detail: path.basename(outputPaths[index] ?? stagedOutput),
      });
    }

    for (const [index, outputPath] of outputPaths.entries()) {
      const stagedOutput = stagedOutputs[index];
      if (!stagedOutput) {
        continue;
      }

      if (existingOutputs.has(outputPath)) {
        const backupPath = path.join(
          transactionDirectory,
          `backup-${String(index + 1).padStart(6, "0")}.xlsx`,
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
        total: outputPaths.length,
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
            outputPaths,
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

  const warnings = grouped.skippedRows > 0 ? [skippedRowsWarning(grouped)] : [];
  if (pivotTablesRemoved > 0) {
    warnings.push(
      `Removed ${pivotTablesRemoved} pivot table${pivotTablesRemoved === 1 ? "" : "s"}: their caches contained rows from other groups, and a cache travels inside the workbook whether or not the pivot is opened. Rebuild the pivot in Excel from each output's own rows if it is required.`,
    );
  }

  return {
    operation: SPLIT_OPERATION,
    artifacts: outputPaths.map((output) => ({
      kind: "file",
      mediaType: WORKBOOK_MEDIA_TYPE,
      path: output,
    })),
    warnings,
    metrics: {
      calcChainEntriesRemoved: 0,
      formulaCellsBlankedForRemovedRows: 0,
      formulaCellsConverted: 0,
      formulaCellsWithoutCachedValues: 0,
      pivotTablesRemoved,
      groups: grouped.groups.length,
      inputFiles: 1,
      inputRows: table.rows.length,
      outputFiles: outputPaths.length,
      outputRows: grouped.groups.reduce(
        (total, group) => total + group.table.rows.length,
        0,
      ),
      rowsDeleted: 0,
      sheetsCopiedUnchanged: 0,
      sheetsFiltered: 1,
      skippedRows: grouped.skippedRows,
      valuesOnly: options.values === true ? 1 : 0,
    },
  };
}
