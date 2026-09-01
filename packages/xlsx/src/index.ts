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
  type Artifact,
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
import {
  validateColumnMapping,
  type ColumnMapping,
  type ColumnMappingSuggestion,
  type Table,
} from "@consultchimps/tabular";
import type * as XLSX from "xlsx";

import { XLSX_ERRORS } from "./errors.js";
import {
  describeWorkbookModel,
  loadWorkbookModelForDescribe,
  MAX_COLUMN_SAMPLE_VALUES,
  type DescribeWorkbookMetric,
  type DescribeWorkbookOptions,
  type WorkbookColumnDescription,
  type WorkbookDescription,
  type WorkbookDescriptionOutcome,
  type WorkbookExcelTableDescription,
  type WorkbookNamedRangeDescription,
  type WorkbookSheetDescription,
  type WorksheetVisibility,
} from "./operations/describe.js";
import {
  preservedSplitExtension,
  splitMediaType,
} from "./split/all-worksheet.js";
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
  INSPECT_OPERATION,
  isMacroWorkbookName,
  MACRO_WORKBOOK_MEDIA_TYPE,
  MAPPING_MEDIA_TYPE,
  MERGE_OPERATION,
  parseExcelTableDefinitions,
  parseWorkbookBytes,
  preservedSplitTemplateBytes,
  refuseMappingWithSuggestion,
  resolveSplitSource,
  serializeColumnMapping,
  skippedRowsWarning,
  splitOutputFileNames,
  SPLIT_OPERATION,
  suggestMappingForTables,
  unmappedColumnsWarning,
  WORKBOOK_EXTENSION,
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
/**
 * The conformance contract: what this package promises to do to each tracked
 * workbook structure, per operation, with a recorded reason for every cell it
 * has not decided yet. Exported so that documentation and tooling can be
 * generated from the same table the corpus tests enforce, instead of restating
 * it in prose that drifts.
 *
 * A cell states one operation's ordinary outcome, not a prediction for a
 * particular run: it carries no options, no input ordering and no mode, so it
 * cannot express a behavior that depends on them. Two live examples, both
 * documented at /docs/tools/excel-preservation: `merge["vba-project"]` reads
 * `strip-warn`, yet a macro project survives when the first input is the only
 * one carrying it and the output is named `.xlsm`; and the `split` column
 * describes the default whole-workbook split, while `preserveWorkbook: false`,
 * `range` and `sheet` rebuild one worksheet of values and `table` refuses an
 * A1 formula it would have to move. Present a cell as what the operation
 * usually does and read the conditions beside it - never as this run's verdict.
 */
export {
  CONTRACT,
  OPERATIONS,
  TRACKED_STRUCTURES,
  UNDECIDED_DESCRIBE_STRUCTURES,
  UNDECIDED_MERGE_STRUCTURES,
  UNDECIDED_SPLIT_STRUCTURES,
  type ContractBehavior,
  type Operation as ContractOperation,
  type Structure as ContractStructure,
} from "./contract.js";
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
export { MAX_COLUMN_SAMPLE_VALUES };
export type {
  DescribeWorkbookMetric,
  DescribeWorkbookOptions,
  WorkbookColumnDescription,
  WorkbookDescription,
  WorkbookDescriptionOutcome,
  WorkbookExcelTableDescription,
  WorkbookNamedRangeDescription,
  WorkbookSheetDescription,
  WorksheetVisibility,
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
   * Path to a version 1 column mapping document. It is read, parsed, and
   * validated before any workbook is opened, and applied to each worksheet
   * table before the union. Cannot be combined with `suggestMappingOutput`.
   */
  mappingFile?: string | undefined;
  /**
   * Match columns whose headers differ only in case, spacing, or punctuation
   * (for example "Failed Checks" and "Failed_Checks") instead of requiring
   * the exact same header in every worksheet. A mapping matches by normalized
   * column key whatever this option says; the flag governs only the columns
   * the mapping did not claim.
   */
  normalizeHeaders?: boolean | undefined;
  outputSheetName?: string | undefined;
  overwrite?: boolean | undefined;
  /**
   * Where to write a drafted mapping built from the headers that were read.
   * The draft is written for review, never applied. Cannot be combined with
   * `mappingFile`.
   */
  suggestMappingOutput?: string | undefined;
  values?: boolean | undefined;
}

export interface ConsolidateWorkbooksResult extends OperationResult<ConsolidateWorkbooksMetric> {
  /**
   * The drafted mapping and the evidence behind it, present only when
   * `suggestMappingOutput` asked for one. Nothing applies it for the caller.
   */
  suggestion?: ColumnMappingSuggestion | undefined;
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

/**
 * Describe a workbook's structure without producing anything: worksheets with
 * their visibility, dimensions and header preview, Excel Tables, named ranges,
 * and a bounded sample of each column's values.
 *
 * The inspection reads the file and writes nothing, so the outcome carries the
 * description beside a structured result with no artifacts. `describeWorkbookBytes`
 * is its byte-level twin and produces a structurally identical description for
 * the same workbook.
 */
export async function describeWorkbook(
  filePath: string,
  options: DescribeWorkbookOptions = {},
): Promise<WorkbookDescriptionOutcome> {
  // Before any filesystem work, as the byte twin already does. An operation
  // handed an aborted signal must report the cancellation rather than the
  // first problem it happens to meet on the way - reading a missing file would
  // otherwise answer XLSX_READ_FAILED to a caller who had already stopped
  // caring, and a large valid workbook would be loaded in full for nothing.
  throwIfAborted(options.signal, INSPECT_OPERATION);
  const absolutePath = path.resolve(filePath);
  const workbook = await loadWorkbookModelForDescribe(
    await readWorkbookBytes(absolutePath),
    absolutePath,
    { filePath: absolutePath },
  );
  return describeWorkbookModel(workbook, path.basename(absolutePath), options);
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
  absoluteSuggestOutput?: string | undefined;
}

/**
 * Read, parse, and validate a mapping document from disk. Reading the file is
 * this package's job because `@consultchimps/tabular` never touches a
 * filesystem; the document's rules stay that package's to enforce, so an
 * ambiguous mapping still fails as `TABLE_MAPPING_INVALID`.
 */
async function readColumnMappingFile(
  mappingFilePath: string,
): Promise<ColumnMapping> {
  const absolutePath = path.resolve(mappingFilePath);
  const details = { mappingFile: absolutePath };
  let text: string;

  try {
    text = await readFile(absolutePath, "utf8");
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_MAPPING_FILE_UNREADABLE,
      `Could not read the column mapping file: ${absolutePath}`,
      { cause: error, details },
    );
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(text);
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_MAPPING_FILE_INVALID,
      `The column mapping file is not valid JSON: ${absolutePath}`,
      { cause: error, details },
    );
  }

  return validateColumnMapping(parsed);
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
  if (
    options.mappingFile !== undefined &&
    options.suggestMappingOutput !== undefined
  ) {
    refuseMappingWithSuggestion();
  }

  const absoluteInputs = options.inputs.map((inputPath) =>
    path.resolve(inputPath),
  );
  const absoluteOutput = path.resolve(options.output);
  refuseInputOverwrite(absoluteOutput, absoluteInputs);

  if (options.suggestMappingOutput === undefined) {
    return { absoluteInputs, absoluteOutput };
  }

  const absoluteSuggestOutput = path.resolve(options.suggestMappingOutput);
  refuseInputOverwrite(absoluteSuggestOutput, absoluteInputs);
  if (absoluteSuggestOutput === absoluteOutput) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_MAPPING_SUGGEST_CONFLICT,
      `The drafted mapping and the consolidated workbook cannot share one destination: ${absoluteOutput}. Give the draft a filename of its own.`,
      {
        details: {
          outputPath: absoluteOutput,
          problem: "suggestion_shares_output",
        },
      },
    );
  }
  return { absoluteInputs, absoluteOutput, absoluteSuggestOutput };
}

export async function planConsolidateWorkbooks(
  options: ConsolidateWorkbooksOptions,
): Promise<OperationPlan<ConsolidateWorkbooksPlanMetric>> {
  const { absoluteInputs, absoluteOutput, absoluteSuggestOutput } =
    resolveConsolidateWorkbooks(options);

  // A plan promises the run's destinations, so an unusable mapping has to fail
  // here too rather than surviving until the workbooks are read.
  if (options.mappingFile !== undefined) {
    await readColumnMappingFile(options.mappingFile);
  }

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
  const outputs: PlannedOutput[] = [
    {
      kind: "file",
      mediaType: WORKBOOK_MEDIA_TYPE,
      path: absoluteOutput,
      exists,
    },
  ];
  const warnings =
    exists && options.overwrite !== true
      ? [
          "The planned output workbook already exists; executing without overwrite will fail.",
        ]
      : [];

  if (absoluteSuggestOutput !== undefined) {
    const suggestExists = await pathExists(absoluteSuggestOutput);
    outputs.push({
      kind: "file",
      mediaType: MAPPING_MEDIA_TYPE,
      path: absoluteSuggestOutput,
      exists: suggestExists,
    });
    if (suggestExists && options.overwrite !== true) {
      warnings.push(
        "The planned mapping draft already exists; executing without overwrite will fail.",
      );
    }
  }

  return {
    operation: CONSOLIDATE_OPERATION,
    inputs: absoluteInputs,
    outputs,
    warnings,
    metrics: {
      inputFiles: absoluteInputs.length,
      outputFiles: outputs.length,
    },
  };
}

export async function consolidateWorkbooks(
  options: ConsolidateWorkbooksOptions,
): Promise<ConsolidateWorkbooksResult> {
  throwIfAborted(options.signal, CONSOLIDATE_OPERATION);
  const { absoluteInputs, absoluteOutput, absoluteSuggestOutput } =
    resolveConsolidateWorkbooks(options);

  // Both the mapping and the draft's destination are settled before a single
  // workbook is opened, so an unusable mapping or an occupied destination
  // costs nothing and leaves nothing behind.
  const mapping =
    options.mappingFile === undefined
      ? undefined
      : await readColumnMappingFile(options.mappingFile);
  if (absoluteSuggestOutput !== undefined) {
    await ensureOutputAvailable(absoluteSuggestOutput, {
      overwrite: options.overwrite,
    });
  }

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

  const { table, unmappedColumns } = consolidateTables(tables, {
    ...options,
    mapping,
  });
  const suggestion =
    absoluteSuggestOutput === undefined
      ? undefined
      : suggestMappingForTables(tables);
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

  const artifacts: Artifact[] = [
    {
      kind: "file",
      mediaType: WORKBOOK_MEDIA_TYPE,
      path: output,
    },
  ];
  if (absoluteSuggestOutput !== undefined && suggestion) {
    await ensureParentDirectory(absoluteSuggestOutput);
    await ensureOutputAvailable(absoluteSuggestOutput, {
      overwrite: options.overwrite,
    });
    await writeFile(
      absoluteSuggestOutput,
      serializeColumnMapping(suggestion.mapping),
      "utf8",
    );
    artifacts.push({
      kind: "file",
      mediaType: MAPPING_MEDIA_TYPE,
      path: absoluteSuggestOutput,
    });
    options.onProgress?.({
      operation: CONSOLIDATE_OPERATION,
      stage: "writing-mapping-draft",
      completed: 1,
      total: 1,
      detail: path.basename(absoluteSuggestOutput),
    });
  }

  const result: ConsolidateWorkbooksResult = {
    operation: CONSOLIDATE_OPERATION,
    artifacts,
    warnings:
      unmappedColumns.length > 0
        ? [unmappedColumnsWarning(unmappedColumns)]
        : [],
    metrics: {
      inputFiles: absoluteInputs.length,
      inputTables: tables.length,
      outputColumns: table.columns.length,
      outputRows: table.rows.length,
      suggestedColumns: suggestion?.mapping.columns.length ?? 0,
      unmappedColumns: unmappedColumns.length,
    },
  };
  if (suggestion) {
    result.suggestion = suggestion;
  }
  return result;
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
  /** The media type every output of this split carries. */
  mediaType: string;
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

  // A preserved split hands back the source package, so its outputs have to be
  // named and typed after that package; a rebuilding split writes a fresh
  // ordinary workbook and stays .xlsx.
  const extension = resolved.preserveWorkbook
    ? await preservedSplitExtension(workbookBytes, absoluteInput, {
        details,
        label: absoluteInput,
      })
    : WORKBOOK_EXTENSION;

  const absoluteOutputDirectory = path.resolve(options.outputDirectory);
  const filenamePrefix = safeNameFragment(
    options.filenamePrefix ?? path.parse(absoluteInput).name,
    "split",
  );
  const outputPaths = splitOutputFileNames(
    filenamePrefix,
    resolved.grouped.groups.map((group) => group.value),
    extension,
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
    mediaType: resolved.preserveWorkbook
      ? splitMediaType(extension)
      : WORKBOOK_MEDIA_TYPE,
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
    mediaType: resolved.mediaType,
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
    mediaType,
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
      mediaType,
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
