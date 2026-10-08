import { uncachedFormulaWarnings } from "./uncached-formulas.js";
export {
  uncachedFormulaHint,
  uncachedFormulaWarnings,
} from "./uncached-formulas.js";
import { randomUUID } from "node:crypto";
import { constants } from "node:fs";
import {
  copyFile,
  link,
  mkdtemp,
  open,
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
  FILES_ERRORS,
  isPathWithin,
  isSameFilesystemPath,
  openRandomAccessSource,
  pathExists,
  refuseInputOverwrite,
  type FileSource,
} from "@consultchimps/files";
import {
  validateColumnMapping,
  type ColumnMapping,
  type ColumnMappingSuggestion,
  type Table,
} from "@consultchimps/tabular";
import {
  unprotectWorkbookBytes,
  type UnprotectWorkbookMetric,
} from "./bytes.js";

import { XLSX_ERRORS } from "./errors.js";
import {
  planConsolidation,
  writeConsolidation,
  type ConsolidationSettings,
  type ConsolidationSink,
  type ConsolidationSource,
  type OpenedSource,
} from "./operations/consolidate/consolidate.js";
import { StreamedWorkbook } from "./operations/consolidate/reader.js";
import { openWorkbookBytes } from "./operations/sheet-grid.js";
import {
  readWorksheetReports,
  type WorksheetImportReport,
} from "./operations/worksheets.js";

// The worksheet report types beside the file surface that returns them, so a
// caller of readWorkbookWorksheets can name its result from the same entry
// point; the byte surface exports the same three.
export type { WorksheetImportReport } from "./operations/worksheets.js";
export type { WorksheetRegion, WorksheetTableReport } from "./shared.js";
import {
  describeStreamedWorkbook,
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
  preservedSplitExtensionOf,
  splitMediaType,
} from "./split/all-worksheet.js";
import { stagedFile } from "./split/staged-file.js";
import {
  openRegionPackage,
  resolveRegionSplit,
  uncachedForValues,
  writeRegionGroups,
  type ResolvedRegionSplit,
} from "./split/region-split.js";
import {
  type FullWorkbookSplitMetric,
  type FullWorkbookSplitSummary,
  type SplitOutputDetail,
  planFullWorkbookSplit,
  splitFullWorkbookByColumn,
} from "./workbook-column-split.js";
import {
  appendWorkbookSheets,
  buildTableWorkbookBytes,
  CONSOLIDATE_OPERATION,
  CONSOLIDATED_SHEET_NAME,
  assertSheetName,
  consolidationInputs,
  createMergeState,
  finishMergedWorkbook,
  INSPECT_OPERATION,
  isMacroWorkbookName,
  MACRO_WORKBOOK_MEDIA_TYPE,
  MAPPING_MEDIA_TYPE,
  MERGE_OPERATION,
  singleSourceUncachedFormulas,
  refuseMappingWithSuggestion,
  resolvePreserveWorkbook,
  serializeColumnMapping,
  skippedRowsWarning,
  splitOutputFileNames,
  SPLIT_OPERATION,
  unmappedColumnsWarning,
  WORKBOOK_EXTENSION,
  WORKBOOK_MEDIA_TYPE,
  workbookExcelTables,
  workbookNamedRanges,
  workbookWorksheetRecords,
  workbookWorksheetReports,
  yieldToEventLoop,
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
  WorksheetTableReport,
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
  UNDECIDED_UNPROTECT_STRUCTURES,
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
export type { UnprotectWorkbookMetric };
export interface UnprotectWorkbookOptions extends OperationControlOptions {
  input: string;
  output: string;
  overwrite?: boolean | undefined;
}
export async function unprotectWorkbook(
  options: UnprotectWorkbookOptions,
): Promise<OperationResult<UnprotectWorkbookMetric>> {
  const input = path.resolve(options.input);
  const output = path.resolve(options.output);
  refuseInputOverwrite(output, [input]);
  await ensureOutputAvailable(output, { overwrite: options.overwrite });
  const outcome = await unprotectWorkbookBytes({
    input: { name: path.basename(input), bytes: await readFile(input) },
    outputName: path.basename(output),
    onProgress: options.onProgress,
    signal: options.signal,
  });
  await ensureParentDirectory(output);
  await writeFile(output, outcome.outputs[0]!.bytes);
  return {
    ...outcome.result,
    // Keep the media type the bytes surface derived from the package, replacing
    // only the portable output name with the real destination path.
    artifacts: [{ ...outcome.result.artifacts[0]!, path: output }],
  };
}
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
 * Open a workbook on disk for the streaming reader, reporting both a missing
 * file and an unreadable workbook as the same stable error.
 */
async function openWorkbookFile(
  absolutePath: string,
): Promise<StreamedWorkbook> {
  return openWorkbookBytes(await readWorkbookBytes(absolutePath), {
    file: path.basename(absolutePath),
    source: absolutePath,
    details: { filePath: absolutePath },
  });
}

/**
 * Every selected worksheet of one file as the table reader saw it. Its cells
 * come from the streaming reader consolidation reads through, so the tables
 * `readWorkbookTables` returns hold the values a consolidation stacks.
 */
async function readWorkbookTableReports(
  filePath: string,
  options: ReadWorkbookOptions,
): Promise<WorksheetTableReport[]> {
  const absolutePath = path.resolve(filePath);
  return workbookWorksheetReports(
    await openWorkbookFile(absolutePath),
    path.basename(absolutePath),
    options,
  );
}

export async function readWorkbookTables(
  filePath: string,
  options: ReadWorkbookOptions = {},
): Promise<Table[]> {
  return consolidationInputs(await readWorkbookTableReports(filePath, options))
    .tables;
}

/**
 * Read every visible worksheet that holds data, with the rectangle each read
 * covered and the cells in it holding a formula the workbook carries no
 * calculated value for.
 *
 * `readWorkbookTables` is this list with the tables taken out of it. The file
 * twin of `readWorkbookWorksheetsBytes`: both adapt their input and hand the
 * same operation the same bytes, so neither can answer differently.
 */
export async function readWorkbookWorksheets(
  filePath: string,
  options: ReadWorkbookOptions = {},
): Promise<WorksheetImportReport[]> {
  const absolutePath = path.resolve(filePath);
  return readWorksheetReports(
    await readWorkbookBytes(absolutePath),
    path.basename(absolutePath),
    options,
  );
}

export async function readWorksheetRecords(
  filePath: string,
  options: ReadWorksheetRecordsOptions,
): Promise<WorksheetRecords> {
  const absolutePath = path.resolve(filePath);
  return workbookWorksheetRecords(
    await openWorkbookFile(absolutePath),
    options,
  );
}

export async function readWorkbookExcelTables(
  filePath: string,
  options: ReadWorkbookExcelTablesOptions = {},
): Promise<WorkbookExcelTable[]> {
  const absolutePath = path.resolve(filePath);
  return workbookExcelTables(
    await openWorkbookFile(absolutePath),
    path.basename(absolutePath),
    options,
  );
}

export async function readWorkbookNamedRanges(
  filePath: string,
  options: ReadWorkbookNamedRangesOptions = {},
): Promise<WorkbookNamedRange[]> {
  const absolutePath = path.resolve(filePath);
  return workbookNamedRanges(
    await openWorkbookFile(absolutePath),
    path.basename(absolutePath),
    options,
  );
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
  // Read in pieces through a file handle (ADR 0006), never whole.
  const opened = await openWorkbookSource(absolutePath);
  try {
    const workbook = await StreamedWorkbook.open(opened, {
      file: path.basename(absolutePath),
      source: absolutePath,
      details: { filePath: absolutePath },
    });
    return await describeStreamedWorkbook(
      workbook,
      path.basename(absolutePath),
      options,
    );
  } finally {
    await opened.close();
  }
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

/**
 * Open a workbook file for random access, reporting one that cannot be opened
 * with the read error every workbook reader raises.
 */
async function openWorkbookSource(absolutePath: string): Promise<OpenedSource> {
  try {
    return await openRandomAccessSource(absolutePath);
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${absolutePath}`,
      { cause: error, details: { filePath: absolutePath } },
    );
  }
}

/**
 * Write an output as it is produced, into a staging file beside the
 * destination that replaces it only once every byte is written, so a run that
 * fails or is cancelled partway leaves no partial workbook behind. The
 * destination is checked first, as `writeTable` checks it.
 */
/**
 * Move a finished staging file into place. Without overwrite, the destination
 * is linked rather than renamed onto, so a file another process created there
 * while the output was being written is refused, not replaced.
 */
function isAlreadyThere(error: unknown): boolean {
  return (error as NodeJS.ErrnoException | undefined)?.code === "EEXIST";
}

/** The refusal the destination check gives before writing, for a file that appeared since. */
async function refuseExistingOutput(destination: string): Promise<never> {
  await ensureOutputAvailable(destination);
  throw new ConsultChimpsError(
    FILES_ERRORS.FILES_OUTPUT_EXISTS,
    `Output already exists: ${destination}`,
    { details: { outputPath: destination } },
  );
}

/**
 * Remove the staging name once the output is published. The output already
 * stands complete, so a staging name that cannot be removed (a scanner holding
 * it, say) is left behind rather than failing a run that succeeded.
 */
async function removePublishedStaging(staging: string): Promise<void> {
  await rm(staging, { force: true }).catch(() => undefined);
}

async function publishStaging(
  staging: string,
  destination: string,
  overwrite: boolean,
): Promise<void> {
  if (overwrite) {
    await rename(staging, destination);
    return;
  }
  try {
    await link(staging, destination);
  } catch (linkError) {
    if (!isAlreadyThere(linkError)) {
      // A filesystem without hard links: copy, still refusing to replace.
      try {
        await copyFile(staging, destination, constants.COPYFILE_EXCL);
      } catch (copyError) {
        if (!isAlreadyThere(copyError)) throw copyError;
        await refuseExistingOutput(destination);
      }
      await removePublishedStaging(staging);
      return;
    }
    await refuseExistingOutput(destination);
  }
  await removePublishedStaging(staging);
}

async function writeStagedFile(
  outputPath: string,
  options: { overwrite?: boolean | undefined },
  produce: (sink: ConsolidationSink) => Promise<void>,
): Promise<string> {
  const absoluteOutput = await ensureParentDirectory(outputPath);
  await ensureOutputAvailable(absoluteOutput, { overwrite: options.overwrite });
  const staging = path.join(
    path.dirname(absoluteOutput),
    // A short fixed name, so a destination whose own name is near the
    // filesystem's limit still has room for its staging file.
    `.consultchimps-${randomUUID()}.partial`,
  );
  const handle = await open(staging, "wx");
  let pending: Uint8Array[] = [];
  const sink: ConsolidationSink = {
    write(chunk) {
      pending.push(chunk);
    },
    async flush() {
      if (pending.length === 0) return;
      const bytes = Buffer.concat(pending);
      pending = [];
      // A write may take less than it was given; the rest follows.
      for (let offset = 0; offset < bytes.length;) {
        const { bytesWritten } = await handle.write(
          bytes,
          offset,
          bytes.length - offset,
        );
        if (bytesWritten === 0) {
          // The run fails and the staging file is removed, rather than retrying forever.
          throw new Error(
            `The consolidated workbook could not be written to ${absoluteOutput}: the disk accepted no more data.`,
          );
        }
        offset += bytesWritten;
      }
    },
  };
  let closed = false;
  try {
    await produce(sink);
    await sink.flush();
    await handle.close();
    closed = true;
    await publishStaging(staging, absoluteOutput, options.overwrite === true);
  } catch (error) {
    if (!closed) await handle.close().catch(() => undefined);
    await rm(staging, { force: true }).catch(() => undefined);
    throw error;
  }
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

/**
 * Refuse a destination that already holds something other than a file, such as
 * a folder. `ensureOutputAvailable` cannot answer this on its own: with
 * overwrite enabled it returns before looking, and overwrite means replacing a
 * file, never replacing a folder with one.
 */
async function refuseNonFileDestination(outputPath: string): Promise<void> {
  try {
    if ((await stat(outputPath)).isFile()) {
      return;
    }
  } catch {
    // Nothing is there, which is the ordinary case and no problem at all.
    return;
  }
  throw new ConsultChimpsError(
    XLSX_ERRORS.XLSX_MAPPING_OUTPUT_NOT_FILE,
    `The drafted mapping cannot be written to ${outputPath}, because something that is not a file already stands there. Choose a filename for the draft, or move what is in the way.`,
    { details: { outputPath } },
  );
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
  // The mapping document is an input of this run like any workbook, so an
  // output aimed at it is refused before `overwrite` gets a chance to destroy
  // the very file the run was told to read.
  refuseInputOverwrite(
    absoluteOutput,
    options.mappingFile === undefined
      ? absoluteInputs
      : [...absoluteInputs, path.resolve(options.mappingFile)],
  );

  if (options.suggestMappingOutput === undefined) {
    return { absoluteInputs, absoluteOutput };
  }

  const absoluteSuggestOutput = path.resolve(options.suggestMappingOutput);
  refuseInputOverwrite(absoluteSuggestOutput, absoluteInputs);
  // Two destinations collide when they name one file - which on Windows and
  // the usual macOS volume includes a difference of case alone - and also when
  // one sits beneath the other, since no filesystem lets "report.xlsx" be a
  // file and a directory at once. Either way the second write would destroy or
  // fail over the first, so both are refused before anything is written.
  if (
    isSameFilesystemPath(absoluteSuggestOutput, absoluteOutput) ||
    isPathWithin(absoluteSuggestOutput, absoluteOutput) ||
    isPathWithin(absoluteOutput, absoluteSuggestOutput)
  ) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_MAPPING_SUGGEST_CONFLICT,
      `The drafted mapping (${absoluteSuggestOutput}) and the consolidated workbook (${absoluteOutput}) cannot share one destination, and neither can sit inside the other. Give the draft a location of its own.`,
      {
        details: {
          outputPath: absoluteOutput,
          problem: "suggestion_shares_output",
          suggestOutputPath: absoluteSuggestOutput,
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
    // A plan lists every file the run will read, and a mapping document is one
    // of them. `inputFiles` stays the workbook count, which is what it has
    // always meant and what the result reports.
    inputs:
      options.mappingFile === undefined
        ? absoluteInputs
        : [...absoluteInputs, path.resolve(options.mappingFile)],
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
    // Overwrite permits replacing a file, never turning a folder into one, and
    // it skips the availability check that would otherwise have noticed. The
    // check is made here so a folder standing at the draft's destination is
    // reported before the workbook is written rather than as an EISDIR after.
    await refuseNonFileDestination(absoluteSuggestOutput);
    await ensureOutputAvailable(absoluteSuggestOutput, {
      overwrite: options.overwrite,
    });
    // The draft's folder is made now rather than beside its write, because a
    // parent that cannot be a folder - a plain file already standing where one
    // is wanted - is only discovered by trying. Discovering it after the
    // workbook has been written would leave that workbook behind on a run that
    // reports failure.
    await ensureParentDirectory(absoluteSuggestOutput);
  }

  const sources: ConsolidationSource[] = absoluteInputs.map(
    (absoluteInput) => ({
      file: path.basename(absoluteInput),
      source: absoluteInput,
      details: { filePath: absoluteInput },
      open: () => openWorkbookSource(absoluteInput),
    }),
  );
  const settings: ConsolidationSettings = {
    headerRow: options.headerRow,
    includeHiddenSheets: options.includeHiddenSheets,
    sheets: options.sheets,
    addSourceColumns: options.addSourceColumns,
    normalizeHeaders: options.normalizeHeaders,
    mapping,
    suggestMapping: absoluteSuggestOutput !== undefined,
    signal: options.signal,
    onProgress: options.onProgress,
    outputContext: "files",
    yieldControl: false,
  };
  const plan = await planConsolidation(sources, settings);
  const { suggestion, unmappedColumns } = plan;
  throwIfAborted(options.signal, CONSOLIDATE_OPERATION);
  const sheetName = options.outputSheetName ?? CONSOLIDATED_SHEET_NAME;
  assertSheetName(sheetName);
  const output = await writeStagedFile(
    absoluteOutput,
    { overwrite: options.overwrite },
    (sink) => writeConsolidation(sources, plan, settings, sheetName, sink),
  );
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
    warnings: [
      ...(unmappedColumns.length > 0
        ? [unmappedColumnsWarning([...unmappedColumns])]
        : []),
      ...uncachedFormulaWarnings(plan.uncachedFormulas, "they came out blank"),
    ],
    metrics: {
      formulaCellsWithoutCachedValues: plan.uncachedFormulas.length,
      inputFiles: absoluteInputs.length,
      inputTables: plan.inputTables,
      outputColumns: plan.columns.length,
      outputRows: plan.rowCount,
      skippedSpacerColumns: plan.skippedSpacerColumns,
      skippedTitleRows: plan.skippedTitleRows,
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
      formulaCellsWithoutCachedValues: merged.uncachedFormulas.length,
    },
  };
}

interface ResolvedSplit {
  absoluteInput: string;
  absoluteOutputDirectory: string;
  existingOutputs: Set<string>;
  /** The media type every output of this split carries. */
  mediaType: string;
  outputPaths: string[];
  preserveWorkbook: boolean;
  resolved: ResolvedRegionSplit;
  /** The input, open for reading in pieces until the split is done with it. */
  source: FileSource;
}

async function resolveSplitWorkbookByColumn(
  options: SplitWorkbookByColumnOptions,
): Promise<ResolvedSplit> {
  const absoluteInput = path.resolve(options.input);
  const details = { inputPath: absoluteInput };
  let source: FileSource;

  try {
    source = await openRandomAccessSource(absoluteInput);
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${absoluteInput}`,
      { cause: error, details },
    );
  }

  try {
    const context = {
      details,
      file: path.basename(absoluteInput),
      label: absoluteInput,
    };
    const preserveWorkbook = resolvePreserveWorkbook(options);
    const resolved = await resolveRegionSplit(
      source,
      context,
      options,
      preserveWorkbook,
    );

    // A preserved split hands back the source package, so its outputs have
    // to be named and typed after that package; a rebuilding split writes a
    // fresh ordinary workbook and stays .xlsx.
    const extension = preserveWorkbook
      ? await preservedSplitExtensionOf(source, absoluteInput, {
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
      resolved.groups.map((group) => group.value),
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
      absoluteInput,
      absoluteOutputDirectory,
      existingOutputs,
      mediaType: preserveWorkbook
        ? splitMediaType(extension)
        : WORKBOOK_MEDIA_TYPE,
      outputPaths,
      preserveWorkbook,
      resolved,
      source,
    };
  } catch (error) {
    await source.close().catch(() => undefined);
    throw error;
  }
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
  const split = await resolveSplitWorkbookByColumn(options);
  await split.source.close();
  const { resolved } = split;
  const outputs: PlannedOutput[] = split.outputPaths.map((outputPath) => ({
    kind: "file",
    mediaType: split.mediaType,
    path: outputPath,
    exists: split.existingOutputs.has(outputPath),
  }));

  const warnings: string[] = [];
  if (resolved.skippedRows > 0) {
    warnings.push(skippedRowsWarning(resolved.skippedRows, resolved.column));
  }
  // What the plan read; a values-only preserved run also counts what its
  // conversion loses, which only the run can see.
  const uncached = singleSourceUncachedFormulas(
    resolved.uncachedFormulas,
    split.preserveWorkbook,
    options.values === true,
  );
  for (const warning of uncached.warnings) warnings.push(warning);
  const collisions = split.existingOutputs.size;
  if (collisions > 0 && options.overwrite !== true) {
    warnings.push(
      `${collisions} planned output file${
        collisions === 1 ? " already exists" : "s already exist"
      }; executing without overwrite will fail.`,
    );
  }

  return {
    operation: SPLIT_OPERATION,
    inputs: [split.absoluteInput],
    outputs,
    warnings,
    metrics: {
      calcChainEntriesRemoved: 0,
      formulaCellsBlankedForRemovedRows: 0,
      formulaCellsConverted: 0,
      formulaCellsWithoutCachedValues: uncached.count,
      pivotTablesRemoved: 0,
      groups: resolved.groups.length,
      inputFiles: 1,
      inputRows: resolved.inputRows,
      outputFiles: split.outputPaths.length,
      rowsDeleted: 0,
      sheetsCopiedUnchanged: 0,
      sheetsFiltered: 1,
      skippedRows: resolved.skippedRows,
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
 * asks for one of the narrower modes - a preserved Excel Table rewrite, or a
 * compact single-worksheet rebuild - which `split/region-split.ts` owns. Both
 * read the input in pieces and write each output as it is produced.
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
  const split = await resolveSplitWorkbookByColumn(options);
  try {
    return await splitResolvedRegion(options, split);
  } finally {
    await split.source.close().catch(() => undefined);
  }
}

async function splitResolvedRegion(
  options: SplitWorkbookByColumnOptions,
  split: ResolvedSplit,
): Promise<SplitWorkbookByColumnResult> {
  const {
    absoluteOutputDirectory,
    existingOutputs,
    mediaType,
    outputPaths,
    preserveWorkbook,
    resolved,
    source,
  } = split;

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
  let uncached: ReturnType<typeof singleSourceUncachedFormulas>;
  let pivotTablesRemoved: number;

  try {
    const splitPackage =
      preserveWorkbook && options.values === true
        ? await openRegionPackage(source)
        : undefined;
    uncached = singleSourceUncachedFormulas(
      resolved.uncachedFormulas,
      preserveWorkbook,
      options.values === true,
      splitPackage ? await uncachedForValues(splitPackage) : [],
    );
    pivotTablesRemoved = await writeRegionGroups(
      source,
      resolved,
      options.values === true,
      async (index) => {
        throwIfAborted(options.signal, SPLIT_OPERATION);
        const stagedOutput = path.join(
          transactionDirectory,
          `output-${String(index + 1).padStart(6, "0")}.xlsx`,
        );
        const file = await stagedFile(stagedOutput);
        return {
          write: (chunk) => file.write(chunk),
          abort: () => file.abort(),
          close: async () => {
            await file.close();
            stagedOutputs[index] = stagedOutput;
            options.onProgress?.({
              operation: SPLIT_OPERATION,
              stage: "staging-workbooks",
              completed: index + 1,
              total: resolved.groups.length,
              detail: path.basename(outputPaths[index] ?? stagedOutput),
            });
          },
        };
      },
      async () => {
        await yieldToEventLoop();
        throwIfAborted(options.signal, SPLIT_OPERATION);
      },
      splitPackage,
    );

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

  const warnings =
    resolved.skippedRows > 0
      ? [skippedRowsWarning(resolved.skippedRows, resolved.column)]
      : [];
  for (const warning of uncached.warnings) warnings.push(warning);
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
      formulaCellsWithoutCachedValues: uncached.count,
      pivotTablesRemoved,
      groups: resolved.groups.length,
      inputFiles: 1,
      inputRows: resolved.inputRows,
      outputFiles: outputPaths.length,
      outputRows: resolved.groups.reduce(
        (total, group) => total + group.rows,
        0,
      ),
      rowsDeleted: 0,
      sheetsCopiedUnchanged: 0,
      sheetsFiltered: 1,
      skippedRows: resolved.skippedRows,
      valuesOnly: options.values === true ? 1 : 0,
    },
  };
}
