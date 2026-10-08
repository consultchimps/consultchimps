/**
 * Byte-level workbook operations for environments without a filesystem, such
 * as browsers. Inputs and outputs are in-memory bytes; artifact paths in the
 * structured results carry portable output names. This module must stay free
 * of node:fs and node:path imports.
 */
import { uncachedFormulaWarnings } from "./uncached-formulas.js";
export {
  uncachedFormulaHint,
  uncachedFormulaWarnings,
} from "./uncached-formulas.js";
import {
  ConsultChimpsError,
  throwIfAborted,
  type Artifact,
  type ByteArtifact,
  type ByteOperationOutcome,
  type OperationControlOptions,
  type OperationPlan,
  type OperationResult,
  type RandomAccessSource,
} from "@consultchimps/core";

import {
  validateColumnMapping,
  type ColumnMapping,
  type ColumnMappingSuggestion,
  type Table,
} from "@consultchimps/tabular";

import { XLSX_ERRORS } from "./errors.js";
import {
  MACRO_WORKBOOK_MAIN_CONTENT_TYPE,
  WORKBOOK_MAIN_PART,
  WorkbookPackage,
} from "./package/index.js";
import { bytesSource } from "./package/index.js";
import {
  planConsolidation,
  writeConsolidation,
  type ConsolidationSettings,
  type ConsolidationSource,
} from "./operations/consolidate/consolidate.js";
import { StreamedWorkbook } from "./operations/consolidate/reader.js";
import { openWorkbookBytes } from "./operations/sheet-grid.js";
import {
  openRegionPackage,
  resolveRegionSplit,
  uncachedForValues,
  writeRegionGroups,
  type ResolvedRegionSplit,
} from "./split/region-split.js";
import {
  readWorksheetReports,
  type WorksheetImportReport,
} from "./operations/worksheets.js";
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
  analyzeAllWorksheetSplit,
  plannedAllWorksheetSplitMetrics,
  plannedAllWorksheetSplitWarnings,
  runAllWorksheetSplit,
  splitMediaType,
  workbookExtensionOf,
  type AllWorksheetSplitAnalysis,
  type AllWorksheetSplitSelection,
  type AllWorksheetSplitSummary,
  type SplitOutputDetail,
  type SplitSourceIdentity,
  preservedSplitExtensionOf,
} from "./split/all-worksheet.js";
import {
  appendWorkbookSheets,
  CONSOLIDATE_OPERATION,
  CONSOLIDATED_SHEET_NAME,
  assertSheetName,
  consolidationInputs,
  createMergeState,
  finishMergedWorkbook,
  INSPECT_OPERATION,
  isMacroWorkbookName,
  MACRO_WORKBOOK_EXTENSION,
  MACRO_WORKBOOK_MEDIA_TYPE,
  MAPPING_MEDIA_TYPE,
  MERGE_OPERATION,
  singleSourceUncachedFormulas,
  refuseMappingWithSuggestion,
  resolvePreserveWorkbook,
  safeNameFragment,
  serializeColumnMapping,
  skippedRowsWarning,
  splitOutputFileNames,
  SPLIT_OPERATION,
  SUGGESTED_MAPPING_FILE_NAME,
  unmappedColumnsWarning,
  WORKBOOK_EXTENSION,
  WORKBOOK_MEDIA_TYPE,
  withoutWorkbookExtension,
  workbookExcelTables,
  workbookNamedRanges,
  workbookWorksheetRecords,
  workbookWorksheetReports,
  yieldToEventLoop,
  type ConsolidateWorkbooksMetric,
  type MergeWorkbooksMetric,
  type ReadWorkbookExcelTablesOptions,
  type ReadWorkbookNamedRangesOptions,
  type ReadWorkbookOptions,
  type SplitWorkbookByColumnMetric,
  type SplitWorkbookByColumnPlanMetric,
  type WorkbookExcelTable,
  type WorkbookNamedRange,
  type WorksheetRecords,
} from "./shared.js";

export const UNPROTECT_OPERATION = "sheets.unprotect";
export type UnprotectWorkbookMetric =
  "sheetProtectionsRemoved" | "workbookProtectionsRemoved";
export interface UnprotectWorkbookBytesOptions extends OperationControlOptions {
  input: WorkbookInputBytes;
  outputName?: string;
}

/**
 * Refuse an output name whose extension contradicts the package's declared
 * workbook content type, before anything is written.
 *
 * Unprotect preserves the source package and never adds or removes a VBA
 * project, so its outputs are whatever the input already was: a macro-enabled
 * package or an ordinary one. Naming an ordinary package `.xlsm`, or a macro
 * package `.xlsx`, would write a file whose contents and name disagree, which
 * Excel opens with a corruption warning. The default output name equals the
 * input name, so this catches an explicit rename and a mislabelled input alike.
 * A name ending in neither extension is left to the caller: this refuses only
 * the two forms that make a definite, contradictory claim.
 */
function refuseUnprotectPackageTypeMismatch(
  declaresMacroWorkbook: boolean,
  outputName: string,
): void {
  const nameExtension = /\.xlsm$/iu.test(outputName)
    ? ".xlsm"
    : /\.xlsx$/iu.test(outputName)
      ? ".xlsx"
      : undefined;
  if (nameExtension === undefined) {
    return;
  }
  const declaredExtension = declaresMacroWorkbook ? ".xlsm" : ".xlsx";
  if (declaredExtension === nameExtension) {
    return;
  }
  throw new ConsultChimpsError(
    XLSX_ERRORS.XLSX_UNPROTECT_PACKAGE_TYPE_MISMATCH,
    declaresMacroWorkbook
      ? `The output name "${outputName}" ends in ".xlsx" but the workbook is macro-enabled. Name the output with an .xlsm extension, or save the workbook as an ordinary .xlsx workbook in Excel first. Writing it as it is would produce a file whose contents and name disagree.`
      : `The output name "${outputName}" ends in ".xlsm" but the workbook is an ordinary Excel workbook with no macro project. Name the output with an .xlsx extension, or save the workbook as a macro-enabled workbook in Excel first. Writing it as it is would produce a file whose contents and name disagree.`,
    {
      details: {
        declaredExtension,
        macroEnabled: declaresMacroWorkbook,
        nameExtension,
        outputName,
      },
    },
  );
}

export async function unprotectWorkbookBytes(
  options: UnprotectWorkbookBytesOptions,
): Promise<ByteOperationOutcome<UnprotectWorkbookMetric>> {
  throwIfAborted(options.signal, UNPROTECT_OPERATION, "memory");
  const inputName = options.input.name.trim();
  if (!/\.(?:xlsx|xlsm)$/iu.test(inputName))
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_UNPROTECT_UNSUPPORTED_FILE,
      "Excel Unprotect accepts only .xlsx and .xlsm files.",
    );
  let archive: WorkbookPackage;
  try {
    archive = await WorkbookPackage.load(options.input.bytes, {
      sourceLabel: options.input.name,
    });
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_UNPROTECT_UNSUPPORTED_FILE,
      "This file is not a readable OOXML Excel workbook. Encrypted or password-required Office files are not supported.",
      { cause: error },
    );
  }
  const workbook = archive.part("xl/workbook.xml");
  if (!workbook || !archive.part("[Content_Types].xml"))
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_UNPROTECT_UNSUPPORTED_FILE,
      "This file is not a valid OOXML Excel workbook. Encrypted or password-required Office files are not supported.",
    );
  // The media type describes the bytes, so it follows the package's own
  // workbook content type rather than the output name: unprotect never adds or
  // removes a VBA project, so the bytes stay whatever the input already was.
  // The output name has to agree with that type rather than override it, which
  // is what `refuseUnprotectPackageTypeMismatch` enforces below.
  // contentTypeOverride parses [Content_Types].xml on demand, so a malformed or
  // DOCTYPE-bearing declaration must surface as the operation's own read failure
  // here rather than leaking a raw parser error, the same reason the split wraps it.
  let declaresMacroWorkbook: boolean;
  try {
    declaresMacroWorkbook =
      archive.contentTypeOverride(WORKBOOK_MAIN_PART)?.trim() ===
      MACRO_WORKBOOK_MAIN_CONTENT_TYPE;
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_UNPROTECT_UNSUPPORTED_FILE,
      "This file is not a valid OOXML Excel workbook. Encrypted or password-required Office files are not supported.",
      { cause: error },
    );
  }
  const outputName = options.outputName ?? inputName;
  // The name is checked against the package before any part is rewritten, so a
  // contradicting name costs nothing and leaves the input untouched.
  refuseUnprotectPackageTypeMismatch(declaresMacroWorkbook, outputName);
  // The XML edits are the package layer's to make: unprotect asks for the
  // protection elements to be dropped by name and never touches the markup
  // itself. Protection elements in OOXML are always empty attribute-only
  // elements, and the seam matches a namespace prefix such as <x:sheetProtection/>.
  const workbookProtectionsRemoved = archive.removeEmptyElements(
    "xl/workbook.xml",
    "workbookProtection",
  );
  let sheetProtectionsRemoved = 0;
  for (const partPath of archive.partPaths()) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/iu.test(partPath)) continue;
    sheetProtectionsRemoved += archive.removeEmptyElements(
      partPath,
      "sheetProtection",
    );
  }
  const mediaType = declaresMacroWorkbook
    ? MACRO_WORKBOOK_MEDIA_TYPE
    : WORKBOOK_MEDIA_TYPE;
  options.onProgress?.({
    operation: UNPROTECT_OPERATION,
    stage: "writing-output",
    completed: 1,
    total: 1,
    detail: outputName,
  });
  return {
    result: {
      operation: UNPROTECT_OPERATION,
      artifacts: [{ kind: "file", mediaType, path: outputName }],
      warnings: [],
      metrics: {
        sheetProtectionsRemoved,
        workbookProtectionsRemoved,
      },
    },
    outputs: [
      {
        name: outputName,
        bytes: await archive.save(),
        mediaType,
      },
    ],
  };
}

export interface WorkbookInputBytes {
  name: string;
  bytes: Uint8Array;
}

export interface SplitWorkbookBytesOptions extends OperationControlOptions {
  input: WorkbookInputBytes;
  column: string;
  filenamePrefix?: string | undefined;
  headerRow?: number | undefined;
  /** Only applies to a table, range, or worksheet selection. */
  includeBlank?: boolean | undefined;
  /** Only applies to a table, range, or worksheet selection. */
  includeHiddenSheets?: boolean | undefined;
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

export interface SplitWorkbookBytesResult extends OperationResult<SplitWorkbookByColumnMetric> {
  /** Per-output, per-worksheet filtering details for all-worksheet splits. */
  outputs?: SplitOutputDetail[] | undefined;
  /** Input, option, and worksheet summary for all-worksheet splits. */
  summary?: AllWorksheetSplitSummary | undefined;
}

export interface SplitWorkbookBytesOutcome extends ByteOperationOutcome<SplitWorkbookByColumnMetric> {
  result: SplitWorkbookBytesResult;
}

export interface ConsolidateWorkbooksBytesOptions
  extends ReadWorkbookOptions, OperationControlOptions {
  inputs: WorkbookInputBytes[];
  addSourceColumns?: boolean | undefined;
  /**
   * A parsed version 1 column mapping, applied to each worksheet table before
   * the union. This surface has no filesystem, so the caller parses the
   * document; it is validated here before any workbook is read. Cannot be
   * combined with `suggestMapping`.
   */
  mapping?: ColumnMapping | undefined;
  /**
   * Match columns whose headers differ only in case, spacing, or punctuation
   * (for example "Failed Checks" and "Failed_Checks") instead of requiring
   * the exact same header in every worksheet. A mapping matches by normalized
   * column key whatever this option says; the flag governs only the columns
   * the mapping did not claim.
   */
  normalizeHeaders?: boolean | undefined;
  outputName?: string | undefined;
  outputSheetName?: string | undefined;
  /**
   * Draft a mapping from the headers that were read and offer it beside the
   * consolidated workbook, as `mapping-draft.json`. The draft is for review,
   * never applied. Cannot be combined with `mapping`.
   */
  suggestMapping?: boolean | undefined;
}

export interface ConsolidateWorkbooksBytesResult extends OperationResult<ConsolidateWorkbooksMetric> {
  /**
   * The drafted mapping and the evidence behind it, present only when
   * `suggestMapping` asked for one. Nothing applies it for the caller.
   */
  suggestion?: ColumnMappingSuggestion | undefined;
}

export interface ConsolidateWorkbooksBytesOutcome extends ByteOperationOutcome<ConsolidateWorkbooksMetric> {
  result: ConsolidateWorkbooksBytesResult;
}

/** Chunks joined into one buffer. */
function concatenate(chunks: readonly Uint8Array[]): Uint8Array {
  const bytes = new Uint8Array(
    chunks.reduce((total, chunk) => total + chunk.length, 0),
  );
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.length;
  }
  return bytes;
}

/**
 * A workbook read in pieces through `Blob.slice`, so a browser `File` is never
 * copied whole into memory. A file changed on disk since it was chosen fails
 * its next read.
 */
export function blobSource(name: string, blob: Blob): RandomAccessSource {
  return {
    name,
    size: blob.size,
    // Cancellation is left to the operation reading, which reports it in its
    // own terms between reads.
    readAt: async (offset, length) => {
      if (
        !Number.isSafeInteger(offset) ||
        !Number.isSafeInteger(length) ||
        offset < 0 ||
        length < 0 ||
        offset + length > blob.size
      ) {
        throw new RangeError("The requested range is outside the workbook.");
      }
      return new Uint8Array(
        await blob.slice(offset, offset + length).arrayBuffer(),
      );
    },
  };
}

/** Where a streamed output's bytes go, in order. */
export interface ByteSink {
  /** Take the next chunk. The writer does not touch it again. */
  write(chunk: Uint8Array): void;
  /** Called between reads; a sink that buffers writes out what it holds. */
  flush(): Promise<void>;
  /**
   * Called once when the operation fails or is cancelled after it may have
   * written: discard everything written, so no partial output is left.
   */
  abort(): Promise<void>;
}

export interface ConsolidateWorkbookSourcesOptions extends Omit<
  ConsolidateWorkbooksBytesOptions,
  "inputs"
> {
  /** Each workbook, read in pieces. Its `name` is what `_source_file` records. */
  inputs: RandomAccessSource[];
  /** Receives the consolidated workbook as it is written. */
  output: ByteSink;
}

export interface ConsolidateWorkbookSourcesOutcome {
  result: ConsolidateWorkbooksBytesResult;
  /** The name of the workbook written to `output`. */
  outputName: string;
  /** The mapping draft, present only when `suggestMapping` asked for one. */
  mappingDraft?: ByteArtifact | undefined;
}

export interface MergeWorkbooksBytesOptions extends OperationControlOptions {
  inputs: WorkbookInputBytes[];
  includeSheetIndex?: boolean | undefined;
  outputName?: string | undefined;
  values?: boolean | undefined;
}

export interface ReadWorksheetRecordsBytesOptions {
  headerRow?: number | undefined;
  worksheet?: string | undefined;
}

/**
 * Whether this split keeps the whole workbook and filters every worksheet that
 * carries the column, rather than rebuilding one selected source.
 *
 * The rule is the file surface's rule, character for character: naming a
 * table, a range, or a worksheet asks for that narrower source, and
 * `preserveWorkbook: false` asks for a compact rebuild. Anything else gets the
 * workbook-preserving split, so the same options mean the same thing whether a
 * caller has a filesystem or only bytes.
 */
function isAllWorksheetSplit(options: SplitSelectionShape): boolean {
  return (
    !options.table &&
    !options.range &&
    !options.sheet &&
    options.preserveWorkbook !== false
  );
}

type SplitSelectionShape = Omit<SplitWorkbookBytesOptions, "input">;

/**
 * The name every output of this split is built from.
 *
 * A byte split's outputs are downloads that land wherever the caller puts
 * them, with no chosen directory to tell one job's results from another's, so
 * they keep the source-derived prefix this surface has always used. The file
 * surface, which writes into a directory the caller named, uses the group
 * value alone.
 */
function splitFilenamePrefix(
  options: SplitSelectionShape,
  inputName: string,
): string {
  return safeNameFragment(
    options.filenamePrefix ?? withoutWorkbookExtension(inputName),
    "split",
  );
}

interface ResolvedAllWorksheetSplitBytes {
  analysis: AllWorksheetSplitAnalysis;
  identity: SplitSourceIdentity;
  outputNames: string[];
  selection: AllWorksheetSplitSelection;
}

async function resolveAllWorksheetSplitSource(
  input: RandomAccessSource,
  options: SplitSelectionShape,
): Promise<ResolvedAllWorksheetSplitBytes> {
  const identity: SplitSourceIdentity = {
    details: { source: input.name },
    label: input.name,
  };
  const extension = workbookExtensionOf(input.name, identity);
  const selection: AllWorksheetSplitSelection = {
    column: options.column,
    headerRow: options.headerRow,
    strict: options.strict,
    values: options.values,
  };
  const analysis = await analyzeAllWorksheetSplit(
    input,
    extension,
    selection,
    identity,
  );

  return {
    analysis,
    identity,
    outputNames: splitOutputFileNames(
      splitFilenamePrefix(options, input.name),
      analysis.groups.map((group) => group.display),
      extension,
    ),
    selection,
  };
}

interface ResolvedRegionSplitSource {
  resolved: ResolvedRegionSplit;
  preserveWorkbook: boolean;
  mediaType: string;
  outputNames: string[];
}

async function resolveRegionSplitSource(
  input: RandomAccessSource,
  options: SplitSelectionShape,
): Promise<ResolvedRegionSplitSource> {
  const context = {
    details: { source: input.name },
    file: input.name,
    label: input.name,
  };
  const preserveWorkbook = resolvePreserveWorkbook(options);
  const resolved = await resolveRegionSplit(
    input,
    context,
    options,
    preserveWorkbook,
  );
  const extension = preserveWorkbook
    ? await preservedSplitExtensionOf(input, input.name, context)
    : WORKBOOK_EXTENSION;
  return {
    resolved,
    preserveWorkbook,
    mediaType: preserveWorkbook
      ? splitMediaType(extension)
      : WORKBOOK_MEDIA_TYPE,
    outputNames: splitOutputFileNames(
      splitFilenamePrefix(options, input.name),
      resolved.groups.map((group) => group.value),
      extension,
    ),
  };
}

export async function planSplitWorkbookBytes(
  options: SplitWorkbookBytesOptions,
): Promise<OperationPlan<SplitWorkbookByColumnPlanMetric>> {
  const { input, ...rest } = options;
  return planSplitWorkbookSource({
    ...rest,
    input: bytesSource(input.name, input.bytes),
  });
}

export interface PlanSplitWorkbookSourceOptions extends SplitSelectionShape {
  /** The workbook, read in pieces. Its `name` names the outputs. */
  input: RandomAccessSource;
}

/**
 * `planSplitWorkbookBytes` over a workbook read in pieces, such as a browser
 * `File` through `blobSource`, so it is never held whole.
 */
export async function planSplitWorkbookSource(
  options: PlanSplitWorkbookSourceOptions,
): Promise<OperationPlan<SplitWorkbookByColumnPlanMetric>> {
  const { input } = options;
  if (isAllWorksheetSplit(options)) {
    const resolved = await resolveAllWorksheetSplitSource(input, options);
    return {
      operation: SPLIT_OPERATION,
      inputs: [input.name],
      outputs: resolved.outputNames.map((name) => ({
        kind: "file",
        mediaType: resolved.analysis.mediaType,
        path: name,
        exists: false,
      })),
      warnings: [
        ...(resolved.analysis.skippedRows > 0
          ? [
              `Skipped ${resolved.analysis.skippedRows} row${resolved.analysis.skippedRows === 1 ? "" : "s"} with blank values in "${options.column}"; no blank-value workbook was created.`,
            ]
          : []),
        ...plannedAllWorksheetSplitWarnings(resolved.analysis),
      ],
      metrics: plannedAllWorksheetSplitMetrics(
        resolved.analysis,
        resolved.selection,
        resolved.outputNames.length,
      ),
    };
  }

  const { resolved, preserveWorkbook, mediaType, outputNames } =
    await resolveRegionSplitSource(input, options);
  const uncached = singleSourceUncachedFormulas(
    resolved.uncachedFormulas,
    preserveWorkbook,
    options.values === true,
  );
  const warnings = [
    ...(resolved.skippedRows > 0
      ? [skippedRowsWarning(resolved.skippedRows, resolved.column)]
      : []),
    ...uncached.warnings,
  ];

  return {
    operation: SPLIT_OPERATION,
    inputs: [input.name],
    outputs: outputNames.map((name) => ({
      kind: "file",
      mediaType,
      path: name,
      exists: false,
    })),
    warnings,
    metrics: {
      calcChainEntriesRemoved: 0,
      formulaCellsBlankedForRemovedRows: 0,
      formulaCellsConverted: 0,
      formulaCellsWithoutCachedValues: uncached.count,
      groups: resolved.groups.length,
      inputFiles: 1,
      inputRows: resolved.inputRows,
      outputFiles: outputNames.length,
      pivotTablesRemoved: 0,
      rowsDeleted: 0,
      sheetsCopiedUnchanged: 0,
      sheetsFiltered: 1,
      skippedRows: resolved.skippedRows,
      valuesOnly: options.values === true ? 1 : 0,
    },
  };
}

export async function splitWorkbookBytes(
  options: SplitWorkbookBytesOptions,
): Promise<SplitWorkbookBytesOutcome> {
  const { input, ...rest } = options;
  const outputs: ByteArtifact[] = [];
  const { result } = await splitWorkbookSource({
    ...rest,
    input: bytesSource(input.name, input.bytes),
    output: (name, mediaType) => {
      const chunks: Uint8Array[] = [];
      return {
        write: (chunk) => {
          chunks.push(chunk);
        },
        flush: () => {
          outputs.push({ name, bytes: concatenate(chunks), mediaType });
          return Promise.resolve();
        },
        abort: () => {
          chunks.length = 0;
          return Promise.resolve();
        },
      };
    },
  });
  return { result, outputs };
}

export interface SplitWorkbookSourceOptions extends SplitSelectionShape {
  /** The workbook, read in pieces. Its `name` names the outputs. */
  input: RandomAccessSource;
  /**
   * Open where one output goes. Each receives its bytes in order through
   * `write`, then `flush` once it is complete; `abort` when the split fails
   * or is cancelled after it opened.
   */
  output: (name: string, mediaType: string) => ByteSink | Promise<ByteSink>;
}

export interface SplitWorkbookSourceOutcome {
  result: SplitWorkbookBytesResult;
}

/**
 * `splitWorkbookBytes` over a workbook read in pieces and outputs written as
 * they are produced: a browser reads its `File` through `blobSource` and
 * writes each output to a file of its own. The outputs hold exactly the bytes
 * `splitWorkbookBytes` returns.
 */
export async function splitWorkbookSource(
  options: SplitWorkbookSourceOptions,
): Promise<SplitWorkbookSourceOutcome> {
  throwIfAborted(options.signal, SPLIT_OPERATION, "memory");
  const opened: ByteSink[] = [];
  const open = async (name: string, mediaType: string): Promise<ByteSink> => {
    const sink = await options.output(name, mediaType);
    opened.push(sink);
    return sink;
  };
  try {
    return isAllWorksheetSplit(options)
      ? await splitAllWorksheetsSource(options, open)
      : await splitRegionSource(options, open);
  } catch (error) {
    for (const sink of opened) await sink.abort().catch(() => undefined);
    throw error;
  }
}

async function splitRegionSource(
  options: SplitWorkbookSourceOptions,
  open: (name: string, mediaType: string) => Promise<ByteSink>,
): Promise<SplitWorkbookSourceOutcome> {
  const { input } = options;
  const { resolved, preserveWorkbook, mediaType, outputNames } =
    await resolveRegionSplitSource(input, options);
  const splitPackage =
    preserveWorkbook && options.values === true
      ? await openRegionPackage(input)
      : undefined;
  const conversionLosses = splitPackage
    ? await uncachedForValues(splitPackage)
    : [];
  const uncached = singleSourceUncachedFormulas(
    resolved.uncachedFormulas,
    preserveWorkbook,
    options.values === true,
    conversionLosses,
  );
  const between = async (): Promise<void> => {
    await yieldToEventLoop();
    throwIfAborted(options.signal, SPLIT_OPERATION, "memory");
  };
  const pivotTablesRemoved = await writeRegionGroups(
    input,
    resolved,
    options.values === true,
    async (index) => {
      throwIfAborted(options.signal, SPLIT_OPERATION, "memory");
      const name = outputNames[index]!;
      const sink = await open(name, mediaType);
      return {
        write: (chunk) => sink.write(chunk),
        close: async () => {
          await sink.flush();
          options.onProgress?.({
            operation: SPLIT_OPERATION,
            stage: "building-workbooks",
            completed: index + 1,
            total: resolved.groups.length,
            detail: name,
          });
        },
      };
    },
    between,
    splitPackage,
  );
  throwIfAborted(options.signal, SPLIT_OPERATION, "memory");

  return {
    result: {
      operation: SPLIT_OPERATION,
      artifacts: outputNames.map((name) => ({
        kind: "file",
        mediaType,
        path: name,
      })),
      warnings: [
        ...(resolved.skippedRows > 0
          ? [skippedRowsWarning(resolved.skippedRows, resolved.column)]
          : []),
        ...uncached.warnings,
        ...(pivotTablesRemoved > 0
          ? [
              `Removed ${pivotTablesRemoved} pivot table${pivotTablesRemoved === 1 ? "" : "s"}: their caches contained rows from other groups, and a cache travels inside the workbook whether or not the pivot is opened. Rebuild the pivot in Excel from each output's own rows if it is required.`,
            ]
          : []),
      ],
      metrics: {
        calcChainEntriesRemoved: 0,
        formulaCellsBlankedForRemovedRows: 0,
        formulaCellsConverted: 0,
        formulaCellsWithoutCachedValues: uncached.count,
        groups: resolved.groups.length,
        inputFiles: 1,
        inputRows: resolved.inputRows,
        outputFiles: outputNames.length,
        outputRows: resolved.groups.reduce(
          (total, group) => total + group.rows,
          0,
        ),
        pivotTablesRemoved,
        rowsDeleted: 0,
        sheetsCopiedUnchanged: 0,
        sheetsFiltered: 1,
        skippedRows: resolved.skippedRows,
        valuesOnly: options.values === true ? 1 : 0,
      },
    },
  };
}

async function splitAllWorksheetsSource(
  options: SplitWorkbookSourceOptions,
  open: (name: string, mediaType: string) => Promise<ByteSink>,
): Promise<SplitWorkbookSourceOutcome> {
  const { analysis, identity, outputNames, selection } =
    await resolveAllWorksheetSplitSource(options.input, options);

  const run = await runAllWorksheetSplit({
    analysis,
    identity,
    outputContext: "memory",
    outputNames,
    selection,
    signal: options.signal,
    open: async (index) => {
      const sink = await open(outputNames[index]!, analysis.mediaType);
      return {
        write: (chunk) => sink.write(chunk),
        close: async (detail) => {
          await sink.flush();
          options.onProgress?.({
            operation: SPLIT_OPERATION,
            stage: "building-workbooks",
            completed: index + 1,
            total: analysis.groups.length,
            detail: detail.output,
          });
        },
      };
    },
  });

  return {
    result: {
      operation: SPLIT_OPERATION,
      artifacts: outputNames.map((name) => ({
        kind: "file",
        mediaType: analysis.mediaType,
        path: name,
      })),
      warnings: run.warnings,
      metrics: run.metrics,
      outputs: run.outputs,
      summary: run.summary,
    },
  };
}

/**
 * Stack the rows of every worksheet that yields a table in every input workbook
 * into one table and write it as a single workbook. A worksheet yields a table
 * when it is visible (or `includeHiddenSheets` is set), passes the `sheets`
 * filter, has a resolvable header row, and holds at least one non-blank row
 * below it.
 *
 * This is the byte-level twin of `consolidateWorkbooks`: it reads the same
 * worksheet tables, hands them to the same consolidation core, and serializes
 * with the same deterministic writer, so a browser and the command line
 * produce byte-identical workbooks from the same inputs and options.
 */
export async function consolidateWorkbooksBytes(
  options: ConsolidateWorkbooksBytesOptions,
): Promise<ConsolidateWorkbooksBytesOutcome> {
  const { inputs, ...rest } = options;
  const chunks: Uint8Array[] = [];
  let size = 0;
  const { result, outputName, mappingDraft } = await consolidateWorkbookSources(
    {
      ...rest,
      inputs: inputs.map((input) => bytesSource(input.name, input.bytes)),
      output: {
        write: (chunk) => {
          chunks.push(chunk);
          size += chunk.length;
        },
        flush: () => Promise.resolve(),
        abort: () => {
          chunks.length = 0;
          return Promise.resolve();
        },
      },
    },
  );
  const bytes = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.length;
  }
  const outputs: ByteArtifact[] = [
    { name: outputName, bytes, mediaType: WORKBOOK_MEDIA_TYPE },
  ];
  if (mappingDraft) outputs.push(mappingDraft);
  return { result, outputs };
}

/**
 * `consolidateWorkbooksBytes` over workbooks read in pieces and an output
 * written as it is produced, so neither is held whole: a browser reads each
 * `File` through `Blob.slice` and writes to a file of its own. `output`
 * receives exactly the bytes `consolidateWorkbooksBytes` returns. On failure
 * or cancellation `output.abort()` runs before the operation rejects.
 */
export async function consolidateWorkbookSources(
  options: ConsolidateWorkbookSourcesOptions,
): Promise<ConsolidateWorkbookSourcesOutcome> {
  try {
    return await consolidateIntoSink(options);
  } catch (error) {
    // The original failure is what the caller needs; a sink that also fails
    // to discard cannot make that answer more useful.
    await options.output.abort().catch(() => undefined);
    throw error;
  }
}

async function consolidateIntoSink(
  options: ConsolidateWorkbookSourcesOptions,
): Promise<ConsolidateWorkbookSourcesOutcome> {
  throwIfAborted(options.signal, CONSOLIDATE_OPERATION, "memory");
  if (options.inputs.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_INPUTS,
      "At least one workbook is required.",
    );
  }
  if (options.mapping !== undefined && options.suggestMapping === true) {
    refuseMappingWithSuggestion();
  }
  // Validated before a single workbook is parsed, so an unusable mapping
  // never costs the caller a read.
  const mapping =
    options.mapping === undefined
      ? undefined
      : validateColumnMapping(options.mapping);

  const outputName = `${safeNameFragment(
    withoutWorkbookExtension(options.outputName ?? "consolidated"),
    "consolidated",
  )}${WORKBOOK_EXTENSION}`;

  const sources: ConsolidationSource[] = options.inputs.map((input) => ({
    file: input.name,
    source: input.name,
    details: { source: input.name },
    open: () =>
      Promise.resolve({
        name: input.name,
        size: input.size,
        readAt: (offset, length, signal) =>
          input.readAt(offset, length, signal),
        close: () => Promise.resolve(),
      }),
  }));
  const settings: ConsolidationSettings = {
    headerRow: options.headerRow,
    includeHiddenSheets: options.includeHiddenSheets,
    sheets: options.sheets,
    addSourceColumns: options.addSourceColumns,
    normalizeHeaders: options.normalizeHeaders,
    mapping,
    suggestMapping: options.suggestMapping === true,
    signal: options.signal,
    onProgress: options.onProgress,
    outputContext: "memory",
    yieldControl: true,
  };
  const plan = await planConsolidation(sources, settings);
  const { suggestion, unmappedColumns } = plan;
  await yieldToEventLoop();
  throwIfAborted(options.signal, CONSOLIDATE_OPERATION, "memory");
  const sheetName = options.outputSheetName ?? CONSOLIDATED_SHEET_NAME;
  assertSheetName(sheetName);
  await writeConsolidation(sources, plan, settings, sheetName, options.output);
  // A cancellation during the last write or flush still cancels: the caller
  // asked to stop before the output was complete.
  throwIfAborted(options.signal, CONSOLIDATE_OPERATION, "memory");
  options.onProgress?.({
    operation: CONSOLIDATE_OPERATION,
    stage: "writing-output",
    completed: 1,
    total: 1,
    detail: outputName,
  });

  const artifacts: Artifact[] = [
    {
      kind: "file",
      mediaType: WORKBOOK_MEDIA_TYPE,
      path: outputName,
    },
  ];
  let mappingDraft: ByteArtifact | undefined;
  if (suggestion) {
    mappingDraft = {
      name: SUGGESTED_MAPPING_FILE_NAME,
      bytes: new TextEncoder().encode(
        serializeColumnMapping(suggestion.mapping),
      ),
      mediaType: MAPPING_MEDIA_TYPE,
    };
    artifacts.push({
      kind: "file",
      mediaType: MAPPING_MEDIA_TYPE,
      path: SUGGESTED_MAPPING_FILE_NAME,
    });
    options.onProgress?.({
      operation: CONSOLIDATE_OPERATION,
      stage: "writing-mapping-draft",
      completed: 1,
      total: 1,
      detail: SUGGESTED_MAPPING_FILE_NAME,
    });
  }

  const result: ConsolidateWorkbooksBytesResult = {
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
      inputFiles: options.inputs.length,
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

  return { result, outputName, mappingDraft };
}

/**
 * Combine every worksheet of every input workbook into one workbook, keeping
 * each worksheet's cells and formatting and recording where it came from.
 */
export async function mergeWorkbooksBytes(
  options: MergeWorkbooksBytesOptions,
): Promise<ByteOperationOutcome<MergeWorkbooksMetric>> {
  const { inputs, ...rest } = options;
  const chunks: Uint8Array[] = [];
  const { result, outputName } = await mergeWorkbookSources({
    ...rest,
    inputs: inputs.map((input) => bytesSource(input.name, input.bytes)),
    output: {
      write: (chunk) => {
        chunks.push(chunk);
      },
      flush: () => Promise.resolve(),
      abort: () => {
        chunks.length = 0;
        return Promise.resolve();
      },
    },
  });
  return {
    result,
    outputs: [
      {
        name: outputName,
        bytes: concatenate(chunks),
        mediaType: result.artifacts[0]!.mediaType!,
      },
    ],
  };
}

export interface MergeWorkbookSourcesOptions extends Omit<
  MergeWorkbooksBytesOptions,
  "inputs"
> {
  /** Each workbook, read in pieces. Its `name` is what the index records. */
  inputs: RandomAccessSource[];
  /** Receives the merged workbook as it is written. */
  output: ByteSink;
}

export interface MergeWorkbookSourcesOutcome {
  result: OperationResult<MergeWorkbooksMetric>;
  /** The name of the workbook written to `output`. */
  outputName: string;
}

/**
 * `mergeWorkbooksBytes` over workbooks read in pieces and an output written
 * as it is produced, so neither is held whole: a browser reads each `File`
 * through `blobSource`. `output` receives exactly the bytes
 * `mergeWorkbooksBytes` returns; on failure or cancellation its `abort` runs
 * before the operation rejects.
 */
export async function mergeWorkbookSources(
  options: MergeWorkbookSourcesOptions,
): Promise<MergeWorkbookSourcesOutcome> {
  try {
    return await mergeIntoSink(options);
  } catch (error) {
    await options.output.abort().catch(() => undefined);
    throw error;
  }
}

async function mergeIntoSink(
  options: MergeWorkbookSourcesOptions,
): Promise<MergeWorkbookSourcesOutcome> {
  throwIfAborted(options.signal, MERGE_OPERATION, "memory");
  if (options.inputs.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_INPUTS,
      "At least one workbook is required.",
    );
  }

  const requested = options.outputName ?? `merged${WORKBOOK_EXTENSION}`;
  // A caller that asks for a macro-enabled name keeps it, because that is what
  // decides whether a single input's macro project may travel (see the merge
  // transplant). Every other name lands on .xlsx, as it always has.
  const macroOutput = isMacroWorkbookName(requested);
  const outputName = `${safeNameFragment(
    withoutWorkbookExtension(requested),
    "merged",
  )}${macroOutput ? MACRO_WORKBOOK_EXTENSION : WORKBOOK_EXTENSION}`;
  const buildOptions = { ...options, macroOutput };
  // Reads before the output starts yield and honour a cancellation too.
  const state = createMergeState(buildOptions, async () => {
    await yieldToEventLoop();
    throwIfAborted(options.signal, MERGE_OPERATION, "memory");
  });

  for (const [index, input] of options.inputs.entries()) {
    throwIfAborted(options.signal, MERGE_OPERATION, "memory");
    await appendWorkbookSheets(state, input.name, input);
    options.onProgress?.({
      operation: MERGE_OPERATION,
      stage: "merging-inputs",
      completed: index + 1,
      total: options.inputs.length,
      detail: input.name,
    });
  }

  const merged = await finishMergedWorkbook(state, buildOptions);
  const mediaType = merged.macroEnabled
    ? MACRO_WORKBOOK_MEDIA_TYPE
    : WORKBOOK_MEDIA_TYPE;
  await merged.write(
    (chunk) => {
      options.output.write(chunk);
    },
    async () => {
      await options.output.flush();
      await yieldToEventLoop();
      throwIfAborted(options.signal, MERGE_OPERATION, "memory");
    },
  );
  await options.output.flush();
  // The merged workbook was written asynchronously; honour a cancellation
  // that arrived while it was being written.
  throwIfAborted(options.signal, MERGE_OPERATION, "memory");

  return {
    outputName,
    result: {
      operation: MERGE_OPERATION,
      artifacts: [
        {
          kind: "file",
          mediaType,
          path: outputName,
        },
      ],
      warnings: merged.warnings,
      metrics: {
        inputFiles: options.inputs.length,
        outputSheets: merged.outputSheets,
        hiddenSheets: merged.hiddenSheets,
        formulaCellsWithoutCachedValues: merged.uncachedFormulas.length,
      },
    },
  };
}

/** Open an in-memory workbook for the streaming reader. */
function openInputBytes(input: WorkbookInputBytes): Promise<StreamedWorkbook> {
  return openWorkbookBytes(input.bytes, {
    file: input.name,
    source: input.name,
    details: { source: input.name },
  });
}

/**
 * Read every visible worksheet that holds data as a `Table`, from bytes.
 *
 * The byte twin of `readWorkbookTables`: same selection options, same header
 * resolution, same `Table` shape, with blank and repeated headers filled in and
 * numbered rather than refused. The byte surface had readers for Excel Tables
 * and named ranges but none for the worksheets themselves, so a browser caller
 * that wanted a worksheet's data had to take the record reader's display text
 * instead of the stored values.
 */
export async function readWorkbookTablesBytes(
  input: WorkbookInputBytes,
  options: ReadWorkbookOptions = {},
): Promise<Table[]> {
  return consolidationInputs(
    await workbookWorksheetReports(
      await openInputBytes(input),
      input.name,
      options,
    ),
  ).tables;
}

/**
 * Read every selected worksheet from bytes, reporting each one whether or not
 * it yielded a table, and how many cells of the read hold a formula the
 * workbook carries no calculated value for.
 *
 * `readWorkbookTablesBytes` is this list with the tables taken out of it, so a
 * caller that only wants the data can keep asking for the data. A caller that
 * has to tell a blank cell from a value the workbook never worked out, which no
 * `Table` can express, asks for this instead. The consolidation, merge, and
 * split operations are unchanged and go on treating an uncalculated formula as
 * empty.
 */
export async function readWorkbookWorksheetsBytes(
  input: WorkbookInputBytes,
  options: ReadWorkbookOptions = {},
): Promise<WorksheetImportReport[]> {
  return readWorksheetReports(input.bytes, input.name, options);
}

/**
 * Read one worksheet as text records, the shape template population and other
 * record-driven operations consume.
 */
export async function readWorksheetRecordsBytes(
  input: WorkbookInputBytes,
  options: ReadWorksheetRecordsBytesOptions = {},
): Promise<WorksheetRecords> {
  return workbookWorksheetRecords(await openInputBytes(input), options);
}

/**
 * Read the Excel Tables a workbook defines, with their data, from bytes.
 *
 * The byte twin of `readWorkbookExcelTables`: same definitions, same
 * selection options, same `WorkbookExcelTable` shape. This surface simply had
 * no way to reach them before, which is what left its `WorkbookExcelTable`
 * re-export pointing at a type nothing here produced.
 */
export async function readWorkbookExcelTablesBytes(
  input: WorkbookInputBytes,
  options: ReadWorkbookExcelTablesOptions = {},
): Promise<WorkbookExcelTable[]> {
  return workbookExcelTables(await openInputBytes(input), input.name, options);
}

/**
 * Read the named ranges a workbook defines, with their data, from bytes. The
 * byte twin of `readWorkbookNamedRanges`.
 */
export async function readWorkbookNamedRangesBytes(
  input: WorkbookInputBytes,
  options: ReadWorkbookNamedRangesOptions = {},
): Promise<WorkbookNamedRange[]> {
  return workbookNamedRanges(await openInputBytes(input), input.name, options);
}

/**
 * Describe a workbook's structure from bytes, without producing anything:
 * worksheets with their visibility, dimensions and header preview, Excel
 * Tables, named ranges, and a bounded sample of each column's values.
 *
 * The byte twin of `describeWorkbook`. Both hand the same parsed workbook to
 * the same engine, so one workbook yields a structurally identical description
 * whether the caller has a filesystem or only bytes.
 */
export async function describeWorkbookBytes(
  input: WorkbookInputBytes,
  options: DescribeWorkbookOptions = {},
): Promise<WorkbookDescriptionOutcome> {
  return describeWorkbookSource(bytesSource(input.name, input.bytes), options);
}

/**
 * `describeWorkbookBytes` over a workbook read in pieces, so it is never held
 * whole: a browser reads a `File` through `blobSource`. The description is the
 * one `describeWorkbookBytes` gives for the same bytes.
 */
export async function describeWorkbookSource(
  input: RandomAccessSource,
  options: DescribeWorkbookOptions = {},
): Promise<WorkbookDescriptionOutcome> {
  throwIfAborted(options.signal, INSPECT_OPERATION, "memory");
  const workbook = await StreamedWorkbook.open(input, {
    file: input.name,
    source: input.name,
    details: { source: input.name },
  });
  return describeStreamedWorkbook(workbook, input.name, options, "memory");
}

export { XLSX_ERRORS, type XlsxErrorCode } from "./errors.js";
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
export type {
  ConsolidateWorkbooksMetric,
  MergeWorkbooksMetric,
  ReadWorkbookExcelTablesOptions,
  ReadWorkbookNamedRangesOptions,
  ReadWorkbookOptions,
  SplitWorkbookByColumnMetric,
  SplitWorkbookByColumnPlanMetric,
  WorkbookExcelTable,
  WorkbookNamedRange,
  WorksheetRecords,
  WorksheetRegion,
  WorksheetTableReport,
} from "./shared.js";
export type { WorksheetImportReport } from "./operations/worksheets.js";
export type {
  AllWorksheetSplitSummary,
  SplitOutputDetail,
  SplitSheetDetail,
} from "./split/all-worksheet.js";
