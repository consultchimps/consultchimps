/**
 * Byte-level workbook operations for environments without a filesystem, such
 * as browsers. Inputs and outputs are in-memory bytes; artifact paths in the
 * structured results carry portable output names. This module must stay free
 * of node:fs and node:path imports.
 */
import {
  ConsultChimpsError,
  throwIfAborted,
  type ByteArtifact,
  type ByteOperationOutcome,
  type OperationControlOptions,
  type OperationPlan,
  type OperationResult,
} from "@consultchimps/core";

import type { Table } from "@consultchimps/tabular";

import { XLSX_ERRORS } from "./errors.js";
import {
  analyzeAllWorksheetSplit,
  plannedAllWorksheetSplitMetrics,
  runAllWorksheetSplit,
  workbookExtensionOf,
  type AllWorksheetSplitAnalysis,
  type AllWorksheetSplitSelection,
  type AllWorksheetSplitSummary,
  type SplitOutputDetail,
  type SplitSourceIdentity,
} from "./split/all-worksheet.js";
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
  MACRO_WORKBOOK_EXTENSION,
  MACRO_WORKBOOK_MEDIA_TYPE,
  MERGE_OPERATION,
  parseWorkbookBytes,
  preservedSplitTemplateBytes,
  resolveSplitSource,
  safeNameFragment,
  skippedRowsWarning,
  splitOutputFileNames,
  SPLIT_OPERATION,
  WORKBOOK_EXTENSION,
  WORKBOOK_MEDIA_TYPE,
  withoutWorkbookExtension,
  workbookTables,
  workbookWorksheetRecords,
  type ConsolidateWorkbooksMetric,
  type MergeWorkbooksMetric,
  type ReadWorkbookOptions,
  type ResolvedSplitSource,
  type SplitWorkbookByColumnMetric,
  type SplitWorkbookByColumnPlanMetric,
  type WorksheetRecords,
} from "./shared.js";

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
   * Match columns whose headers differ only in case, spacing, or punctuation
   * (for example "Failed Checks" and "Failed_Checks") instead of requiring
   * the exact same header in every worksheet.
   */
  normalizeHeaders?: boolean | undefined;
  outputName?: string | undefined;
  outputSheetName?: string | undefined;
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

interface ResolvedSplitBytes extends ResolvedSplitSource {
  outputNames: string[];
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
function isAllWorksheetSplit(options: SplitWorkbookBytesOptions): boolean {
  return (
    !options.table &&
    !options.range &&
    !options.sheet &&
    options.preserveWorkbook !== false
  );
}

/**
 * The name every output of this split is built from.
 *
 * A byte split's outputs are downloads that land wherever the caller puts
 * them, with no chosen directory to tell one job's results from another's, so
 * they keep the source-derived prefix this surface has always used. The file
 * surface, which writes into a directory the caller named, uses the group
 * value alone.
 */
function splitFilenamePrefix(options: SplitWorkbookBytesOptions): string {
  return safeNameFragment(
    options.filenamePrefix ?? withoutWorkbookExtension(options.input.name),
    "split",
  );
}

interface ResolvedAllWorksheetSplitBytes {
  analysis: AllWorksheetSplitAnalysis;
  identity: SplitSourceIdentity;
  outputNames: string[];
  selection: AllWorksheetSplitSelection;
}

async function resolveAllWorksheetSplitBytes(
  options: SplitWorkbookBytesOptions,
): Promise<ResolvedAllWorksheetSplitBytes> {
  const identity: SplitSourceIdentity = {
    details: { source: options.input.name },
    label: options.input.name,
  };
  const extension = workbookExtensionOf(options.input.name, identity);
  const selection: AllWorksheetSplitSelection = {
    column: options.column,
    headerRow: options.headerRow,
    strict: options.strict,
    values: options.values,
  };
  const analysis = await analyzeAllWorksheetSplit(
    options.input.bytes,
    extension,
    selection,
    identity,
  );

  return {
    analysis,
    identity,
    outputNames: splitOutputFileNames(
      splitFilenamePrefix(options),
      analysis.groups.map((group) => group.display),
      extension,
    ),
    selection,
  };
}

async function resolveSplitWorkbookBytes(
  options: SplitWorkbookBytesOptions,
): Promise<ResolvedSplitBytes> {
  const resolved = await resolveSplitSource(
    options.input.bytes,
    {
      details: { source: options.input.name },
      file: options.input.name,
      label: options.input.name,
    },
    options,
  );

  return {
    ...resolved,
    outputNames: splitOutputFileNames(
      splitFilenamePrefix(options),
      resolved.grouped.groups.map((group) => group.value),
    ),
  };
}

/**
 * Report the workbooks a split would produce, and the rows it would skip,
 * without building any bytes.
 */
export async function planSplitWorkbookBytes(
  options: SplitWorkbookBytesOptions,
): Promise<OperationPlan<SplitWorkbookByColumnPlanMetric>> {
  if (isAllWorksheetSplit(options)) {
    const resolved = await resolveAllWorksheetSplitBytes(options);
    return {
      operation: SPLIT_OPERATION,
      inputs: [options.input.name],
      outputs: resolved.outputNames.map((name) => ({
        kind: "file",
        mediaType: resolved.analysis.mediaType,
        path: name,
        exists: false,
      })),
      warnings:
        resolved.analysis.skippedRows > 0
          ? [
              `Skipped ${resolved.analysis.skippedRows} row${resolved.analysis.skippedRows === 1 ? "" : "s"} with blank values in "${options.column}"; no blank-value workbook was created.`,
            ]
          : [],
      metrics: plannedAllWorksheetSplitMetrics(
        resolved.analysis,
        resolved.selection,
        resolved.outputNames.length,
      ),
    };
  }

  const resolved = await resolveSplitWorkbookBytes(options);
  const warnings =
    resolved.grouped.skippedRows > 0
      ? [skippedRowsWarning(resolved.grouped)]
      : [];

  return {
    operation: SPLIT_OPERATION,
    inputs: [options.input.name],
    outputs: resolved.outputNames.map((name) => ({
      kind: "file",
      mediaType: WORKBOOK_MEDIA_TYPE,
      path: name,
      exists: false,
    })),
    warnings,
    // A single-source split reports zero for the work only the all-worksheet
    // engine does, so both modes answer to the same metric names.
    metrics: {
      calcChainEntriesRemoved: 0,
      formulaCellsBlankedForRemovedRows: 0,
      formulaCellsConverted: 0,
      formulaCellsWithoutCachedValues: 0,
      groups: resolved.grouped.groups.length,
      inputFiles: 1,
      inputRows: resolved.table.rows.length,
      outputFiles: resolved.outputNames.length,
      pivotTablesRemoved: 0,
      rowsDeleted: 0,
      sheetsCopiedUnchanged: 0,
      sheetsFiltered: 1,
      skippedRows: resolved.grouped.skippedRows,
      valuesOnly: options.values === true ? 1 : 0,
    },
  };
}

/**
 * Split one workbook's rows into one workbook per distinct column value.
 *
 * Two engines sit behind one signature, chosen by exactly the rule the file
 * surface uses. With no table, range, or worksheet selection the split keeps
 * the whole workbook and filters every worksheet that carries the column, on
 * the layered engine, so every reference describing a moved row moves with it.
 * Selecting a region instead asks for one of the older, narrower modes - a
 * preserved Excel Table rewrite, or a compact single-worksheet rebuild - which
 * `shared.ts` still owns; `buildSplitGroupBytes` says why they stayed there.
 */
export async function splitWorkbookBytes(
  options: SplitWorkbookBytesOptions,
): Promise<SplitWorkbookBytesOutcome> {
  throwIfAborted(options.signal, SPLIT_OPERATION, "memory");
  if (isAllWorksheetSplit(options)) {
    return splitAllWorksheetsBytes(options);
  }
  const {
    grouped,
    outputNames,
    preservedTableDefinition,
    preserveWorkbook,
    table,
  } = await resolveSplitWorkbookBytes(options);
  const templateBytes = preserveWorkbook
    ? await preservedSplitTemplateBytes(options.input.bytes, options.values)
    : undefined;
  const outputs: ByteArtifact[] = [];
  let pivotTablesRemoved = 0;

  for (const [index, group] of grouped.groups.entries()) {
    throwIfAborted(options.signal, SPLIT_OPERATION, "memory");
    const name = outputNames[index]!;
    outputs.push({
      name,
      bytes: await buildSplitGroupBytes(group, {
        onPivotTablesRemoved: (removed) => {
          pivotTablesRemoved += removed;
        },
        preservedTableDefinition,
        sheetName: table.source?.sheet ?? "Split",
        templateBytes,
      }),
      mediaType: WORKBOOK_MEDIA_TYPE,
    });
    options.onProgress?.({
      operation: SPLIT_OPERATION,
      stage: "building-workbooks",
      completed: index + 1,
      total: grouped.groups.length,
      detail: name,
    });
  }

  // The last workbook was serialized asynchronously; honour a cancellation
  // that arrived while it was being built.
  throwIfAborted(options.signal, SPLIT_OPERATION, "memory");

  return {
    result: {
      operation: SPLIT_OPERATION,
      artifacts: outputs.map((output) => ({
        kind: "file",
        mediaType: WORKBOOK_MEDIA_TYPE,
        path: output.name,
      })),
      warnings: [
        ...(grouped.skippedRows > 0 ? [skippedRowsWarning(grouped)] : []),
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
        formulaCellsWithoutCachedValues: 0,
        groups: grouped.groups.length,
        inputFiles: 1,
        inputRows: table.rows.length,
        outputFiles: outputs.length,
        outputRows: grouped.groups.reduce(
          (total, group) => total + group.table.rows.length,
          0,
        ),
        pivotTablesRemoved,
        rowsDeleted: 0,
        sheetsCopiedUnchanged: 0,
        sheetsFiltered: 1,
        skippedRows: grouped.skippedRows,
        valuesOnly: options.values === true ? 1 : 0,
      },
    },
    outputs,
  };
}

/**
 * The byte surface's all-worksheet split: hold every finished workbook in
 * memory, because a caller with no filesystem has nowhere else to put one, and
 * report the same details and warnings the file surface reports.
 */
async function splitAllWorksheetsBytes(
  options: SplitWorkbookBytesOptions,
): Promise<SplitWorkbookBytesOutcome> {
  const { analysis, identity, outputNames, selection } =
    await resolveAllWorksheetSplitBytes(options);
  const outputs: ByteArtifact[] = [];

  const run = await runAllWorksheetSplit({
    analysis,
    identity,
    outputContext: "memory",
    outputNames,
    selection,
    signal: options.signal,
    write: (index, bytes, detail) => {
      outputs.push({
        name: detail.output,
        bytes,
        mediaType: analysis.mediaType,
      });
      options.onProgress?.({
        operation: SPLIT_OPERATION,
        stage: "building-workbooks",
        completed: index + 1,
        total: analysis.groups.length,
        detail: detail.output,
      });
    },
  });

  return {
    result: {
      operation: SPLIT_OPERATION,
      artifacts: outputs.map((output) => ({
        kind: "file",
        mediaType: analysis.mediaType,
        path: output.name,
      })),
      warnings: run.warnings,
      metrics: run.metrics,
      outputs: run.outputs,
      summary: run.summary,
    },
    outputs,
  };
}

/**
 * Stack the rows of every useful worksheet in every input workbook into one
 * table and write it as a single workbook.
 *
 * This is the byte-level twin of `consolidateWorkbooks`: it reads the same
 * worksheet tables, hands them to the same consolidation core, and serializes
 * with the same deterministic writer, so a browser and the command line
 * produce byte-identical workbooks from the same inputs and options.
 */
export async function consolidateWorkbooksBytes(
  options: ConsolidateWorkbooksBytesOptions,
): Promise<ByteOperationOutcome<ConsolidateWorkbooksMetric>> {
  throwIfAborted(options.signal, CONSOLIDATE_OPERATION, "memory");
  if (options.inputs.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_INPUTS,
      "At least one workbook is required.",
    );
  }

  const outputName = `${safeNameFragment(
    withoutWorkbookExtension(options.outputName ?? "consolidated"),
    "consolidated",
  )}${WORKBOOK_EXTENSION}`;

  const tables: Table[] = [];
  for (const [index, input] of options.inputs.entries()) {
    throwIfAborted(options.signal, CONSOLIDATE_OPERATION, "memory");
    tables.push(
      ...workbookTables(
        parseWorkbookBytes(input.bytes, input.name, {
          details: { source: input.name },
        }),
        input.name,
        options,
      ),
    );
    options.onProgress?.({
      operation: CONSOLIDATE_OPERATION,
      stage: "reading-workbooks",
      completed: index + 1,
      total: options.inputs.length,
      detail: input.name,
    });
  }

  const table = consolidateTables(tables, options);
  throwIfAborted(options.signal, CONSOLIDATE_OPERATION, "memory");
  const output: ByteArtifact = {
    name: outputName,
    bytes: buildTableWorkbookBytes(
      table,
      options.outputSheetName ?? CONSOLIDATED_SHEET_NAME,
    ),
    mediaType: WORKBOOK_MEDIA_TYPE,
  };
  options.onProgress?.({
    operation: CONSOLIDATE_OPERATION,
    stage: "writing-output",
    completed: 1,
    total: 1,
    detail: outputName,
  });

  return {
    result: {
      operation: CONSOLIDATE_OPERATION,
      artifacts: [
        {
          kind: "file",
          mediaType: WORKBOOK_MEDIA_TYPE,
          path: outputName,
        },
      ],
      warnings: [],
      metrics: {
        inputFiles: options.inputs.length,
        inputTables: tables.length,
        outputColumns: table.columns.length,
        outputRows: table.rows.length,
      },
    },
    outputs: [output],
  };
}

/**
 * Combine every worksheet of every input workbook into one workbook, keeping
 * each worksheet's cells and formatting and recording where it came from.
 */
export async function mergeWorkbooksBytes(
  options: MergeWorkbooksBytesOptions,
): Promise<ByteOperationOutcome<MergeWorkbooksMetric>> {
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
  const state = createMergeState(buildOptions);

  for (const [index, input] of options.inputs.entries()) {
    throwIfAborted(options.signal, MERGE_OPERATION, "memory");
    await appendWorkbookSheets(state, input.name, input.bytes);
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
  const output: ByteArtifact = {
    name: outputName,
    bytes: merged.bytes,
    mediaType,
  };
  // The merged workbook was serialized asynchronously; honour a cancellation
  // that arrived while it was being written.
  throwIfAborted(options.signal, MERGE_OPERATION, "memory");

  return {
    result: {
      operation: MERGE_OPERATION,
      artifacts: [
        {
          kind: "file",
          mediaType,
          path: output.name,
        },
      ],
      warnings: merged.warnings,
      metrics: {
        inputFiles: options.inputs.length,
        outputSheets: merged.outputSheets,
        hiddenSheets: merged.hiddenSheets,
      },
    },
    outputs: [output],
  };
}

/**
 * Read one worksheet as text records, the shape template population and other
 * record-driven operations consume.
 */
export async function readWorksheetRecordsBytes(
  input: WorkbookInputBytes,
  options: ReadWorksheetRecordsBytesOptions = {},
): Promise<WorksheetRecords> {
  return workbookWorksheetRecords(
    parseWorkbookBytes(input.bytes, input.name, {
      cellText: true,
      details: { source: input.name },
    }),
    options,
  );
}

export { XLSX_ERRORS, type XlsxErrorCode } from "./errors.js";
export type {
  ConsolidateWorkbooksMetric,
  MergeWorkbooksMetric,
  ReadWorkbookOptions,
  SplitWorkbookByColumnMetric,
  SplitWorkbookByColumnPlanMetric,
  WorkbookExcelTable,
  WorkbookNamedRange,
  WorksheetRecords,
} from "./shared.js";
export type {
  AllWorksheetSplitSummary,
  SplitOutputDetail,
  SplitSheetDetail,
} from "./split/all-worksheet.js";
