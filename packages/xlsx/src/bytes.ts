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
} from "@consultchimps/core";

import { XLSX_ERRORS } from "./errors.js";
import {
  appendWorkbookSheets,
  buildSplitGroupBytes,
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
  workbookWorksheetRecords,
  type MergeWorkbooksMetric,
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
  includeBlank?: boolean | undefined;
  includeHiddenSheets?: boolean | undefined;
  /**
   * Keep the complete source workbook and replace only the selected Excel
   * Table's rows. Defaults to true when a table is selected; not available
   * for named ranges or plain worksheet splits.
   */
  preserveWorkbook?: boolean | undefined;
  range?: string | undefined;
  sheet?: string | undefined;
  table?: string | undefined;
  values?: boolean | undefined;
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
  const filenamePrefix = safeNameFragment(
    options.filenamePrefix ?? withoutWorkbookExtension(options.input.name),
    "split",
  );

  return {
    ...resolved,
    outputNames: splitOutputFileNames(
      filenamePrefix,
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
    metrics: {
      groups: resolved.grouped.groups.length,
      inputFiles: 1,
      inputRows: resolved.table.rows.length,
      outputFiles: resolved.outputNames.length,
      skippedRows: resolved.grouped.skippedRows,
    },
  };
}

/** Split one workbook's rows into one workbook per distinct column value. */
export async function splitWorkbookBytes(
  options: SplitWorkbookBytesOptions,
): Promise<ByteOperationOutcome<SplitWorkbookByColumnMetric>> {
  throwIfAborted(options.signal, SPLIT_OPERATION, "memory");
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
        groups: grouped.groups.length,
        inputFiles: 1,
        inputRows: table.rows.length,
        outputFiles: outputs.length,
        outputRows: grouped.groups.reduce(
          (total, group) => total + group.table.rows.length,
          0,
        ),
        skippedRows: grouped.skippedRows,
      },
    },
    outputs,
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
  MergeWorkbooksMetric,
  SplitWorkbookByColumnMetric,
  SplitWorkbookByColumnPlanMetric,
  WorkbookExcelTable,
  WorkbookNamedRange,
  WorksheetRecords,
} from "./shared.js";
