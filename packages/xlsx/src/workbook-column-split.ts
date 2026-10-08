/**
 * L5 - the filesystem surface of the all-worksheet split.
 *
 * The operation itself lives in `src/split/all-worksheet.ts`, where the byte
 * surface can reach it too. What is left here is what a filesystem adds around
 * it: resolving and validating paths, refusing to overwrite the input, checking
 * every destination before a single output is built, and committing through a
 * staging directory so a failure halfway leaves the destination as it was.
 */
import { mkdtemp, open, rename, rm, stat } from "node:fs/promises";
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
  openRandomAccessSource,
  refuseInputOverwrite,
  type FileSource,
} from "@consultchimps/files";

import { XLSX_ERRORS } from "./errors.js";
import {
  analyzeAllWorksheetSplit,
  plannedAllWorksheetSplitMetrics,
  plannedAllWorksheetSplitWarnings,
  runAllWorksheetSplit,
  SPLIT_OPERATION,
  workbookExtensionOf,
  type AllWorksheetSplitAnalysis,
  type AllWorksheetSplitMetric,
  type AllWorksheetSplitSelection,
  type AllWorksheetSplitSummary,
  type SplitOutputDetail,
  type SplitSheetDetail,
  type SplitSourceIdentity,
  type WorkbookExtension,
} from "./split/all-worksheet.js";
import { safeFilenameSegment, splitOutputPaths } from "./split-filenames.js";

export type FullWorkbookSplitMetric = AllWorksheetSplitMetric;
export type { SplitOutputDetail, SplitSheetDetail };

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

/** The byte surface's summary, plus the directory this split wrote into. */
export interface FullWorkbookSplitSummary extends AllWorksheetSplitSummary {
  outputDirectory: string;
}

export interface FullWorkbookSplitResult extends OperationResult<FullWorkbookSplitMetric> {
  outputs: SplitOutputDetail[];
  summary: FullWorkbookSplitSummary;
}

interface ResolvedFullWorkbookSplit {
  absoluteInput: string;
  /** The input, open for reading in pieces until the split is done with it. */
  source: FileSource;
  absoluteOutputDirectory: string;
  analysis: AllWorksheetSplitAnalysis;
  existingOutputs: Set<string>;
  extension: WorkbookExtension;
  identity: SplitSourceIdentity;
  outputPaths: string[];
}

function isMissingPathError(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    error.code === "ENOENT"
  );
}

function splitSelection(
  options: FullWorkbookSplitOptions,
): AllWorksheetSplitSelection {
  return {
    column: options.column,
    headerRow: options.headerRow,
    strict: options.strict,
    values: options.values,
  };
}

async function resolveFullWorkbookSplit(
  options: FullWorkbookSplitOptions,
): Promise<ResolvedFullWorkbookSplit> {
  const absoluteInput = path.resolve(options.input);
  const identity: SplitSourceIdentity = {
    details: { inputPath: absoluteInput },
    label: absoluteInput,
  };
  const extension = workbookExtensionOf(absoluteInput, identity);

  let source: FileSource;
  try {
    const inputStat = await stat(absoluteInput);
    if (!inputStat.isFile()) {
      throw new Error("The input path is not a file.");
    }
    source = await openRandomAccessSource(absoluteInput);
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

  try {
    return await resolveOpenedSplit(
      options,
      absoluteInput,
      identity,
      extension,
      source,
    );
  } catch (error) {
    await source.close().catch(() => undefined);
    throw error;
  }
}

async function resolveOpenedSplit(
  options: FullWorkbookSplitOptions,
  absoluteInput: string,
  identity: SplitSourceIdentity,
  extension: WorkbookExtension,
  source: FileSource,
): Promise<ResolvedFullWorkbookSplit> {
  const analysis = await analyzeAllWorksheetSplit(
    source,
    extension,
    splitSelection(options),
    identity,
  );

  const absoluteOutputDirectory = path.resolve(options.outputDirectory);
  const prefix = options.filenamePrefix
    ? safeFilenameSegment(options.filenamePrefix, "split")
    : undefined;
  const outputPaths = splitOutputPaths(
    absoluteOutputDirectory,
    prefix,
    analysis.groups.map((group) => group.display),
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
    source,
    absoluteOutputDirectory,
    analysis,
    existingOutputs,
    extension,
    identity,
    outputPaths,
  };
}

export async function planFullWorkbookSplit(
  options: FullWorkbookSplitOptions,
): Promise<OperationPlan<Exclude<FullWorkbookSplitMetric, "outputRows">>> {
  const resolved = await resolveFullWorkbookSplit(options);
  await resolved.source.close();
  const warnings = plannedAllWorksheetSplitWarnings(resolved.analysis);
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
      mediaType: resolved.analysis.mediaType,
      path: outputPath,
    })),
    warnings,
    metrics: plannedAllWorksheetSplitMetrics(
      resolved.analysis,
      splitSelection(options),
      resolved.outputPaths.length,
    ),
  };
}

export async function splitFullWorkbookByColumn(
  options: FullWorkbookSplitOptions,
): Promise<FullWorkbookSplitResult> {
  throwIfAborted(options.signal, SPLIT_OPERATION);
  const resolved = await resolveFullWorkbookSplit(options);
  try {
    return await splitResolved(options, resolved);
  } finally {
    await resolved.source.close().catch(() => undefined);
  }
}

/** A staged output written as it is produced, in pieces of up to 1 MiB. */
async function stagedFile(filePath: string): Promise<{
  write(chunk: Uint8Array): Promise<void>;
  close(): Promise<void>;
  abort(): Promise<void>;
}> {
  const handle = await open(filePath, "wx");
  let pending: Uint8Array[] = [];
  let pendingSize = 0;
  const flush = async (): Promise<void> => {
    if (pendingSize === 0) return;
    const joined = new Uint8Array(pendingSize);
    let offset = 0;
    for (const chunk of pending) {
      joined.set(chunk, offset);
      offset += chunk.length;
    }
    pending = [];
    pendingSize = 0;
    await handle.write(joined);
  };
  return {
    async write(chunk) {
      pending.push(chunk);
      pendingSize += chunk.length;
      if (pendingSize >= 1024 * 1024) await flush();
    },
    async close() {
      try {
        await flush();
      } finally {
        await handle.close();
      }
    },
    async abort() {
      await handle.close();
    },
  };
}

async function splitResolved(
  options: FullWorkbookSplitOptions,
  resolved: ResolvedFullWorkbookSplit,
): Promise<FullWorkbookSplitResult> {
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

  try {
    const run = await runAllWorksheetSplit({
      analysis: resolved.analysis,
      identity: resolved.identity,
      outputContext: "files",
      outputNames: resolved.outputPaths,
      selection: splitSelection(options),
      signal: options.signal,
      open: async (index) => {
        const stagedOutput = path.join(
          transactionDirectory,
          `output-${String(index + 1).padStart(6, "0")}${resolved.extension}`,
        );
        const file = await stagedFile(stagedOutput);
        return {
          write: (chunk) => file.write(chunk),
          abort: () => file.abort(),
          close: async (detail) => {
            await file.close();
            stagedOutputs.push(stagedOutput);
            options.onProgress?.({
              operation: SPLIT_OPERATION,
              stage: "staging-workbooks",
              completed: index + 1,
              total: resolved.analysis.groups.length,
              detail: `${path.basename(detail.output)} (${detail.sheets.map((sheet) => `${sheet.sheet}: kept ${sheet.retainedRows}, deleted ${sheet.deletedRows}`).join("; ")})`,
            });
          },
        };
      },
    });

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

    return {
      operation: SPLIT_OPERATION,
      artifacts: resolved.outputPaths.map((outputPath) => ({
        kind: "file",
        mediaType: resolved.analysis.mediaType,
        path: outputPath,
      })),
      warnings: run.warnings,
      metrics: run.metrics,
      outputs: run.outputs,
      summary: {
        ...run.summary,
        outputDirectory: resolved.absoluteOutputDirectory,
      },
    };
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
}
