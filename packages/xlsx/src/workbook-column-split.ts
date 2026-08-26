/**
 * L5 - the filesystem surface of the all-worksheet split.
 *
 * The operation itself lives in `src/split/all-worksheet.ts`, where the byte
 * surface can reach it too. What is left here is what a filesystem adds around
 * it: resolving and validating paths, refusing to overwrite the input, checking
 * every destination before a single output is built, and committing through a
 * staging directory so a failure halfway leaves the destination as it was.
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
import {
  analyzeAllWorksheetSplit,
  plannedAllWorksheetSplitMetrics,
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

  const analysis = await analyzeAllWorksheetSplit(
    workbookBytes,
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
      write: async (index, bytes, detail) => {
        const stagedOutput = path.join(
          transactionDirectory,
          `output-${String(index + 1).padStart(6, "0")}${resolved.extension}`,
        );
        await writeFile(stagedOutput, bytes);
        stagedOutputs.push(stagedOutput);
        options.onProgress?.({
          operation: SPLIT_OPERATION,
          stage: "staging-workbooks",
          completed: index + 1,
          total: resolved.analysis.groups.length,
          detail: `${path.basename(detail.output)} (${detail.sheets.map((sheet) => `${sheet.sheet}: kept ${sheet.retainedRows}, deleted ${sheet.deletedRows}`).join("; ")})`,
        });
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
