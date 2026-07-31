import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import {
  ConsultChimpsError,
  throwIfAborted,
  type OperationPlan,
  type OperationControlOptions,
  type OperationResult,
} from "@consultchimps/core";
import {
  ensureDirectory,
  ensureOutputAvailable,
  ensureParentDirectory,
  pathExists,
  refuseInputOverwrite,
} from "@consultchimps/files";
import { PDFDocument } from "pdf-lib";

import {
  createOutputDocument,
  MERGE_OPERATION,
  PDF_ERRORS,
  PDF_MEDIA_TYPE,
  SPLIT_OPERATION,
  splitOutputNames,
  type MergePdfsMetric,
  type MergePdfsPlanMetric,
  type SplitPdfMetric,
} from "./shared.js";

export {
  PDF_ERRORS,
  type MergePdfsMetric,
  type MergePdfsPlanMetric,
  type PdfErrorCode,
  type SplitPdfMetric,
} from "./shared.js";

export interface PdfWriteOptions {
  overwrite?: boolean | undefined;
}

export interface SplitPdfOptions
  extends PdfWriteOptions, OperationControlOptions {
  input: string;
  outputDirectory: string;
  filenamePrefix?: string | undefined;
}

export interface MergePdfsOptions
  extends PdfWriteOptions, OperationControlOptions {
  inputs: string[];
  output: string;
}

async function loadPdf(filePath: string): Promise<PDFDocument> {
  const absolutePath = path.resolve(filePath);

  try {
    const bytes = await readFile(absolutePath);
    return await PDFDocument.load(bytes);
  } catch (error) {
    throw new ConsultChimpsError(
      PDF_ERRORS.PDF_READ_FAILED,
      `Could not read PDF: ${absolutePath}`,
      {
        cause: error,
        details: { filePath: absolutePath },
      },
    );
  }
}

interface ResolvedSplit {
  absoluteInput: string;
  absoluteOutputDirectory: string;
  outputPaths: string[];
  pageCount: number;
  sourceDocument: PDFDocument;
}

async function resolveSplitPdf(
  options: SplitPdfOptions,
): Promise<ResolvedSplit> {
  const absoluteInput = path.resolve(options.input);
  const absoluteOutputDirectory = path.resolve(options.outputDirectory);
  const sourceDocument = await loadPdf(absoluteInput);
  const pageCount = sourceDocument.getPageCount();

  if (pageCount === 0) {
    throw new ConsultChimpsError(
      PDF_ERRORS.PDF_NO_PAGES,
      "The input PDF contains no pages.",
      {
        details: { inputPath: absoluteInput },
      },
    );
  }

  const prefix = options.filenamePrefix ?? path.parse(absoluteInput).name;
  const outputPaths = splitOutputNames(prefix, pageCount).map((name) =>
    path.join(absoluteOutputDirectory, name),
  );

  outputPaths.forEach((outputPath) =>
    refuseInputOverwrite(outputPath, [absoluteInput]),
  );

  return {
    absoluteInput,
    absoluteOutputDirectory,
    outputPaths,
    pageCount,
    sourceDocument,
  };
}

export async function planSplitPdf(
  options: SplitPdfOptions,
): Promise<OperationPlan<SplitPdfMetric>> {
  const resolved = await resolveSplitPdf(options);
  const outputs = await Promise.all(
    resolved.outputPaths.map(async (outputPath) => ({
      kind: "file" as const,
      mediaType: PDF_MEDIA_TYPE,
      path: outputPath,
      exists: await pathExists(outputPath),
    })),
  );

  const collisions = outputs.filter((output) => output.exists).length;
  const warnings =
    collisions > 0 && options.overwrite !== true
      ? [
          `${collisions} planned output file${collisions === 1 ? " already exists" : "s already exist"}; executing without overwrite will fail.`,
        ]
      : [];

  return {
    operation: SPLIT_OPERATION,
    inputs: [resolved.absoluteInput],
    outputs,
    warnings,
    metrics: {
      inputFiles: 1,
      outputFiles: outputs.length,
      pages: resolved.pageCount,
    },
  };
}

export async function splitPdf(
  options: SplitPdfOptions,
): Promise<OperationResult<SplitPdfMetric>> {
  throwIfAborted(options.signal, SPLIT_OPERATION);
  const resolved = await resolveSplitPdf(options);
  await ensureDirectory(resolved.absoluteOutputDirectory);
  await Promise.all(
    resolved.outputPaths.map((outputPath) =>
      ensureOutputAvailable(outputPath, { overwrite: options.overwrite }),
    ),
  );

  for (let index = 0; index < resolved.pageCount; index += 1) {
    throwIfAborted(options.signal, SPLIT_OPERATION);
    const outputDocument = await createOutputDocument();
    const [page] = await outputDocument.copyPages(resolved.sourceDocument, [
      index,
    ]);
    if (!page) {
      throw new ConsultChimpsError(
        PDF_ERRORS.PDF_PAGE_COPY_FAILED,
        `Could not copy page ${index + 1}.`,
      );
    }
    outputDocument.addPage(page);
    const outputPath = resolved.outputPaths[index]!;
    await writeFile(outputPath, await outputDocument.save());
    options.onProgress?.({
      operation: SPLIT_OPERATION,
      stage: "writing-pages",
      completed: index + 1,
      total: resolved.pageCount,
      detail: path.basename(outputPath),
    });
  }

  return {
    operation: SPLIT_OPERATION,
    artifacts: resolved.outputPaths.map((outputPath) => ({
      kind: "file",
      mediaType: PDF_MEDIA_TYPE,
      path: outputPath,
    })),
    warnings: [],
    metrics: {
      inputFiles: 1,
      outputFiles: resolved.outputPaths.length,
      pages: resolved.pageCount,
    },
  };
}

interface ResolvedMerge {
  absoluteInputs: string[];
  absoluteOutput: string;
}

function resolveMergePdfs(options: MergePdfsOptions): ResolvedMerge {
  if (options.inputs.length === 0) {
    throw new ConsultChimpsError(
      PDF_ERRORS.PDF_NO_INPUTS,
      "At least one input PDF is required.",
    );
  }

  const absoluteInputs = options.inputs.map((inputPath) =>
    path.resolve(inputPath),
  );
  const absoluteOutput = path.resolve(options.output);
  refuseInputOverwrite(absoluteOutput, absoluteInputs);

  return { absoluteInputs, absoluteOutput };
}

export async function planMergePdfs(
  options: MergePdfsOptions,
): Promise<OperationPlan<MergePdfsPlanMetric>> {
  const resolved = resolveMergePdfs(options);
  const exists = await pathExists(resolved.absoluteOutput);
  const warnings =
    exists && options.overwrite !== true
      ? [
          "The planned output file already exists; executing without overwrite will fail.",
        ]
      : [];

  return {
    operation: MERGE_OPERATION,
    inputs: resolved.absoluteInputs,
    outputs: [
      {
        kind: "file",
        mediaType: PDF_MEDIA_TYPE,
        path: resolved.absoluteOutput,
        exists,
      },
    ],
    warnings,
    metrics: {
      inputFiles: resolved.absoluteInputs.length,
      outputFiles: 1,
    },
  };
}

export async function mergePdfs(
  options: MergePdfsOptions,
): Promise<OperationResult<MergePdfsMetric>> {
  throwIfAborted(options.signal, MERGE_OPERATION);
  const resolved = resolveMergePdfs(options);
  await ensureOutputAvailable(resolved.absoluteOutput, {
    overwrite: options.overwrite,
  });
  await ensureParentDirectory(resolved.absoluteOutput);

  const outputDocument = await createOutputDocument();
  let pageCount = 0;

  for (const [index, inputPath] of resolved.absoluteInputs.entries()) {
    throwIfAborted(options.signal, MERGE_OPERATION);
    const inputDocument = await loadPdf(inputPath);
    const copiedPages = await outputDocument.copyPages(
      inputDocument,
      inputDocument.getPageIndices(),
    );
    copiedPages.forEach((page) => outputDocument.addPage(page));
    pageCount += copiedPages.length;
    options.onProgress?.({
      operation: MERGE_OPERATION,
      stage: "merging-inputs",
      completed: index + 1,
      total: resolved.absoluteInputs.length,
      detail: path.basename(inputPath),
    });
  }

  throwIfAborted(options.signal, MERGE_OPERATION);
  await writeFile(resolved.absoluteOutput, await outputDocument.save());

  return {
    operation: MERGE_OPERATION,
    artifacts: [
      {
        kind: "file",
        mediaType: PDF_MEDIA_TYPE,
        path: resolved.absoluteOutput,
      },
    ],
    warnings: [],
    metrics: {
      inputFiles: resolved.absoluteInputs.length,
      outputFiles: 1,
      pages: pageCount,
    },
  };
}
