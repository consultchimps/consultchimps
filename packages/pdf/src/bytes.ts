/**
 * Byte-level PDF operations for environments without a filesystem, such as
 * browsers. Inputs and outputs are in-memory bytes; artifact paths in the
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
import type { PDFDocument } from "pdf-lib";

import {
  createOutputDocument,
  MERGE_OPERATION,
  PDF_ERRORS,
  PDF_MEDIA_TYPE,
  parsePdf,
  safeNameFragment,
  SPLIT_OPERATION,
  splitOutputNames,
  withoutPdfExtension,
  type MergePdfsMetric,
  type SplitPdfMetric,
} from "./shared.js";

export interface PdfInputBytes {
  name: string;
  bytes: Uint8Array;
}

export interface SplitPdfBytesOptions extends OperationControlOptions {
  input: PdfInputBytes;
  filenamePrefix?: string | undefined;
}

export interface MergePdfsBytesOptions extends OperationControlOptions {
  inputs: PdfInputBytes[];
  outputName?: string | undefined;
}

interface ResolvedSplitBytes {
  outputNames: string[];
  pageCount: number;
  sourceDocument: PDFDocument;
}

async function resolveSplitPdfBytes(
  options: SplitPdfBytesOptions,
): Promise<ResolvedSplitBytes> {
  const sourceDocument = await parsePdf(
    options.input.bytes,
    options.input.name,
  );
  const pageCount = sourceDocument.getPageCount();
  if (pageCount === 0) {
    throw new ConsultChimpsError(
      PDF_ERRORS.PDF_NO_PAGES,
      "The input PDF contains no pages.",
      {
        details: { source: options.input.name },
      },
    );
  }

  const prefix = safeNameFragment(
    options.filenamePrefix ?? withoutPdfExtension(options.input.name),
    "document",
  );
  return {
    outputNames: splitOutputNames(prefix, pageCount),
    pageCount,
    sourceDocument,
  };
}

export async function planSplitPdfBytes(
  options: SplitPdfBytesOptions,
): Promise<OperationPlan<SplitPdfMetric>> {
  const resolved = await resolveSplitPdfBytes(options);
  return {
    operation: SPLIT_OPERATION,
    inputs: [options.input.name],
    outputs: resolved.outputNames.map((name) => ({
      kind: "file",
      mediaType: PDF_MEDIA_TYPE,
      path: name,
      exists: false,
    })),
    warnings: [],
    metrics: {
      inputFiles: 1,
      outputFiles: resolved.outputNames.length,
      pages: resolved.pageCount,
    },
  };
}

export async function splitPdfBytes(
  options: SplitPdfBytesOptions,
): Promise<ByteOperationOutcome<SplitPdfMetric>> {
  throwIfAborted(options.signal, SPLIT_OPERATION);
  const resolved = await resolveSplitPdfBytes(options);
  const outputs: ByteArtifact[] = [];

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
    const name = resolved.outputNames[index]!;
    outputs.push({
      name,
      bytes: await outputDocument.save(),
      mediaType: PDF_MEDIA_TYPE,
    });
    options.onProgress?.({
      operation: SPLIT_OPERATION,
      stage: "writing-pages",
      completed: index + 1,
      total: resolved.pageCount,
      detail: name,
    });
  }

  return {
    result: {
      operation: SPLIT_OPERATION,
      artifacts: outputs.map((output) => ({
        kind: "file",
        mediaType: PDF_MEDIA_TYPE,
        path: output.name,
      })),
      warnings: [],
      metrics: {
        inputFiles: 1,
        outputFiles: outputs.length,
        pages: resolved.pageCount,
      },
    },
    outputs,
  };
}

export async function mergePdfsBytes(
  options: MergePdfsBytesOptions,
): Promise<ByteOperationOutcome<MergePdfsMetric>> {
  throwIfAborted(options.signal, MERGE_OPERATION);
  if (options.inputs.length === 0) {
    throw new ConsultChimpsError(
      PDF_ERRORS.PDF_NO_INPUTS,
      "At least one input PDF is required.",
    );
  }

  const outputName = `${safeNameFragment(
    withoutPdfExtension(options.outputName ?? "combined.pdf"),
    "combined",
  )}.pdf`;
  const outputDocument = await createOutputDocument();
  let pageCount = 0;

  for (const [index, input] of options.inputs.entries()) {
    throwIfAborted(options.signal, MERGE_OPERATION);
    const inputDocument = await parsePdf(input.bytes, input.name);
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
      total: options.inputs.length,
      detail: input.name,
    });
  }

  throwIfAborted(options.signal, MERGE_OPERATION);
  const output: ByteArtifact = {
    name: outputName,
    bytes: await outputDocument.save(),
    mediaType: PDF_MEDIA_TYPE,
  };

  return {
    result: {
      operation: MERGE_OPERATION,
      artifacts: [
        {
          kind: "file",
          mediaType: PDF_MEDIA_TYPE,
          path: output.name,
        },
      ],
      warnings: [],
      metrics: {
        inputFiles: options.inputs.length,
        outputFiles: 1,
        pages: pageCount,
      },
    },
    outputs: [output],
  };
}

export {
  PDF_ERRORS,
  type MergePdfsMetric,
  type PdfErrorCode,
  type SplitPdfMetric,
} from "./shared.js";
