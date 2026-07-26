import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { ConsultChimpsError, type OperationResult } from "@consultchimps/core";
import {
  ensureDirectory,
  ensureOutputAvailable,
  ensureParentDirectory,
  refuseInputOverwrite,
} from "@consultchimps/files";
import { PDFDocument } from "pdf-lib";

export interface PdfWriteOptions {
  overwrite?: boolean | undefined;
}

export interface SplitPdfOptions extends PdfWriteOptions {
  filenamePrefix?: string | undefined;
}

async function loadPdf(filePath: string): Promise<PDFDocument> {
  const absolutePath = path.resolve(filePath);

  try {
    const bytes = await readFile(absolutePath);
    return await PDFDocument.load(bytes);
  } catch (error) {
    throw new ConsultChimpsError(
      "PDF_READ_FAILED",
      `Could not read PDF: ${absolutePath}`,
      {
        cause: error,
        details: { filePath: absolutePath },
      },
    );
  }
}

export async function splitPdf(
  inputPath: string,
  outputDirectory: string,
  options: SplitPdfOptions = {},
): Promise<OperationResult> {
  const absoluteInput = path.resolve(inputPath);
  const absoluteOutputDirectory = await ensureDirectory(outputDirectory);
  const sourceDocument = await loadPdf(absoluteInput);
  const pageCount = sourceDocument.getPageCount();

  if (pageCount === 0) {
    throw new ConsultChimpsError(
      "PDF_NO_PAGES",
      "The input PDF contains no pages.",
      {
        details: { inputPath: absoluteInput },
      },
    );
  }

  const width = Math.max(3, String(pageCount).length);
  const prefix = options.filenamePrefix ?? path.parse(absoluteInput).name;
  const outputPaths = Array.from({ length: pageCount }, (_, index) =>
    path.join(
      absoluteOutputDirectory,
      `${prefix}-page-${String(index + 1).padStart(width, "0")}.pdf`,
    ),
  );

  await Promise.all(
    outputPaths.map((outputPath) =>
      ensureOutputAvailable(outputPath, { overwrite: options.overwrite }),
    ),
  );

  for (let index = 0; index < pageCount; index += 1) {
    const outputDocument = await PDFDocument.create();
    const [page] = await outputDocument.copyPages(sourceDocument, [index]);
    if (!page) {
      throw new ConsultChimpsError(
        "PDF_PAGE_COPY_FAILED",
        `Could not copy page ${index + 1}.`,
      );
    }
    outputDocument.addPage(page);
    await writeFile(outputPaths[index]!, await outputDocument.save());
  }

  return {
    operation: "pdf.split",
    artifacts: outputPaths.map((outputPath) => ({
      kind: "file",
      mediaType: "application/pdf",
      path: outputPath,
    })),
    warnings: [],
    metrics: {
      inputFiles: 1,
      outputFiles: outputPaths.length,
      pages: pageCount,
    },
  };
}

export async function mergePdfs(
  inputPaths: string[],
  outputPath: string,
  options: PdfWriteOptions = {},
): Promise<OperationResult> {
  if (inputPaths.length === 0) {
    throw new ConsultChimpsError(
      "PDF_NO_INPUTS",
      "At least one input PDF is required.",
    );
  }

  const absoluteInputs = inputPaths.map((inputPath) => path.resolve(inputPath));
  const absoluteOutput = path.resolve(outputPath);
  refuseInputOverwrite(absoluteOutput, absoluteInputs);
  await ensureOutputAvailable(absoluteOutput, { overwrite: options.overwrite });
  await ensureParentDirectory(absoluteOutput);

  const outputDocument = await PDFDocument.create();
  let pageCount = 0;

  for (const inputPath of absoluteInputs) {
    const inputDocument = await loadPdf(inputPath);
    const copiedPages = await outputDocument.copyPages(
      inputDocument,
      inputDocument.getPageIndices(),
    );
    copiedPages.forEach((page) => outputDocument.addPage(page));
    pageCount += copiedPages.length;
  }

  await writeFile(absoluteOutput, await outputDocument.save());

  return {
    operation: "pdf.merge",
    artifacts: [
      {
        kind: "file",
        mediaType: "application/pdf",
        path: absoluteOutput,
      },
    ],
    warnings: [],
    metrics: {
      inputFiles: absoluteInputs.length,
      outputFiles: 1,
      pages: pageCount,
    },
  };
}
