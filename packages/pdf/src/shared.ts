/**
 * Platform-neutral internals shared by the path-based and byte-based PDF
 * operations. This module must stay free of node:fs and node:path imports so
 * the byte entry point can run in browsers.
 */
import { ConsultChimpsError } from "@consultchimps/core";
import { PDFDocument } from "pdf-lib";

export const PDF_MEDIA_TYPE = "application/pdf";
export const SPLIT_OPERATION = "pdf.split";
export const MERGE_OPERATION = "pdf.merge";
// Identical inputs must produce byte-identical outputs, so generated
// documents carry a fixed timestamp instead of the current time.
const FIXED_PDF_DATE = new Date("1980-01-01T00:00:00.000Z");

/**
 * Stable, published error codes thrown by @consultchimps/pdf. Values are part
 * of the versioned public API; never change an existing value.
 */
export const PDF_ERRORS = {
  PDF_NO_INPUTS: "PDF_NO_INPUTS",
  PDF_NO_PAGES: "PDF_NO_PAGES",
  PDF_PAGE_COPY_FAILED: "PDF_PAGE_COPY_FAILED",
  PDF_READ_FAILED: "PDF_READ_FAILED",
} as const;

export type PdfErrorCode = (typeof PDF_ERRORS)[keyof typeof PDF_ERRORS];

export type SplitPdfMetric = "inputFiles" | "outputFiles" | "pages";
export type MergePdfsMetric = "inputFiles" | "outputFiles" | "pages";
export type MergePdfsPlanMetric = "inputFiles" | "outputFiles";

export async function createOutputDocument(): Promise<PDFDocument> {
  const document = await PDFDocument.create();
  document.setCreationDate(FIXED_PDF_DATE);
  document.setModificationDate(FIXED_PDF_DATE);
  return document;
}

export async function parsePdf(
  bytes: Uint8Array,
  sourceLabel: string,
): Promise<PDFDocument> {
  try {
    return await PDFDocument.load(bytes);
  } catch (error) {
    throw new ConsultChimpsError(
      PDF_ERRORS.PDF_READ_FAILED,
      `Could not read PDF: ${sourceLabel}`,
      {
        cause: error,
        details: { source: sourceLabel },
      },
    );
  }
}

export function splitOutputNames(prefix: string, pageCount: number): string[] {
  const width = Math.max(3, String(pageCount).length);
  return Array.from(
    { length: pageCount },
    (_, index) =>
      `${prefix}-page-${String(index + 1).padStart(width, "0")}.pdf`,
  );
}

// Filename sanitization lives in @consultchimps/core so the PDF, XLSX, and
// future operations derive output names identically.
export { safeNameFragment } from "@consultchimps/core";

export function withoutPdfExtension(name: string): string {
  return name.replace(/\.pdf$/iu, "");
}
