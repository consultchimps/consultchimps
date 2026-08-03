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

const UNSAFE_NAME_CHARACTERS = /[<>:"/\\|?*]+/gu;
const WINDOWS_RESERVED_FILENAME =
  /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\..*)?$/iu;

// Linear-time replacement for a trailing [. ]+ regex, which CodeQL flags as
// polynomial on adversarial inputs with long runs of spaces.
function trimTrailingDotsAndSpaces(value: string): string {
  let end = value.length;
  while (end > 0) {
    const character = value[end - 1];
    if (character !== "." && character !== " ") {
      break;
    }
    end -= 1;
  }
  return value.slice(0, end);
}

/**
 * Reduce a caller-supplied name to a portable filename fragment. Byte
 * operations never touch a filesystem, but their output names become
 * download and archive entries that must stay valid everywhere.
 */
function truncateToUtf8Bytes(value: string, maxBytes: number): string {
  const encoder = new TextEncoder();
  let result = "";
  let total = 0;
  for (const character of value) {
    const size = encoder.encode(character).length;
    if (total + size > maxBytes) {
      break;
    }
    result += character;
    total += size;
  }
  return result;
}

function withoutControlCharacters(value: string): string {
  let result = "";
  for (const character of value) {
    const codePoint = character.codePointAt(0);
    if (codePoint !== undefined && codePoint >= 32 && codePoint !== 127) {
      result += character;
    }
  }
  return result;
}

export function safeNameFragment(value: string, fallback: string): string {
  const normalized = trimTrailingDotsAndSpaces(
    withoutControlCharacters(value.normalize("NFKC"))
      .replace(UNSAFE_NAME_CHARACTERS, "-")
      .replace(/\s+/gu, " ")
      .replace(/-+/gu, "-")
      .trim(),
  );
  // Cap the fragment by encoded size, truncating at code-point boundaries,
  // so generated names plus their page suffix stay well inside common
  // 255-byte filename limits even for multi-byte scripts and emoji.
  const limited = trimTrailingDotsAndSpaces(
    truncateToUtf8Bytes(normalized, 80),
  );
  const safe = limited || fallback;
  return WINDOWS_RESERVED_FILENAME.test(safe) ? `_${safe}` : safe;
}

export function withoutPdfExtension(name: string): string {
  return name.replace(/\.pdf$/iu, "");
}
