import type * as DbPackage from "@consultchimps/db";
import type * as DbNodePackage from "@consultchimps/db/node";
import type * as FilesPackage from "@consultchimps/files";
import type * as PdfPackage from "@consultchimps/pdf";
import type * as PptxPackage from "@consultchimps/pptx";
import type * as XlsxPackage from "@consultchimps/xlsx";

/**
 * The packages a command needs, loaded when that command runs rather than when
 * the CLI starts. Each one takes one to several seconds to load, and without
 * this every command, --version and --help included, paid for all of them.
 * Each loader caches its promise, so a package loads at most once per run. The
 * imports above are types only.
 */
export type DbNodeModule = typeof DbNodePackage;

let files: Promise<typeof FilesPackage> | undefined;
let xlsx: Promise<typeof XlsxPackage> | undefined;
let pdf: Promise<typeof PdfPackage> | undefined;
let pptx: Promise<typeof PptxPackage> | undefined;
let db: Promise<typeof DbPackage> | undefined;
let dbNode: Promise<DbNodeModule> | undefined;

export function filesModule(): Promise<typeof FilesPackage> {
  return (files ??= import("@consultchimps/files"));
}

export function xlsxModule(): Promise<typeof XlsxPackage> {
  return (xlsx ??= import("@consultchimps/xlsx"));
}

export function pdfModule(): Promise<typeof PdfPackage> {
  return (pdf ??= import("@consultchimps/pdf"));
}

export function pptxModule(): Promise<typeof PptxPackage> {
  return (pptx ??= import("@consultchimps/pptx"));
}

export function dbModule(): Promise<typeof DbPackage> {
  return (db ??= import("@consultchimps/db"));
}

export function dbNodeModule(): Promise<DbNodeModule> {
  return (dbNode ??= import("@consultchimps/db/node"));
}
