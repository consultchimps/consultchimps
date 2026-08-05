/**
 * Deterministic workbook-package writing. Every rewritten part carries a fixed
 * timestamp and never creates folder entries, so identical inputs and options
 * produce byte-identical workbooks and the output package keeps exactly the
 * entries the source package had.
 */
import type JSZip from "jszip";

const FIXED_PACKAGE_DATE = new Date("1980-01-01T00:00:00.000Z");

export function replacePackagePart(
  archive: JSZip,
  partPath: string,
  content: string,
): void {
  archive.file(partPath, content, {
    createFolders: false,
    date: FIXED_PACKAGE_DATE,
  });
}

export async function generatePackageBytes(
  archive: JSZip,
): Promise<Uint8Array> {
  return archive.generateAsync({
    compression: "DEFLATE",
    type: "uint8array",
  });
}
