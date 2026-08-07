/**
 * Interim I/O for the Tier-1 utilities: the deterministic-write helpers the
 * deleted `package-zip.ts` provided, kept verbatim so the Tier-1 modules stay
 * behavior-identical until they re-express onto L0's `WorkbookPackage` in
 * Phase 2 (at which point this file is deleted with them).
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

export async function generatePackageBytes(archive: JSZip): Promise<Buffer> {
  return archive.generateAsync({
    compression: "DEFLATE",
    type: "nodebuffer",
  });
}
