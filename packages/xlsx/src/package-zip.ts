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
  return archive.generateAsync({ compression: "DEFLATE", type: "nodebuffer" });
}
