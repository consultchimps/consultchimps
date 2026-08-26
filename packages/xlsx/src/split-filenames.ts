import path from "node:path";

import { safeFilenameSegment, splitOutputFilenames } from "./split/names.js";

export { safeFilenameSegment };

export function splitOutputPaths(
  outputDirectory: string,
  filenamePrefix: string | undefined,
  values: Array<boolean | null | number | string>,
  extension = ".xlsx",
): string[] {
  return splitOutputFilenames(filenamePrefix, values, extension).map(
    (filename) => path.join(outputDirectory, filename),
  );
}
