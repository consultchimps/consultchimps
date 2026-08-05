import path from "node:path";

import { safeNameFragment } from "@consultchimps/core";

export function safeFilenameSegment(value: string, fallback: string): string {
  return safeNameFragment(value, fallback);
}

export function splitOutputPaths(
  outputDirectory: string,
  filenamePrefix: string | undefined,
  values: Array<boolean | null | number | string>,
  extension = ".xlsx",
): string[] {
  const usedFilenames = new Set<string>();

  return values.map((value) => {
    const segment =
      value === null ? "blank" : safeFilenameSegment(String(value), "value");
    const base = filenamePrefix ? `${filenamePrefix}-${segment}` : segment;
    let filename = `${base}${extension}`;
    let suffix = 2;

    while (usedFilenames.has(filename.toLocaleLowerCase())) {
      filename = `${base}-${suffix}${extension}`;
      suffix += 1;
    }

    usedFilenames.add(filename.toLocaleLowerCase());
    return path.join(outputDirectory, filename);
  });
}
