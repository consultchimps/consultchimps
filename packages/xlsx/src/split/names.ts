/**
 * Output naming for the all-worksheet split, kept free of `node:path` so both
 * surfaces can reach it. A byte-level split returns portable filenames and a
 * file-level split joins those same names onto a directory, so the collision
 * suffixing that decides them has to live below the path handling rather than
 * inside it.
 */
import { safeNameFragment } from "@consultchimps/core";

export function safeFilenameSegment(value: string, fallback: string): string {
  return safeNameFragment(value, fallback);
}

/**
 * One portable filename per group value, in the order the values arrive.
 *
 * Two values that sanitize to the same name are told apart by a stable `-2`,
 * `-3` suffix, compared case-insensitively because Windows and macOS treat the
 * two names as one file.
 */
export function splitOutputFilenames(
  filenamePrefix: string | undefined,
  values: ReadonlyArray<boolean | null | number | string>,
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
    return filename;
  });
}
