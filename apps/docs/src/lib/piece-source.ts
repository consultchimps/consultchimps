/**
 * A chosen file read in pieces through `Blob.slice`, never whole, for the
 * operations that read workbooks as streams (ADR 0006).
 */
import type { RandomAccessSource } from "@consultchimps/core";
import { blobSource } from "@consultchimps/xlsx/bytes";

import type { NamedFile } from "./operation-tasks";
import { FILE_UNREADABLE, unreadableFile } from "./unreadable-file";

/**
 * A chosen file read in pieces. A read the browser refuses fails as the file
 * being unreadable, as a whole-file read does, not as a damaged workbook.
 */
export function pieceSource(input: NamedFile): RandomAccessSource {
  const source = blobSource(input.name, input.file);
  return {
    name: source.name,
    size: source.size,
    readAt: (offset, length, signal) =>
      source.readAt(offset, length, signal).catch((error: unknown) => {
        if (error instanceof RangeError) throw error;
        throw unreadableFile(input.name, error);
      }),
  };
}

/**
 * The workbook reader reports any failed read as a damaged workbook. When the
 * cause was the file itself going unreadable, that is the answer to give.
 */
export async function readingFiles<T>(work: () => Promise<T>): Promise<T> {
  try {
    return await work();
  } catch (error) {
    let cause: unknown = error;
    for (let depth = 0; cause instanceof Error && depth < 8; depth += 1) {
      if ((cause as { code?: unknown }).code === FILE_UNREADABLE) throw cause;
      cause = cause.cause;
    }
    throw error;
  }
}
