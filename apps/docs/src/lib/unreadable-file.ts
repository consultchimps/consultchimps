/**
 * The failure a chosen file gives once it can no longer be read: moved,
 * offline, removed, or changed since it was picked. Whether the worker reads
 * the file whole or in pieces, the visitor gets the same code and advice.
 */
import { ConsultChimpsError } from "@consultchimps/core";

export const FILE_UNREADABLE = "FILE_UNREADABLE";

export function unreadableFile(
  name: string,
  cause: unknown,
): ConsultChimpsError {
  return new ConsultChimpsError(
    FILE_UNREADABLE,
    `"${name}" could not be read. It may have moved, gone offline, or been removed since it was chosen. Choose it again, or pick another file`,
    { cause, details: { source: name } },
  );
}
