/**
 * Inspection as the operation worker runs it: the chosen file read through
 * `Blob.slice`, never whole. The description is the command line's.
 */
import {
  describeWorkbookSource,
  type DescribeWorkbookOptions,
  type WorkbookDescriptionOutcome,
} from "@consultchimps/xlsx/bytes";

import type { NamedFile } from "./operation-tasks";
import { pieceSource, readingFiles } from "./piece-source";

export function inspectFile(
  input: NamedFile,
  options: DescribeWorkbookOptions,
): Promise<WorkbookDescriptionOutcome> {
  return readingFiles(() =>
    describeWorkbookSource(pieceSource(input), options),
  );
}
