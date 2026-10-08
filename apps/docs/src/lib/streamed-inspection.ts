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
import { PieceReads } from "./piece-source";

/** The operation's name, as the library reports a cancellation. */
const INSPECT_OPERATION = "sheets.inspect";

export function inspectFile(
  input: NamedFile,
  options: DescribeWorkbookOptions,
): Promise<WorkbookDescriptionOutcome> {
  const reads = new PieceReads(INSPECT_OPERATION, options.signal);
  return reads.run(() => describeWorkbookSource(reads.source(input), options));
}
