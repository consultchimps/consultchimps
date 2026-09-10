/**
 * How a referenced record reads in a foreign-key cell and in its picker.
 *
 * A foreign key stores a Record ID, which is stable and meaningless to read, so
 * a cell shows the referenced record's name with the Record ID beside it. Two
 * customers can share a name and the stored value is the id, so showing both is
 * what makes a choice unambiguous, and a record with no name to show is named
 * by its id alone.
 *
 * It is one function because two things need the same answer from two places:
 * the worker builds labels for a table the grid is not showing, and the grid
 * builds them from the rows it is showing when a table refers to itself. A
 * label that differed between the two would look like an edit that had not
 * taken.
 */
import type { CellValue } from "@consultchimps/tabular";

export function referenceLabel(
  recordId: string,
  name: CellValue | undefined,
): string {
  return typeof name === "string" && name.trim() !== ""
    ? `${name} (${recordId})`
    : recordId;
}
