/**
 * The explanations the grid is showing, and whose they are.
 *
 * An explanation belongs to the thing that failed, not to the page. A refusal
 * names one cell: the value it was given, why the database would not hold it,
 * and what the cell holds instead. That sentence stops being useful only when
 * the cell it is about is dealt with.
 *
 * So the rule, once:
 *
 * - A failure is recorded against its cell. A cell here is one column of one
 *   record of one table, which is the same thing an edit is addressed to.
 * - Only the same cell replaces it: a later attempt on that cell clears it if
 *   the database accepts, and replaces it if the database refuses again.
 * - A success anywhere else leaves it standing. Several cells can be refused
 *   and each keeps its own explanation until it is dealt with.
 * - The visitor can dismiss the lot.
 *
 * Before this there was one string, cleared by any success. Commit an invalid
 * value in one cell and a valid one in another, and the second reply erased the
 * first's explanation: a cell had snapped back to its old value with nothing on
 * screen saying why. The keying is what makes that impossible rather than
 * unlikely, because there is no longer a slot for an unrelated success to
 * clear.
 *
 * A read that fails is recorded the same way, against the table it was a read
 * of, so a refused edit and an unreadable table cannot erase each other either.
 */

/** What is standing, oldest first, keyed by what it is about. */
export type CellFailures = ReadonlyMap<string, string>;

/** The empty set, which is what the grid starts and dismisses to. */
export const NO_FAILURES: CellFailures = new Map<string, string>();

// Keys encode their parts rather than joining them with a separator, because a
// table or column name is allowed to contain anything a separator could use.
// Two different cells therefore cannot collide.

/** Which cell an edit was addressed to. */
export function cellKey(
  table: string,
  recordId: string,
  column: string,
): string {
  return JSON.stringify(["cell", table, recordId, column]);
}

/** Which table a read was a read of. */
export function tableKey(table: string): string {
  return JSON.stringify(["table", table]);
}

/**
 * Record a failure. Deleted before it is set so that re-refusing a cell moves
 * its explanation to the end, which is where the newest one is read from.
 */
export function withFailure(
  failures: CellFailures,
  key: string,
  message: string,
): CellFailures {
  const next = new Map(failures);
  next.delete(key);
  next.set(key, message);
  return next;
}

/** Drop what was standing against one cell or table, because it succeeded. */
export function withoutFailure(
  failures: CellFailures,
  key: string,
): CellFailures {
  if (!failures.has(key)) {
    return failures;
  }
  const next = new Map(failures);
  next.delete(key);
  return next;
}

/**
 * What to show, or null when nothing is standing.
 *
 * The newest explanation in full, and a count of the rest. In full because it
 * is the one the visitor just earned and it carries the reason and the way out;
 * a count for the rest because these sentences are long, and stacking them
 * would bury the newest under the ones already read. Each of the others is
 * still there, and shows again as the cells ahead of it are dealt with.
 */
export function failureReport(failures: CellFailures): string | null {
  const messages = [...failures.values()];
  const newest = messages[messages.length - 1];
  if (newest === undefined) {
    return null;
  }
  const others = messages.length - 1;
  if (others === 0) {
    return newest;
  }
  return `${newest}\n\n${
    others === 1
      ? "1 other edit was refused as well, and that cell still holds the value the workspace has"
      : `${others} other edits were refused as well, and those cells still hold the values the workspace has`
  }`;
}
