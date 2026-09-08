/**
 * Which database the workspace worker currently holds.
 *
 * The worker owns one `@consultchimps/db` database at a time, and creating,
 * opening, or closing a workspace replaces it. Everything the grid holds came
 * from one of those databases: a table name, a Record ID, a column. None of
 * those identifiers mean the same thing in the next database, and a Record ID
 * such as `CUST-0001` is very likely to exist in both, so an edit that arrives
 * after a replacement would not fail, it would land on the wrong record.
 *
 * Commands therefore name the workspace they belong to, and this counter is
 * what they are checked against. It is a counter rather than a flag because a
 * workspace can be replaced any number of times, and rather than a comparison
 * of the database object because the page never holds one.
 *
 * The check lives here, beside the data, rather than in the page: the page
 * disables editing while it is busy, but that only stops a visitor from being
 * invited to make an edit that would be refused. It cannot stop one already in
 * flight, and a guard the worker does not enforce is a guard the worker cannot
 * keep.
 */
import { ConsultChimpsError } from "@consultchimps/core";

import {
  WORKSPACE_STALE_EDIT,
  WORKSPACE_STALE_READ,
} from "./workspace-protocol";

/** What a refused command was trying to do, so the message can say so. */
export type WorkspaceAccess = "edit" | "read";

export class WorkspaceGenerations {
  #current = 0;

  /**
   * The workspace held now. It is 0 before the first create or open, which no
   * command can name, so a command that arrives before either is refused.
   */
  get current(): number {
    return this.#current;
  }

  /** Record that the held database has been replaced or released. */
  replaced(): number {
    this.#current += 1;
    return this.#current;
  }

  /**
   * Refuse a command that names a workspace other than the one held. The
   * numbers are deliberately kept out of the message: they are bookkeeping, and
   * what the reader needs is what happened to their edit.
   */
  assertCurrent(expected: number, access: WorkspaceAccess): void {
    if (expected === this.#current) {
      return;
    }
    throw access === "edit"
      ? new ConsultChimpsError(
          WORKSPACE_STALE_EDIT,
          "That edit belongs to a workspace that is no longer open, so it was not applied. Make it again in the workspace open now.",
        )
      : new ConsultChimpsError(
          WORKSPACE_STALE_READ,
          "Those records belong to a workspace that is no longer open.",
        );
  }
}
