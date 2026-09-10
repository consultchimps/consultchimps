/**
 * The workspace the worker holds, and which one it is.
 *
 * The worker owns one `@consultchimps/db` database at a time. Creating or
 * opening a workspace replaces it, closing it releases it, and an import adds
 * tables to it. Everything the grid holds came from one of those moments: a
 * table name, a Record ID, a column, and a snapshot of the rows. None of those
 * mean the same thing after the database has moved on, and a Record ID such as
 * `CUST-0001` is very likely to exist in the next database too, so a command
 * that arrives late would not fail. It would land on the wrong record, or write
 * into a table it never saw.
 *
 * Commands therefore name the workspace they belong to, and this counter is
 * what they are checked against. It is a counter rather than a flag because a
 * workspace can change any number of times, and rather than a comparison of the
 * database object because the page never holds one.
 *
 * The database and the counter live behind one object for a reason: the rule is
 * that every change to the held database moves the generation, and the only way
 * to keep that rule is to leave no way to change one without the other. There is
 * no setter here that does not move the counter, so a handler cannot forget.
 *
 * The check lives here, beside the data, rather than in the page: the page
 * disables editing while it is busy, but that only stops a visitor from being
 * invited to make an edit that would be refused. It cannot stop one already in
 * flight, and a guard the worker does not enforce is a guard it cannot keep.
 */
import { ConsultChimpsError } from "@consultchimps/core";

import {
  WORKSPACE_STALE_EDIT,
  WORKSPACE_STALE_READ,
} from "./workspace-protocol";

/** What a refused command was trying to do, so the message can say so. */
export type WorkspaceAccess = "edit" | "read";

/**
 * All this needs of a workspace is that it can be closed, which is the only
 * thing it does to one. Typing it that narrowly is what lets the rules above be
 * tested without an engine.
 */
export interface ClosableWorkspace {
  close(): void;
}

export class HeldWorkspace<T extends ClosableWorkspace> {
  #held: T | null = null;
  #generation = 0;

  /**
   * The workspace held now. It is 0 before the first create or open, which no
   * command can name, so a command arriving before either is refused.
   */
  get generation(): number {
    return this.#generation;
  }

  /**
   * The held workspace. Every command that acts on one needs one to be open,
   * and says so the same way.
   */
  require(): T {
    if (this.#held === null) {
      throw new ConsultChimpsError(
        "WORKSPACE_NONE_OPEN",
        "There is no workspace open yet. Start a new workspace or open one first.",
      );
    }
    return this.#held;
  }

  /**
   * Hold a new workspace, closing the previous one first so its allocation is
   * never orphaned. Called only once the new workspace is built, so a failed
   * open leaves the previous one untouched.
   */
  replace(next: T): void {
    this.#held?.close();
    this.#held = next;
    this.#generation += 1;
  }

  /** Release the held workspace. What was read from it is stale from here on. */
  release(): void {
    this.#held?.close();
    this.#held = null;
    this.#generation += 1;
  }

  /**
   * Record that the held workspace itself changed: an import adds tables and
   * rows to the database already open rather than replacing it, and a snapshot
   * taken before that knows neither.
   */
  changed(): void {
    this.#generation += 1;
  }

  /**
   * Refuse a command that names a workspace other than the one held. The
   * numbers are deliberately kept out of the message: they are bookkeeping, and
   * what the reader needs is what happened to their edit.
   */
  assertCurrent(expected: number, access: WorkspaceAccess): void {
    if (expected === this.#generation) {
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
