/**
 * The page-side half of the workspace worker: one stateful module Web Worker
 * that owns the database, driven through promises.
 *
 * A `WorkspaceClient` wraps one worker instance. The worker is created on the
 * first command rather than on construction, so merely mounting the page never
 * fetches it, and Next resolves the `new Worker(new URL(...), { type: "module" })`
 * form at build time into its own chunk, which keeps the pattern compatible with
 * the site's static export.
 *
 * Because the worker holds a single database, commands are run one at a time:
 * each call waits for the previous to settle before its command is posted, so a
 * create can never interleave with an open in the engine. That ordering lives
 * here rather than in the page, which only guards the buttons for the user's
 * sake.
 */
import { ConsultChimpsError } from "@consultchimps/core";
import type { CellValue } from "@consultchimps/tabular";

import type {
  UpdateWorkspaceCellCommand,
  WorkspaceCommand,
  WorkspaceEvent,
  WorkspaceSummary,
  WorkspaceTable,
} from "./workspace-protocol";

/**
 * One cell edit: which record, which column, and the new value. It is the
 * update command without its wire fields, so the two cannot drift apart.
 */
export type WorkspaceCellEdit = Omit<UpdateWorkspaceCellCommand, "type" | "id">;

/** Raised when the worker cannot start, so no command can be served. */
export const WORKSPACE_WORKER_UNAVAILABLE = "WORKSPACE_WORKER_UNAVAILABLE";

interface PendingCommand {
  readonly resolve: (event: WorkspaceEvent) => void;
  readonly reject: (error: unknown) => void;
}

export class WorkspaceClient {
  #worker: Worker | null = null;
  #pending = new Map<number, PendingCommand>();
  #lastId = 0;
  // A promise chain that serializes commands: each new command is appended and
  // runs only after the previous one settles.
  #queue: Promise<unknown> = Promise.resolve();

  /** Start a new, empty workspace. Resolves with its summary. */
  async create(): Promise<WorkspaceSummary> {
    const event = await this.#run((id) => ({ type: "create", id }));
    return this.#expectReady(event);
  }

  /**
   * Open a workspace from file bytes. A copy is taken and its buffer transferred
   * to the worker, so the caller's `bytes` stay valid.
   */
  async open(bytes: Uint8Array): Promise<WorkspaceSummary> {
    const buffer = bytes.slice().buffer;
    const event = await this.#run(
      (id) => ({ type: "open", id, buffer }),
      [buffer],
    );
    return this.#expectReady(event);
  }

  /** Serialize the held workspace to bytes for saving. */
  async serialize(): Promise<Uint8Array> {
    const event = await this.#run((id) => ({ type: "serialize", id }));
    if (event.type !== "serialized") {
      throw this.#unexpected(event);
    }
    return new Uint8Array(event.buffer);
  }

  /** Release the held workspace. */
  async close(): Promise<void> {
    const event = await this.#run((id) => ({ type: "close", id }));
    if (event.type !== "closed") {
      throw this.#unexpected(event);
    }
  }

  /* -----------------------------------------------------------------------
   * Record grid
   *
   * These share the queue above, so a table read cannot overtake a cell edit
   * that is still in the engine, and a workspace opened while an edit is in
   * flight cannot swap the database underneath it.
   * --------------------------------------------------------------------- */

  /** Name the tables the held workspace contains, ordered by name. */
  async listTables(): Promise<readonly string[]> {
    const event = await this.#run((id) => ({ type: "listTables", id }));
    if (event.type !== "tables") {
      throw this.#unexpected(event);
    }
    return event.tables;
  }

  /** Read one table's columns and every row it holds. */
  async readTable(name: string): Promise<WorkspaceTable> {
    const event = await this.#run((id) => ({ type: "readTable", id, name }));
    if (event.type !== "table") {
      throw this.#unexpected(event);
    }
    return event.table;
  }

  /**
   * Write one cell. Resolves with the value the database now holds, which is
   * not always the value that was sent, and rejects when the database refuses
   * it, so the caller can show the stored value or put the cell back.
   */
  async updateCell(edit: WorkspaceCellEdit): Promise<CellValue> {
    const event = await this.#run((id) => ({
      type: "updateCell",
      id,
      table: edit.table,
      recordId: edit.recordId,
      column: edit.column,
      value: edit.value,
    }));
    if (event.type !== "cellUpdated") {
      throw this.#unexpected(event);
    }
    return event.value;
  }

  /** Tear down the worker entirely, failing anything still pending. */
  terminate(): void {
    this.#worker?.terminate();
    this.#worker = null;
    this.#failEveryPending(
      "The workspace was closed before the task finished.",
    );
  }

  #expectReady(event: WorkspaceEvent): WorkspaceSummary {
    if (event.type !== "ready") {
      throw this.#unexpected(event);
    }
    return event.summary;
  }

  #unexpected(event: WorkspaceEvent): Error {
    return new Error(`Unexpected workspace worker reply: ${event.type}`);
  }

  #ensureWorker(): Worker {
    if (this.#worker) {
      return this.#worker;
    }
    const created = new Worker(
      new URL("../workers/workspace.worker.ts", import.meta.url),
      { type: "module" },
    );
    created.addEventListener("message", (event: MessageEvent<WorkspaceEvent>) =>
      this.#handleEvent(event.data),
    );
    created.addEventListener("error", () => {
      // A worker that cannot start would otherwise leave every caller waiting.
      // Drop the instance so the next command tries again from scratch.
      created.terminate();
      this.#worker = null;
      this.#failEveryPending(
        "The workspace engine could not start in this browser.",
      );
    });
    this.#worker = created;
    return created;
  }

  #handleEvent(event: WorkspaceEvent): void {
    const pending = this.#pending.get(event.id);
    if (!pending) {
      return;
    }
    this.#pending.delete(event.id);
    if (event.type === "error") {
      pending.reject(
        event.code === undefined
          ? new Error(event.message)
          : // Rebuilt with this bundle's class so the shared error formatting,
            // which tests for a ConsultChimpsError, still recognizes it.
            new ConsultChimpsError(event.code, event.message),
      );
      return;
    }
    pending.resolve(event);
  }

  #failEveryPending(reason: string): void {
    const failed = [...this.#pending.values()];
    this.#pending.clear();
    for (const pending of failed) {
      pending.reject(
        new ConsultChimpsError(WORKSPACE_WORKER_UNAVAILABLE, reason),
      );
    }
  }

  // Append a command to the serial queue and resolve with its reply. The chain
  // advances whether the command succeeds or fails, so one failure never stalls
  // the next command.
  #run(
    build: (id: number) => WorkspaceCommand,
    transfer: Transferable[] = [],
  ): Promise<WorkspaceEvent> {
    const result = this.#queue.then(() => this.#send(build, transfer));
    // Keep the chain alive even when this command rejects.
    this.#queue = result.then(
      () => undefined,
      () => undefined,
    );
    return result;
  }

  #send(
    build: (id: number) => WorkspaceCommand,
    transfer: Transferable[],
  ): Promise<WorkspaceEvent> {
    const worker = this.#ensureWorker();
    this.#lastId += 1;
    const id = this.#lastId;
    return new Promise<WorkspaceEvent>((resolve, reject) => {
      this.#pending.set(id, { resolve, reject });
      worker.postMessage(build(id), transfer);
    });
  }
}
