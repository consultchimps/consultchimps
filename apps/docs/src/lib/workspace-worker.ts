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

import type { WorkspaceImportKind } from "./accepted-files";

import type {
  ImportedTableSummary,
  ImportSourceDescription,
  ImportTableChoice,
  UpdateWorkspaceCellCommand,
  UpdateWorkspaceCellsCommand,
  WorkspaceCommand,
  WorkspaceEvent,
  WorkspaceSummary,
  WorkspaceTable,
} from "./workspace-protocol";

/** What an import created, as the page reports it. */
export interface WorkspaceImportResult {
  /** The workspace as it now stands, table listing included. */
  readonly summary: WorkspaceSummary;
  /** The tables this import created. */
  readonly tables: readonly ImportedTableSummary[];
}

/**
 * One cell edit: which workspace, which record, which column, and the new
 * value. It is the update command without its wire fields, so the two cannot
 * drift apart.
 */
export type WorkspaceCellEdit = Omit<UpdateWorkspaceCellCommand, "type" | "id">;

/**
 * One gesture: the cells a paste or a fill writes, in one command. It is the
 * batch command without its wire fields, for the same reason as above.
 */
export type WorkspaceCellsEdit = Omit<
  UpdateWorkspaceCellsCommand,
  "type" | "id"
>;

/**
 * What became of one cell of a gesture, as the page reads it.
 *
 * A refusal arrives as data on the wire and is rebuilt into an error here, the
 * same way a whole command's failure is, so the caller has one kind of thing to
 * hand to the shared error formatting.
 */
export type WorkspaceCellOutcome =
  | {
      readonly accepted: true;
      readonly recordId: string;
      readonly column: string;
      readonly value: CellValue;
    }
  | {
      readonly accepted: false;
      readonly recordId: string;
      readonly column: string;
      readonly error: unknown;
    };

/** Raised when the worker cannot start, so no command can be served. */
export const WORKSPACE_WORKER_UNAVAILABLE = "WORKSPACE_WORKER_UNAVAILABLE";

const CLOSED_REASON = "The workspace was closed before the task finished.";

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
  // Set by terminate() and never cleared: a torn-down client must not start a
  // worker again, however many commands were still waiting in the queue.
  #terminated = false;

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

  /* -------------------------------------------------------------------------
   * Import. Both commands carry the file, so the worker holds no state between
   * choosing a file and importing from it. A copy is taken and its buffer
   * transferred, so the caller's `bytes` stay valid and can be sent again.
   * ---------------------------------------------------------------------- */

  /** List the tables a `.xlsx`, `.xlsm`, or `.csv` file could contribute. */
  async describeImport(
    fileName: string,
    kind: WorkspaceImportKind,
    bytes: Uint8Array,
  ): Promise<readonly ImportSourceDescription[]> {
    const buffer = bytes.slice().buffer;
    const event = await this.#run(
      (id) => ({ type: "describeImport", id, fileName, kind, buffer }),
      [buffer],
    );
    if (event.type !== "importSources") {
      throw this.#unexpected(event);
    }
    return event.sources;
  }

  /** Create the chosen tables in the held workspace and fill them. */
  async importFile(
    fileName: string,
    kind: WorkspaceImportKind,
    bytes: Uint8Array,
    tables: readonly ImportTableChoice[],
  ): Promise<WorkspaceImportResult> {
    const buffer = bytes.slice().buffer;
    const event = await this.#run(
      (id) => ({ type: "import", id, fileName, kind, buffer, tables }),
      [buffer],
    );
    if (event.type !== "imported") {
      throw this.#unexpected(event);
    }
    return { summary: event.summary, tables: event.tables };
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
   * flight cannot swap the database underneath it. Which tables exist is not
   * asked here: that is the workspace summary every create, open, and import
   * already reports.
   * --------------------------------------------------------------------- */

  /**
   * Read one table's columns and every row it holds, from the workspace
   * `generation` names. Rejects with `WORKSPACE_STALE_READ` once that workspace
   * has been replaced or closed.
   */
  async readTable(name: string, generation: number): Promise<WorkspaceTable> {
    const event = await this.#run((id) => ({
      type: "readTable",
      id,
      name,
      generation,
    }));
    if (event.type !== "table") {
      throw this.#unexpected(event);
    }
    return event.table;
  }

  /**
   * Write one cell. Resolves with the value the database now holds, which is
   * not always the value that was sent, and rejects when the database refuses
   * it, so the caller can show the stored value or put the cell back. An edit
   * naming a workspace that is no longer held is refused with
   * `WORKSPACE_STALE_EDIT` rather than applied to the one that replaced it.
   */
  async updateCell(edit: WorkspaceCellEdit): Promise<CellValue> {
    const event = await this.#run((id) => ({
      type: "updateCell",
      id,
      generation: edit.generation,
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

  /**
   * Write many cells as one step, which is what one paste or one fill is.
   *
   * Resolves with one outcome per write, in the order they were sent: the value
   * the database now holds, or the refusal for that cell. The accepted writes
   * commit together, so a refusal in the middle of a gesture does not discard
   * the cells beside it. The whole gesture is rejected, and nothing is written,
   * when it names a workspace that is no longer held or asks for more cells than
   * one step applies.
   */
  async updateCells(
    gesture: WorkspaceCellsEdit,
  ): Promise<readonly WorkspaceCellOutcome[]> {
    const event = await this.#run((id) => ({
      type: "updateCells",
      id,
      generation: gesture.generation,
      table: gesture.table,
      writes: gesture.writes,
    }));
    if (event.type !== "cellsUpdated") {
      throw this.#unexpected(event);
    }
    return event.results.map((result) =>
      result.accepted
        ? {
            accepted: true,
            recordId: result.recordId,
            column: result.column,
            value: result.value,
          }
        : {
            accepted: false,
            recordId: result.recordId,
            column: result.column,
            error: this.#rebuild(result.message, result.code),
          },
    );
  }

  /**
   * Tear down the worker entirely, failing anything still pending, and refuse
   * every command from now on. Commands still waiting in the queue when this
   * runs reach `#ensureWorker` only after the pending one is rejected, so the
   * flag is what stops them from creating a worker the page has already left
   * behind. Calling this more than once is harmless.
   */
  terminate(): void {
    this.#terminated = true;
    this.#worker?.terminate();
    this.#worker = null;
    this.#failEveryPending(CLOSED_REASON);
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

  /**
   * Rebuild a failure that travelled as a message and a code. A
   * `ConsultChimpsError` cannot be structure-cloned as itself, and this bundle's
   * class is what the shared error formatting tests for.
   */
  #rebuild(message: string, code: string | undefined): Error {
    return code === undefined
      ? new Error(message)
      : new ConsultChimpsError(code, message);
  }

  #ensureWorker(): Worker {
    if (this.#terminated) {
      throw new ConsultChimpsError(WORKSPACE_WORKER_UNAVAILABLE, CLOSED_REASON);
    }
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
      pending.reject(this.#rebuild(event.message, event.code));
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
