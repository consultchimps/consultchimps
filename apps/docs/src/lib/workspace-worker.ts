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

import type { WorkspaceImportKind } from "./accepted-files";

import type {
  ImportedTableSummary,
  ImportSourceDescription,
  ImportTableChoice,
  WorkspaceCommand,
  WorkspaceEvent,
  WorkspaceSummary,
} from "./workspace-protocol";

/** What an import created, as the page reports it. */
export interface WorkspaceImportResult {
  /** The workspace as it now stands, table listing included. */
  readonly summary: WorkspaceSummary;
  /** The tables this import created. */
  readonly tables: readonly ImportedTableSummary[];
}

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
