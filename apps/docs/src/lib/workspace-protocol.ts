/**
 * The wire contract between the workspace page and the workspace Web Worker.
 *
 * Unlike the operation worker, which runs one stateless byte task at a time,
 * this worker OWNS a single `@consultchimps/db` `Database` for the life of the
 * page: it is the only holder of the sql.js instance, and the main thread keeps
 * view state alone. Every command therefore acts on that one in-worker database.
 *
 * The protocol is deliberately minimal for this first shell (create, open,
 * serialize, close) and shaped to grow: later work adds data import, queries,
 * and a grid as new command and event variants beside these, without changing
 * how a command is matched to its reply or how bytes are transferred.
 *
 * Two constraints shape the shapes below, the same ones the operation protocol
 * meets:
 *
 * - Bytes travel as a transferable `ArrayBuffer`, never a `Uint8Array` view
 *   into a larger pool, so opening or saving a large workspace moves the buffer
 *   instead of copying it.
 * - A `ConsultChimpsError` cannot be structure-cloned as itself, so a failure
 *   travels as its message and code and is rebuilt on the main thread, exactly
 *   as the operation worker does it.
 */

/**
 * What the worker reports about the database it now holds, enough for the shell
 * to show that a workspace is open and describe it without reaching into the
 * engine. It grows as later work exposes more of the schema.
 */
export interface WorkspaceSummary {
  /** The number of user tables in the workspace. An empty workspace has none. */
  readonly tableCount: number;
  /** The schema format version stored in the file, for display and support. */
  readonly schemaFormatVersion: number;
}

/** Start a new, empty workspace, discarding any the worker already holds. */
export interface CreateWorkspaceCommand {
  readonly type: "create";
  readonly id: number;
}

/**
 * Open a workspace from the bytes of a `.sqlite` file. The buffer is transferred
 * to the worker, so the caller must not read it afterward.
 */
export interface OpenWorkspaceCommand {
  readonly type: "open";
  readonly id: number;
  readonly buffer: ArrayBuffer;
}

/** Serialize the held workspace back to bytes for saving. */
export interface SerializeWorkspaceCommand {
  readonly type: "serialize";
  readonly id: number;
}

/** Release the held workspace and its memory. */
export interface CloseWorkspaceCommand {
  readonly type: "close";
  readonly id: number;
}

export type WorkspaceCommand =
  | CreateWorkspaceCommand
  | OpenWorkspaceCommand
  | SerializeWorkspaceCommand
  | CloseWorkspaceCommand;

/**
 * The worker holds a workspace after a create or open. The summary lets the
 * shell render the open state without a follow-up round trip.
 */
export interface WorkspaceReadyEvent {
  readonly type: "ready";
  readonly id: number;
  readonly summary: WorkspaceSummary;
}

/** The serialized workspace bytes, transferred back as a standalone buffer. */
export interface WorkspaceSerializedEvent {
  readonly type: "serialized";
  readonly id: number;
  readonly buffer: ArrayBuffer;
}

/** The held workspace has been released. */
export interface WorkspaceClosedEvent {
  readonly type: "closed";
  readonly id: number;
}

/**
 * A command failed. `code` is present when the failure was a
 * `ConsultChimpsError`, so the main thread can rebuild one and the shared error
 * formatting recognizes it.
 */
export interface WorkspaceErrorEvent {
  readonly type: "error";
  readonly id: number;
  readonly message: string;
  readonly code?: string | undefined;
}

export type WorkspaceEvent =
  | WorkspaceReadyEvent
  | WorkspaceSerializedEvent
  | WorkspaceClosedEvent
  | WorkspaceErrorEvent;
