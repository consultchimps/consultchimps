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
 *
 * The types the grid commands carry are the library's own (`ColumnType`,
 * `CellValue`), not restatements of them, so a column kind the database gains
 * is a compile error here rather than a shape that quietly stops matching.
 */
import type { ColumnType } from "@consultchimps/db";
import type { CellValue, TableRow } from "@consultchimps/tabular";

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
  | CloseWorkspaceCommand
  // The record-grid commands, declared at the foot of this file.
  | ListWorkspaceTablesCommand
  | ReadWorkspaceTableCommand
  | UpdateWorkspaceCellCommand;

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
  | WorkspaceErrorEvent
  // The record-grid events, declared at the foot of this file.
  | WorkspaceTablesEvent
  | WorkspaceTableEvent
  | WorkspaceCellUpdatedEvent;

/* -------------------------------------------------------------------------
 * Record grid
 *
 * Everything below serves the grid: list the tables, read one table's schema
 * and rows, and persist one cell edit. The grid holds no engine of its own, so
 * a cell it shows was read here and a cell it changes is written here.
 *
 * A cell edit is deliberately one command per cell rather than a batched row
 * save. The database is the source of truth for whether a value is acceptable,
 * so the grid needs the answer for the cell the visitor just left, and it needs
 * it before they judge the next one.
 * ------------------------------------------------------------------------- */

/**
 * The most records a foreign-key column offers as pickable options. A workspace
 * table can hold far more rows than anyone can scroll, and every option is
 * built, cloned across the worker boundary, and held by the editor, so the list
 * is capped rather than allowed to grow with the referenced table. Beyond the
 * cap the column stays editable by typing a Record ID, and the database is what
 * decides whether that Record ID exists.
 */
export const WORKSPACE_REFERENCE_LIMIT = 500;

/** One option a foreign-key column offers: what is stored, and what is shown. */
export interface WorkspaceReference {
  /** The referenced Record ID, which is the value the cell stores. */
  readonly value: string;
  /** The readable label shown in the cell and in the picker. */
  readonly label: string;
}

/** A column as the grid needs it: enough to render it and choose an editor. */
export interface WorkspaceColumn {
  readonly name: string;
  readonly type: ColumnType;
  /** Whether the column accepts an empty cell. */
  readonly nullable: boolean;
  /**
   * For a foreign-key column, the records it may point at, at most
   * `WORKSPACE_REFERENCE_LIMIT` of them. Null for every other column.
   */
  readonly references: readonly WorkspaceReference[] | null;
  /** Whether the referenced table holds more records than `references` lists. */
  readonly referencesTruncated: boolean;
}

/** One table of the workspace: its columns and all of its rows. */
export interface WorkspaceTable {
  /** The table name as the schema declares it, not as it was asked for. */
  readonly name: string;
  /** The user columns, in schema order. The Record ID is not among them. */
  readonly columns: readonly WorkspaceColumn[];
  /**
   * Every record in insertion order, keyed by column name, each carrying its
   * Record ID under `RECORD_ID_COLUMN`.
   */
  readonly rows: readonly TableRow[];
}

/** Name the tables the held workspace contains. */
export interface ListWorkspaceTablesCommand {
  readonly type: "listTables";
  readonly id: number;
}

/** Read one table's columns and rows. */
export interface ReadWorkspaceTableCommand {
  readonly type: "readTable";
  readonly id: number;
  readonly name: string;
}

/**
 * Write one cell, found by its Record ID rather than by any position, so a
 * sorted, filtered, or concurrently reloaded grid can never write to the wrong
 * record.
 */
export interface UpdateWorkspaceCellCommand {
  readonly type: "updateCell";
  readonly id: number;
  readonly table: string;
  readonly recordId: string;
  readonly column: string;
  readonly value: CellValue;
}

/** The tables the held workspace contains, ordered by name. */
export interface WorkspaceTablesEvent {
  readonly type: "tables";
  readonly id: number;
  readonly tables: readonly string[];
}

/** One table's columns and rows. */
export interface WorkspaceTableEvent {
  readonly type: "table";
  readonly id: number;
  readonly table: WorkspaceTable;
}

/**
 * A cell was written. The value is what the database now holds, read back
 * through its declared type, which is not always the value that was sent: the
 * text "yes" in a boolean column comes back as `true`. The grid shows this
 * rather than what the visitor typed, so the cell never disagrees with the file
 * that will be saved.
 */
export interface WorkspaceCellUpdatedEvent {
  readonly type: "cellUpdated";
  readonly id: number;
  readonly value: CellValue;
}
