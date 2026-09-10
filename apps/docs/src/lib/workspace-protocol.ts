/**
 * The wire contract between the workspace page and the workspace Web Worker.
 *
 * Unlike the operation worker, which runs one stateless byte task at a time,
 * this worker OWNS a single `@consultchimps/db` `Database` for the life of the
 * page: it is the only holder of the sql.js instance, and the main thread keeps
 * view state alone. Every command therefore acts on that one in-worker database.
 *
 * The protocol started minimal (create, open, serialize, close) and grows by
 * adding command and event variants beside those, never by changing how a
 * command is matched to its reply or how bytes are transferred. Import is the
 * first such addition, in its own block below; queries and a grid follow the
 * same way.
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

import type { WorkspaceImportKind } from "./accepted-files";

/**
 * One column of a table in the workspace, as the shell lists it.
 */
export interface WorkspaceColumnSummary {
  readonly name: string;
  /** The column type the schema declares, which import inferred for it. */
  readonly type: ColumnType;
}

/**
 * One table in the workspace: what it is called, how much it holds, how its
 * records are numbered, and what its columns are. Until a grid exists, this
 * listing is how a person sees that their data arrived.
 */
export interface WorkspaceTableSummary {
  readonly name: string;
  readonly rowCount: number;
  readonly recordIdPrefix: string;
  readonly recordIdPadding: number;
  readonly columns: readonly WorkspaceColumnSummary[];
}

/**
 * What the worker reports about the database it now holds, enough for the shell
 * to show that a workspace is open and describe it without reaching into the
 * engine. It grows as later work exposes more of the schema.
 */
export interface WorkspaceSummary {
  /**
   * Which database the worker holds. It moves on every create, open, close, and
   * import, and the grid quotes it back on every read and write so the worker
   * can refuse one meant for a workspace it no longer has. See
   * `WORKSPACE_STALE_EDIT`.
   */
  readonly generation: number;
  /** The number of user tables in the workspace. An empty workspace has none. */
  readonly tableCount: number;
  /** The schema format version stored in the file, for display and support. */
  readonly schemaFormatVersion: number;
  /** Every table in the workspace, ordered by name. */
  readonly tables: readonly WorkspaceTableSummary[];
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

/* ---------------------------------------------------------------------------
 * Import: describing what a chosen file could contribute, and creating it.
 * ------------------------------------------------------------------------ */

/**
 * One table a chosen file could become: a worksheet of a workbook, or the whole
 * file for delimited text. The suggestions are starting points the page shows
 * in an editable form; nothing is created until an import command names them.
 */
export interface ImportSourceDescription {
  /** The worksheet name, or the file name for delimited text. */
  readonly name: string;
  /** Rows below the header row. */
  readonly rowCount: number;
  /** Columns in the header row. */
  readonly columnCount: number;
  /**
   * Cells holding a formula the workbook carries no calculated value for. Any
   * number above zero means the worksheet cannot be imported: those cells read
   * as empty, so importing would quietly leave holes where the values belong.
   */
  readonly uncachedFormulaCells: number;
  /**
   * Cells holding an error value. Any number above zero means the worksheet
   * cannot be imported: those cells read as the internal code Excel numbers
   * each error by, so importing would store numbers nobody entered.
   */
  readonly errorCells: number;
  /** A safe table name derived from the source name. */
  readonly suggestedTableName: string;
  /** A Record ID prefix derived from that table name. */
  readonly suggestedRecordIdPrefix: string;
}

/** One table the visitor asked for, with the names they settled on. */
export interface ImportTableChoice {
  /** The `ImportSourceDescription.name` this table comes from. */
  readonly source: string;
  readonly tableName: string;
  readonly recordIdPrefix: string;
  readonly recordIdPadding: number;
}

/**
 * List what a `.xlsx`, `.xlsm`, or `.csv` file could contribute, without
 * touching the workspace. The buffer is transferred to the worker, so the caller
 * must not read it afterward.
 */
export interface DescribeImportCommand {
  readonly type: "describeImport";
  readonly id: number;
  readonly fileName: string;
  /**
   * Which family the file belongs to, decided once by the shared accepted-files
   * contract where both the media type the browser reported and the name were
   * available. The worker is handed the answer rather than working it out again
   * from the name, which is all that reaches it.
   */
  readonly kind: WorkspaceImportKind;
  readonly buffer: ArrayBuffer;
}

/**
 * Create the chosen tables in the held workspace and fill them from the file.
 * The file travels again rather than being held between the two commands, so
 * the worker keeps no state beyond the one database it owns.
 */
export interface ImportCommand {
  readonly type: "import";
  readonly id: number;
  readonly fileName: string;
  /** The family, as `DescribeImportCommand` explains. */
  readonly kind: WorkspaceImportKind;
  readonly buffer: ArrayBuffer;
  readonly tables: readonly ImportTableChoice[];
}

export type WorkspaceCommand =
  | CreateWorkspaceCommand
  | OpenWorkspaceCommand
  | SerializeWorkspaceCommand
  | CloseWorkspaceCommand
  | DescribeImportCommand
  | ImportCommand
  // The record-grid commands, declared at the foot of this file.
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

/** What the chosen file could contribute, in the order the file lists it. */
export interface WorkspaceImportSourcesEvent {
  readonly type: "importSources";
  readonly id: number;
  readonly sources: readonly ImportSourceDescription[];
}

/** One table an import created. */
export interface ImportedTableSummary {
  readonly name: string;
  readonly rowCount: number;
  readonly recordIdPrefix: string;
  /** The first and last generated Record ID, absent for an empty table. */
  readonly firstRecordId: string | null;
  readonly lastRecordId: string | null;
  /**
   * Columns of the source that were not created. Only a Record ID column can
   * appear here, and the page says so rather than reporting a clean import of
   * data it quietly left behind.
   */
  readonly ignoredColumns: readonly string[];
  /**
   * Columns stored under a different name from the one the file wrote, because
   * the header was longer than a name may be, or collided with another once it
   * had been shortened. Every value is still there; only the name changed, and
   * saying so is the difference between a report and a claim.
   */
  readonly renamedColumns: ReadonlyArray<{
    readonly from: string;
    readonly to: string;
  }>;
}

/**
 * The import finished. The summary is the workspace as it now stands, so the
 * shell replaces its table listing from one reply rather than asking again.
 */
export interface WorkspaceImportedEvent {
  readonly type: "imported";
  readonly id: number;
  readonly summary: WorkspaceSummary;
  readonly tables: readonly ImportedTableSummary[];
}

export type WorkspaceEvent =
  | WorkspaceReadyEvent
  | WorkspaceSerializedEvent
  | WorkspaceClosedEvent
  | WorkspaceErrorEvent
  | WorkspaceImportSourcesEvent
  | WorkspaceImportedEvent
  // The record-grid events, declared at the foot of this file.
  | WorkspaceTableEvent
  | WorkspaceCellUpdatedEvent;

/* -------------------------------------------------------------------------
 * Record grid
 *
 * Everything below serves the grid: read one table's schema and rows, and
 * persist one cell edit. The grid holds no engine of its own, so a cell it
 * shows was read here and a cell it changes is written here. Which tables it
 * may show is not asked for separately: that is the workspace summary the shell
 * already holds, so the listing and the rows cannot disagree about what exists.
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

/**
 * Raised when a command names a workspace the worker no longer holds.
 *
 * The rows a grid shows came from one database. Creating, opening, or closing a
 * workspace replaces that database, and the commands already in the queue do not
 * know it: a cell edit posted after an open would otherwise be applied to the
 * new database, against a table and Record ID that mean something else there, or
 * to a record that happens to carry the same Record ID. Every read and write
 * therefore names the workspace it belongs to, and the worker refuses the ones
 * that do not match instead of applying them somewhere they were never meant to
 * go. The page disables editing while it is busy, which is why a visitor should
 * never meet this; the check is what makes that true rather than likely.
 */
export const WORKSPACE_STALE_EDIT = "WORKSPACE_STALE_EDIT";

/** The read equivalent: the rows asked for belong to a replaced workspace. */
export const WORKSPACE_STALE_READ = "WORKSPACE_STALE_READ";

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
  /** Which database these rows came from. See `WORKSPACE_STALE_EDIT`. */
  readonly generation: number;
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

/** Read one table's columns and rows, from the workspace `generation` names. */
export interface ReadWorkspaceTableCommand {
  readonly type: "readTable";
  readonly id: number;
  readonly name: string;
  readonly generation: number;
}

/**
 * Write one cell, found by its Record ID rather than by any position, so a
 * sorted, filtered, or concurrently reloaded grid can never write to the wrong
 * record, and named against the workspace the row was read from, so a replaced
 * database can never be written to by an edit meant for the previous one.
 */
export interface UpdateWorkspaceCellCommand {
  readonly type: "updateCell";
  readonly id: number;
  readonly generation: number;
  readonly table: string;
  readonly recordId: string;
  readonly column: string;
  readonly value: CellValue;
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
