/**
 * The worker that owns the in-browser workspace database.
 *
 * The whole point of this worker is ownership: exactly one `@consultchimps/db`
 * `Database`, and with it the one sql.js WebAssembly instance, lives here for
 * the life of the page. The main thread never touches the engine; it sends the
 * commands in `workspace-protocol.ts` and renders what comes back. Keeping the
 * engine off the main thread is what keeps the tab responsive when a later PR
 * loads a large workspace or runs a query.
 *
 * sql.js needs its wasm located. The docs build copies the wasm into
 * `public/sql-wasm/`, and the export is served under `basePath` (empty locally,
 * "/consultchimps" on GitHub Pages), so `locateFile` builds the URL from that
 * base, the same way the rest of the app prefixes raw fetch targets. Serving the
 * wasm from our own origin, never a CDN, is the local-first rule in
 * docs/adr/0003 Decision 2.
 */
import { Database, identifierKey, RECORD_ID_COLUMN } from "@consultchimps/db";
import { ConsultChimpsError, isConsultChimpsError } from "@consultchimps/core";
import type { CellValue } from "@consultchimps/tabular";

import { basePath } from "@/lib/shared";
import { WorkspaceGenerations } from "@/lib/workspace-generation";
import {
  WORKSPACE_REFERENCE_LIMIT,
  type WorkspaceColumn,
  type WorkspaceCommand,
  type WorkspaceEvent,
  type WorkspaceReference,
  type WorkspaceSummary,
  type WorkspaceTable,
} from "@/lib/workspace-protocol";

/**
 * The worker global, typed locally. Pulling in the `webworker` lib would
 * collide with the DOM lib the rest of the app compiles against, and the two
 * members used here are the whole surface, the same reason the operation worker
 * types its scope this way.
 */
const scope = self as unknown as {
  addEventListener(
    type: "message",
    listener: (event: MessageEvent<WorkspaceCommand>) => void,
  ): void;
  postMessage(message: WorkspaceEvent, transfer?: Transferable[]): void;
};

const engineConfig = {
  locateFile: (file: string): string => `${basePath}/sql-wasm/${file}`,
};

/** The one workspace this worker owns, or null before the first create or open. */
let database: Database | null = null;

/**
 * Which database that is. Every grid read and write names the workspace it
 * belongs to and is checked against this, so a command posted before a create,
 * open, or close and delivered after it is refused rather than applied to the
 * database that took its place. See `lib/workspace-generation.ts`.
 */
const generations = new WorkspaceGenerations();

function summarize(current: Database): WorkspaceSummary {
  return {
    tableCount: current.getSchema().length,
    schemaFormatVersion: current.schemaFormatVersion(),
  };
}

// Replace the held workspace, closing the previous one first so its sql.js
// allocation is never orphaned. Assigned only after the new database is built,
// so a failed open leaves the previous workspace untouched.
function replaceDatabase(next: Database): void {
  if (database !== null) {
    database.close();
  }
  database = next;
  generations.replaced();
}

function reportError(id: number, error: unknown): void {
  scope.postMessage({
    type: "error",
    id,
    message:
      error instanceof Error
        ? error.message
        : "An unexpected problem occurred.",
    code: isConsultChimpsError(error) ? error.code : undefined,
  });
}

async function handleCreate(id: number): Promise<void> {
  const created = await Database.create(engineConfig);
  replaceDatabase(created);
  scope.postMessage({ type: "ready", id, summary: summarize(created) });
}

async function handleOpen(id: number, buffer: ArrayBuffer): Promise<void> {
  // The buffer was transferred in, so this view owns it outright.
  const opened = await Database.open(new Uint8Array(buffer), engineConfig);
  replaceDatabase(opened);
  scope.postMessage({ type: "ready", id, summary: summarize(opened) });
}

function handleSerialize(id: number): void {
  if (database === null) {
    throw new ConsultChimpsError(
      "WORKSPACE_NONE_OPEN",
      "There is no workspace to save yet. Start a new workspace or open one first.",
    );
  }
  const bytes = database.serialize();
  // Copy into an exactly sized buffer so the transfer moves only the workspace
  // bytes, never a larger pool the view might sit inside, and so the worker's
  // own database keeps a live buffer to serialize again.
  const copy = bytes.slice();
  scope.postMessage({ type: "serialized", id, buffer: copy.buffer }, [
    copy.buffer,
  ]);
}

function handleClose(id: number): void {
  if (database !== null) {
    database.close();
    database = null;
  }
  // Releasing the workspace makes everything read from it stale too, so the
  // counter moves here as well as on a replacement.
  generations.replaced();
  scope.postMessage({ type: "closed", id });
}

/* -------------------------------------------------------------------------
 * Record grid
 *
 * The grid holds rows, never a database, so every read and every write it makes
 * lands here. Type enforcement is not repeated: an edit is handed to
 * `Database.updateRecord`, which routes it through the one conversion point the
 * library owns, and a value that does not fit the column comes back as a
 * ConsultChimpsError the page can show.
 * ------------------------------------------------------------------------- */

function requireDatabase(): Database {
  if (database === null) {
    throw new ConsultChimpsError(
      "WORKSPACE_NONE_OPEN",
      "There is no workspace open yet. Start a new workspace or open one first.",
    );
  }
  return database;
}

function handleListTables(id: number): void {
  const current = requireDatabase();
  scope.postMessage({
    type: "tables",
    id,
    // Handed out so every later read and write can name the workspace it came
    // from, and be refused once that workspace is gone.
    generation: generations.current,
    tables: current.getSchema().map((schema) => schema.name),
  });
}

/**
 * The options a foreign-key column offers. The label is the referenced record's
 * first text column, which is the closest thing the schema has to a name, with
 * the Record ID kept alongside it: two customers can share a name, and the
 * value being stored is the Record ID, so showing it is what makes the choice
 * unambiguous. A column with no text to show falls back to the Record ID alone.
 */
function referenceOptions(
  current: Database,
  tableName: string,
): {
  references: WorkspaceReference[];
  truncated: boolean;
} {
  const schema = current.getTableSchema(tableName);
  const foreignKeyColumns = new Set(
    schema.foreignKeys.map((foreignKey) => identifierKey(foreignKey.column)),
  );
  // A foreign key column holds Record IDs, so it names nothing; the first
  // ordinary text column is the readable one.
  const labelColumn = schema.columns.find(
    (column) =>
      column.type === "text" &&
      !foreignKeyColumns.has(identifierKey(column.name)),
  );
  const rows = current.readRecords(schema.name);
  const references = rows
    .slice(0, WORKSPACE_REFERENCE_LIMIT)
    .map((row): WorkspaceReference => {
      const value = String(row[RECORD_ID_COLUMN]);
      const label =
        labelColumn === undefined ? null : (row[labelColumn.name] ?? null);
      return {
        value,
        label:
          typeof label === "string" && label.trim() !== ""
            ? `${label} (${value})`
            : value,
      };
    });
  return { references, truncated: rows.length > references.length };
}

function handleReadTable(id: number, name: string, generation: number): void {
  generations.assertCurrent(generation, "read");
  const current = requireDatabase();
  const schema = current.getTableSchema(name);
  const referencedBy = new Map(
    schema.foreignKeys.map((foreignKey) => [
      identifierKey(foreignKey.column),
      foreignKey.referencesTable,
    ]),
  );
  // One lookup per referenced table, not per column, so two foreign keys onto
  // the same table read it once.
  const optionsByTable = new Map<
    string,
    { references: WorkspaceReference[]; truncated: boolean }
  >();
  const columns = schema.columns.map((column): WorkspaceColumn => {
    const referencesTable = referencedBy.get(identifierKey(column.name));
    if (referencesTable === undefined) {
      return {
        name: column.name,
        type: column.type,
        nullable: column.nullable !== false,
        references: null,
        referencesTruncated: false,
      };
    }
    const key = identifierKey(referencesTable);
    let options = optionsByTable.get(key);
    if (options === undefined) {
      options = referenceOptions(current, referencesTable);
      optionsByTable.set(key, options);
    }
    return {
      name: column.name,
      type: column.type,
      nullable: column.nullable !== false,
      references: options.references,
      referencesTruncated: options.truncated,
    };
  });

  const table: WorkspaceTable = {
    generation: generations.current,
    name: schema.name,
    columns,
    rows: current.readRecords(schema.name),
  };
  scope.postMessage({ type: "table", id, table });
}

function handleUpdateCell(
  id: number,
  generation: number,
  table: string,
  recordId: string,
  column: string,
  value: CellValue,
): void {
  // Before anything else: an edit made in a workspace that has since been
  // replaced names a table and a Record ID that mean something different here,
  // and the Record ID very likely exists in this database too. Refuse it rather
  // than write it to whatever record happens to match.
  generations.assertCurrent(generation, "edit");
  const current = requireDatabase();
  const updated = current.updateRecord(table, recordId, { [column]: value });
  // Read the stored value back under the column name the schema declared, which
  // is what updateRecord keys its reply by, whatever casing the grid sent.
  const [stored] = Object.values(updated.values);
  scope.postMessage({ type: "cellUpdated", id, value: stored ?? null });
}

async function dispatch(command: WorkspaceCommand): Promise<void> {
  switch (command.type) {
    case "create":
      return handleCreate(command.id);
    case "open":
      return handleOpen(command.id, command.buffer);
    case "serialize":
      return handleSerialize(command.id);
    case "close":
      return handleClose(command.id);
    case "listTables":
      return handleListTables(command.id);
    case "readTable":
      return handleReadTable(command.id, command.name, command.generation);
    case "updateCell":
      return handleUpdateCell(
        command.id,
        command.generation,
        command.table,
        command.recordId,
        command.column,
        command.value,
      );
  }
}

scope.addEventListener("message", (event) => {
  const command = event.data;
  void dispatch(command).catch((error: unknown) => {
    reportError(command.id, error);
  });
});
