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
import {
  Database,
  identifierKey,
  importTables,
  RECORD_ID_COLUMN,
} from "@consultchimps/db";
import { isConsultChimpsError } from "@consultchimps/core";
import type { CellValue } from "@consultchimps/tabular";

import type { WorkspaceImportKind } from "@/lib/accepted-files";
import { basePath } from "@/lib/shared";
import { HeldWorkspace } from "@/lib/workspace-generation";
import {
  WORKSPACE_REFERENCE_LIMIT,
  type ImportTableChoice,
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

/**
 * The one workspace this worker owns, and which one it is. Every grid read and
 * write names a workspace and is checked against this, so a command posted
 * before a create, open, close, or import and delivered after it is refused
 * rather than applied to the database that took its place. The database is only
 * reachable through this holder, so there is no way to change it without the
 * generation moving with it. See `lib/workspace-generation.ts`.
 */
const held = new HeldWorkspace<Database>();

function summarize(current: Database): WorkspaceSummary {
  const schema = current.getSchema();
  return {
    // Read here rather than asked for separately, so the table listing and the
    // generation a grid quotes back always describe the same database.
    generation: held.generation,
    tableCount: schema.length,
    schemaFormatVersion: current.schemaFormatVersion(),
    // Counted in the engine rather than by reading the rows, so listing a large
    // workspace stays cheap.
    tables: schema.map((table) => ({
      name: table.name,
      rowCount: current.countRecords(table.name),
      recordIdPrefix: table.recordId.prefix,
      recordIdPadding: table.recordId.padding,
      columns: table.columns.map((column) => ({
        name: column.name,
        type: column.type,
      })),
    })),
  };
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
  held.replace(created);
  scope.postMessage({ type: "ready", id, summary: summarize(created) });
}

async function handleOpen(id: number, buffer: ArrayBuffer): Promise<void> {
  // The buffer was transferred in, so this view owns it outright.
  const opened = await Database.open(new Uint8Array(buffer), engineConfig);
  held.replace(opened);
  scope.postMessage({ type: "ready", id, summary: summarize(opened) });
}

function handleSerialize(id: number): void {
  const bytes = held.require().serialize();
  // Copy into an exactly sized buffer so the transfer moves only the workspace
  // bytes, never a larger pool the view might sit inside, and so the worker's
  // own database keeps a live buffer to serialize again.
  const copy = bytes.slice();
  scope.postMessage({ type: "serialized", id, buffer: copy.buffer }, [
    copy.buffer,
  ]);
}

function handleClose(id: number): void {
  held.release();
  scope.postMessage({ type: "closed", id });
}

/* ---------------------------------------------------------------------------
 * Import. The file is read here and the tables are created by the library, so
 * the worker only carries bytes between the two.
 * ------------------------------------------------------------------------ */

async function handleDescribeImport(
  id: number,
  fileName: string,
  kind: WorkspaceImportKind,
  buffer: ArrayBuffer,
): Promise<void> {
  // Describing a file touches no workspace, but the page only offers it once
  // one is open, and refusing here keeps that promise true whatever calls it.
  held.require();
  const { describeImportSources } = await import("@/lib/workspace-import-file");
  const sources = await describeImportSources(
    fileName,
    kind,
    new Uint8Array(buffer),
  );
  scope.postMessage({ type: "importSources", id, sources });
}

async function handleImport(
  id: number,
  fileName: string,
  kind: WorkspaceImportKind,
  buffer: ArrayBuffer,
  tables: readonly ImportTableChoice[],
): Promise<void> {
  const current = held.require();
  const { resolveImportRequests } = await import("@/lib/workspace-import-file");
  const requests = await resolveImportRequests(
    fileName,
    kind,
    new Uint8Array(buffer),
    tables,
  );
  // One call for every chosen table, so a failure part way through leaves the
  // workspace exactly as it was rather than half imported.
  const imported = importTables(current, requests);
  // The database now holds tables a grid opened before this import never saw,
  // and its snapshot of the rows it does show may be short of imported ones.
  // Recording the change here is what makes that snapshot stale rather than
  // merely out of date: the worker refuses its edits, and the new summary sends
  // the grid back for a fresh read. It runs only once the import has succeeded,
  // so a refused import leaves every open grid still valid.
  held.changed();
  scope.postMessage({
    type: "imported",
    id,
    summary: summarize(current),
    tables: imported.map((table) => ({
      name: table.name,
      rowCount: table.rowCount,
      recordIdPrefix: table.recordId.prefix,
      firstRecordId: table.firstRecordId,
      lastRecordId: table.lastRecordId,
      ignoredColumns: table.ignoredColumns,
      renamedColumns: table.renamedColumns,
    })),
  });
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
  // Read only the two columns an option needs, and one record past the cap:
  // that single extra row says whether the table holds more than is offered,
  // without counting or materialising the rest of a large referenced table.
  const rows = current.readRecords(schema.name, {
    columns: labelColumn === undefined ? [] : [labelColumn.name],
    limit: WORKSPACE_REFERENCE_LIMIT + 1,
  });
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
  held.assertCurrent(generation, "read");
  const current = held.require();
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
    generation: held.generation,
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
  held.assertCurrent(generation, "edit");
  const current = held.require();
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
    case "describeImport":
      return handleDescribeImport(
        command.id,
        command.fileName,
        command.kind,
        command.buffer,
      );
    case "import":
      return handleImport(
        command.id,
        command.fileName,
        command.kind,
        command.buffer,
        command.tables,
      );
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
