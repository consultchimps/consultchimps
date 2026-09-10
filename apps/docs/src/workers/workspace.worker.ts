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
import { Database, importTables } from "@consultchimps/db";
import { ConsultChimpsError, isConsultChimpsError } from "@consultchimps/core";

import type { WorkspaceImportKind } from "@/lib/accepted-files";
import { basePath } from "@/lib/shared";
import type {
  ImportTableChoice,
  WorkspaceCommand,
  WorkspaceEvent,
  WorkspaceSummary,
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

function summarize(current: Database): WorkspaceSummary {
  const schema = current.getSchema();
  return {
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

// Every command that acts on a workspace needs one to be open, and says so the
// same way.
function requireDatabase(): Database {
  if (database === null) {
    throw new ConsultChimpsError(
      "WORKSPACE_NONE_OPEN",
      "There is no workspace open yet. Start a new workspace or open one first.",
    );
  }
  return database;
}

// Replace the held workspace, closing the previous one first so its sql.js
// allocation is never orphaned. Assigned only after the new database is built,
// so a failed open leaves the previous workspace untouched.
function replaceDatabase(next: Database): void {
  if (database !== null) {
    database.close();
  }
  database = next;
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
  const bytes = requireDatabase().serialize();
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
  requireDatabase();
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
  const current = requireDatabase();
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
  }
}

scope.addEventListener("message", (event) => {
  const command = event.data;
  void dispatch(command).catch((error: unknown) => {
    reportError(command.id, error);
  });
});
