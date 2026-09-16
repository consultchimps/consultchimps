import { ConsultChimpsError } from "@consultchimps/core";
import type {
  Database as WasmDatabase,
  Sqlite3Static,
} from "@sqlite.org/sqlite-wasm";
import {
  limitError,
  readError,
  validateReadOptions,
} from "./sqlite-read/options.js";
import type { SqliteReadOptions } from "./sqlite-read/options.js";
import { runReadQuery } from "./sqlite-read/query.js";
import type { SqliteReadResult, SqliteReadValue } from "./sqlite-read/query.js";

export type {
  SqliteReadOptions,
  SqliteReadRuntimeConfig,
} from "./sqlite-read/options.js";
export type { SqliteReadResult, SqliteReadValue } from "./sqlite-read/query.js";

/** An ephemeral reader, independent of managed databases and persistent storage. */
export interface ReadOnlySqlite {
  query(sql: string, parameters?: readonly SqliteReadValue[]): SqliteReadResult;
  /** Current linear memory reservation, for the host's pipeline accounting. */
  readonly wasmMemoryBytes: number;
  close(): void;
}

/**
 * Open an independent, read-only copy of SQLite bytes using the official WASM
 * runtime. Each call owns a fresh runtime; Node uses the installed WASM by
 * default, while browser callers supply a locator or the binary itself.
 */
export async function openReadOnlySqlite(
  bytes: Uint8Array,
  options: SqliteReadOptions = {},
): Promise<ReadOnlySqlite> {
  const limits = validateReadOptions(options);
  if (!(bytes instanceof Uint8Array))
    throw readError(
      "DB_SQLITE_READ_INVALID_DATABASE",
      "open",
      "Supply SQLite database bytes from a readable catalog.",
    );
  if (bytes.byteLength > limits.maxDatabaseBytes)
    throw limitError("maxDatabaseBytes", limits.maxDatabaseBytes);
  const header = "SQLite format 3\0";
  if (
    bytes.length < 100 ||
    !Array.from(header).every(
      (character, index) => bytes[index] === character.charCodeAt(0),
    )
  ) {
    throw readError(
      "DB_SQLITE_READ_INVALID_DATABASE",
      "open",
      "The supplied bytes are not a readable SQLite database. Obtain a valid catalog and try again.",
    );
  }
  const source = options.runtime;
  if (source?.wasmBinary && source.wasmBinary.length > limits.maxSqliteBytes)
    throw limitError("maxSqliteBytes", limits.maxSqliteBytes);
  let sqlite: Sqlite3Static | undefined;
  try {
    const { default: initialize } = await import("@sqlite.org/sqlite-wasm");
    // Upstream's default-export declaration omits Emscripten module options.
    const load = initialize as unknown as (config: {
      wasmBinary?: Uint8Array;
      locateFile?: (fileName: string) => string;
      print: () => void;
      printErr: () => void;
    }) => Promise<Sqlite3Static>;
    // The pinned runtime logs its own load failures through console.error
    // before rejecting. Only the thrown error below is under this reader's
    // control, so it carries no cause, path, or upstream text.
    sqlite = await load({
      ...(source?.wasmBinary === undefined
        ? {}
        : { wasmBinary: new Uint8Array(source.wasmBinary) }),
      ...(source?.locateFile === undefined
        ? {}
        : { locateFile: (fileName: string) => source.locateFile!(fileName) }),
      print: () => undefined,
      printErr: () => undefined,
    });
  } catch {
    throw readError(
      "DB_SQLITE_READ_RUNTIME_UNAVAILABLE",
      "load",
      "SQLite could not load its WebAssembly runtime. Install the pinned @sqlite.org/sqlite-wasm peer dependency and configure or serve sqlite3.wasm, or supply its bytes.",
    );
  }
  let database: WasmDatabase | undefined;
  let pointer = 0;
  let closed = false;
  // Closing marks the reader unusable first, then releases the handle before
  // the catalog buffer it still references. A failed step keeps its state so a
  // retry releases exactly what remains; nothing is freed twice.
  const close = (): void => {
    closed = true;
    try {
      if (database !== undefined) {
        database.close();
        database = undefined;
      }
      if (pointer) {
        sqlite!.wasm.dealloc(pointer);
        pointer = 0;
      }
      // Dropping our runtime reference lets its linear memory be collected even
      // when the caller retains the closed reader object.
      sqlite = undefined;
    } catch {
      throw readError(
        "DB_SQLITE_READ_CLEANUP_REQUIRED",
        "close",
        "SQLite could not release its reader. Retry close, or discard the owning worker before another read.",
      );
    }
  };
  try {
    database = new sqlite.oo1.DB(":memory:");
    // The heap ceiling is scoped to this fresh WASM instance, not the workspace.
    database.exec(
      `PRAGMA hard_heap_limit = ${limits.maxSqliteBytes}; PRAGMA temp_store = MEMORY; PRAGMA trusted_schema = OFF; PRAGMA query_only = ON;`,
    );
    if (
      Number(database.selectValue("PRAGMA hard_heap_limit")) !==
      limits.maxSqliteBytes
    )
      throw limitError("maxSqliteBytes", limits.maxSqliteBytes);
    const { capi } = sqlite;
    capi.sqlite3_limit(database, capi.SQLITE_LIMIT_ATTACHED, 0);
    // The runtime helper selects its heap by constructor, so a Node Buffer or
    // another Uint8Array subclass is viewed as a plain Uint8Array first. This
    // creates no copy of the caller's bytes.
    pointer = sqlite.wasm.allocFromTypedArray(
      new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength),
    );
    const rc = capi.sqlite3_deserialize(
      database,
      "main",
      pointer,
      bytes.length,
      bytes.length,
      capi.SQLITE_DESERIALIZE_READONLY,
    );
    if (rc !== capi.SQLITE_OK)
      throw readError(
        "DB_SQLITE_READ_INVALID_DATABASE",
        "open",
        "SQLite could not open the catalog. Obtain a valid, standalone SQLite database and try again.",
      );
    // Parse the schema now, before the caller's query limits apply. The value
    // length limit below would otherwise be measured against the catalog's own
    // CREATE statements and misreport a valid file as corrupt. The allocator
    // ceiling above still bounds this step.
    database.exec("SELECT name FROM sqlite_schema LIMIT 0");
    // maxSqlBytes is enforced on the caller's text alone, in runReadQuery. No
    // connection-level SQL length limit is set: it would also measure the
    // statements a virtual table prepares internally and surface as an
    // unattributable SQLITE_TOOBIG. The value length limit stays on the
    // connection so that every materialized value, including a virtual
    // table's internal reads, is bounded by maxResultBytes.
    capi.sqlite3_limit(
      database,
      capi.SQLITE_LIMIT_LENGTH,
      limits.maxResultBytes,
    );
  } catch (error) {
    const allocationFailed =
      error instanceof sqlite.WasmAllocError ||
      (error instanceof sqlite.SQLite3Error &&
        error.resultCode === sqlite.capi.SQLITE_NOMEM);
    try {
      close();
    } catch {
      // The open failure below is the actionable error. Whatever could not be
      // released belongs to this reader alone and is unreachable afterwards.
    }
    if (
      error instanceof ConsultChimpsError &&
      error.code === "DB_SQLITE_READ_LIMIT_EXCEEDED"
    )
      throw error;
    if (allocationFailed)
      throw limitError("maxSqliteBytes", limits.maxSqliteBytes);
    throw readError(
      "DB_SQLITE_READ_INVALID_DATABASE",
      "open",
      "SQLite could not read the catalog within its configured limits. Obtain a valid, standalone SQLite database and try again.",
    );
  }
  return {
    get wasmMemoryBytes() {
      return sqlite?.wasm.heap8u().byteLength ?? 0;
    },
    query(sql, parameters = []) {
      if (closed || database === undefined)
        throw readError(
          "DB_SQLITE_READ_CLOSED",
          "query",
          "This SQLite reader is closed. Open the catalog again before querying it.",
        );
      try {
        return runReadQuery(
          sqlite!,
          database.pointer!,
          sql,
          parameters,
          limits,
        );
      } catch (error) {
        if (error instanceof ConsultChimpsError) throw error;
        if (error instanceof sqlite!.WasmAllocError)
          throw limitError("maxSqliteBytes", limits.maxSqliteBytes);
        throw readError(
          "DB_SQLITE_READ_QUERY_FAILED",
          "query",
          "SQLite could not read the requested result. Check the query, catalog, and configured limits.",
        );
      }
    },
    close,
  };
}
