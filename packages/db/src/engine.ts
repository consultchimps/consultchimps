import initSqlJs from "sql.js";
import type {
  BindParams,
  Database as SqlJsDatabase,
  ParamsObject,
  SqlJsConfig,
  SqlJsStatic,
  SqlValue,
} from "sql.js";

/**
 * The sql.js boundary. Nothing else in the package imports sql.js: this module
 * turns its callback-and-statement API into a small typed surface (load bytes,
 * run statements, read rows, serialize back to bytes) so the schema and bridge
 * layers never touch the engine directly and a future engine swap stays local.
 *
 * sql.js compiles SQLite to WebAssembly and runs it wholly in memory. Loading
 * the runtime is async because the wasm must be fetched and instantiated; every
 * operation after that is synchronous. In Node the wasm resolves beside the
 * sql.js package automatically; a browser worker passes `locateFile` (or
 * `wasmBinary`) so the wasm is fetched from wherever the app serves it.
 */

/** A stored cell value: text, a number, binary, or null. */
export type SqlValueType = SqlValue;

/** Bound statement parameters, positional or named. */
export type SqlParams = SqlValue[] | Record<string, SqlValue>;

/** One result row keyed by column name. */
export type SqlRow = ParamsObject;

/** A raw result set from a multi-row query. */
export interface SqlQueryResult {
  columns: string[];
  values: SqlValue[][];
}

/**
 * How to locate the sql.js WebAssembly binary. Omit it in Node, where the wasm
 * resolves beside the installed package; supply `locateFile` or `wasmBinary` in
 * a browser worker.
 */
export interface SqlEngineConfig {
  locateFile?: (fileName: string) => string;
  wasmBinary?: Uint8Array;
}

// The default runtime is loaded once and reused. A run with a custom config
// (a browser worker pointing at its own wasm) is loaded on demand and not
// cached, since the config is what varies.
let defaultRuntime: Promise<SqlJsStatic> | undefined;

async function loadRuntime(config?: SqlEngineConfig): Promise<SqlJsStatic> {
  if (config !== undefined) {
    // sql.js accepts a Uint8Array for `wasmBinary` at runtime even though the
    // Emscripten type narrows it to ArrayBuffer, so widen through `unknown`.
    return initSqlJs(config as unknown as SqlJsConfig);
  }
  defaultRuntime ??= initSqlJs();
  return defaultRuntime;
}

/**
 * A thin, typed handle over one in-memory SQLite database.
 */
export class SqlDatabase {
  readonly #db: SqlJsDatabase;

  constructor(db: SqlJsDatabase) {
    this.#db = db;
  }

  /** Execute a statement that returns no rows (DDL, insert, update, delete). */
  run(sql: string, params?: SqlParams): void {
    this.#db.run(sql, params as BindParams);
  }

  /** Execute one or more statements, returning each result set. */
  exec(sql: string): SqlQueryResult[] {
    return this.#db.exec(sql).map((result) => ({
      columns: result.columns,
      values: result.values,
    }));
  }

  /** Run a query and return its rows as objects keyed by column name. */
  select(sql: string, params?: SqlParams): SqlRow[] {
    const statement = this.#db.prepare(sql, params as BindParams);
    const rows: SqlRow[] = [];
    try {
      while (statement.step()) {
        rows.push(statement.getAsObject());
      }
    } finally {
      statement.free();
    }
    return rows;
  }

  /** Run a query expected to return a single value, or null when it returns no row. */
  selectValue(sql: string, params?: SqlParams): SqlValue {
    const statement = this.#db.prepare(sql, params as BindParams);
    try {
      if (!statement.step()) {
        return null;
      }
      const values = statement.get();
      return values.length > 0 ? values[0]! : null;
    } finally {
      statement.free();
    }
  }

  /** Serialize the whole database back to bytes for saving. */
  serialize(): Uint8Array {
    return this.#db.export();
  }

  /** Release the database and its memory. */
  close(): void {
    this.#db.close();
  }
}

/**
 * Load a database into memory. Pass the bytes of an existing SQLite file to
 * open it, or omit them to create an empty one. Supply `config` in a browser
 * worker to locate the wasm.
 */
export async function loadSqlDatabase(
  bytes?: Uint8Array,
  config?: SqlEngineConfig,
): Promise<SqlDatabase> {
  const runtime = await loadRuntime(config);
  const db =
    bytes !== undefined ? new runtime.Database(bytes) : new runtime.Database();
  return new SqlDatabase(db);
}
