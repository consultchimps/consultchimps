import type { Sqlite3Static } from "@sqlite.org/sqlite-wasm";
import { limitError, readError, utf8Bytes } from "./options.js";
import type { ReadLimits } from "./options.js";

export type SqliteReadValue = null | string | number | bigint | Uint8Array;

export interface SqliteReadResult {
  readonly columns: readonly string[];
  readonly rows: readonly (readonly SqliteReadValue[])[];
  /** Accounted result storage, including column names, row slots, and values. */
  readonly resultBytes: number;
}

export function queryError(): Error {
  return readError(
    "DB_SQLITE_READ_QUERY_FAILED",
    "query",
    "The SQLite query could not be read. Use one read-only SELECT statement with valid parameters and a readable catalog.",
  );
}

export function runReadQuery(
  sqlite: Sqlite3Static,
  database: number,
  sql: string,
  parameters: readonly SqliteReadValue[],
  limits: ReadLimits,
): SqliteReadResult {
  if (
    typeof sql !== "string" ||
    sql.includes("\0") ||
    !Array.isArray(parameters)
  )
    throw queryError();
  if (utf8Bytes(sql) > limits.maxSqlBytes)
    throw limitError("maxSqlBytes", limits.maxSqlBytes);
  let parameterBytes = 0;
  for (const value of parameters) {
    if (value === null) parameterBytes += 8;
    else if (typeof value === "string") parameterBytes += utf8Bytes(value);
    else if (value instanceof Uint8Array) parameterBytes += value.byteLength;
    else if (typeof value === "number" && Number.isFinite(value))
      parameterBytes += 8;
    else if (
      typeof value === "bigint" &&
      value >= -(1n << 63n) &&
      value < 1n << 63n
    )
      parameterBytes += 8;
    else throw queryError();
    if (parameterBytes > limits.maxResultBytes)
      throw limitError("maxResultBytes", limits.maxResultBytes);
  }
  const { capi, wasm } = sqlite;
  let steps = 0;
  let denied = false;
  const check = (rc: number): void => {
    if (rc === capi.SQLITE_OK) return;
    if (steps > limits.maxSteps) throw limitError("maxSteps", limits.maxSteps);
    if (rc === capi.SQLITE_NOMEM)
      throw limitError("maxSqliteBytes", limits.maxSqliteBytes);
    if (rc === capi.SQLITE_TOOBIG)
      throw limitError("maxResultBytes", limits.maxResultBytes);
    if (denied || rc === capi.SQLITE_READONLY) {
      throw readError(
        "DB_SQLITE_READ_ONLY",
        "query",
        "This reader permits SELECT queries only. Database writes, attachments, and configuration changes are not available.",
      );
    }
    throw queryError();
  };
  const scope = wasm.scopedAllocPush();
  let statement = 0;
  try {
    capi.sqlite3_progress_handler(
      database,
      1,
      () => (++steps > limits.maxSteps ? 1 : 0),
      0,
    );
    check(
      capi.sqlite3_set_authorizer(
        database,
        (_context, action, detail, assigned) => {
          // FTS5 cursors read the schema generation through this one PRAGMA
          // while opening. Its query form is permitted, case-insensitively;
          // its assignment form and every other PRAGMA stay refused.
          const allowed =
            action === capi.SQLITE_SELECT ||
            action === capi.SQLITE_READ ||
            action === capi.SQLITE_FUNCTION ||
            action === capi.SQLITE_RECURSIVE ||
            (action === capi.SQLITE_PRAGMA &&
              typeof detail === "string" &&
              detail.toLowerCase() === "data_version" &&
              typeof assigned !== "string");
          denied ||= !allowed;
          return allowed ? capi.SQLITE_OK : capi.SQLITE_DENY;
        },
        0,
      ),
    );
    const [text, length] = wasm.scopedAllocCString(sql, true);
    const [out, tail] = wasm.scopedAllocPtr(2);
    let cursor = text;
    // Prepare the whole input before executing. A second statement is refused,
    // including a second SELECT; semicolons in literals remain ordinary text.
    while (cursor < text + length) {
      wasm.pokePtr(out!, 0);
      const rc = capi.sqlite3_prepare_v3(
        database,
        cursor,
        text + length - cursor,
        0,
        out!,
        tail!,
      );
      const next = wasm.peekPtr(out!);
      if (rc !== capi.SQLITE_OK) {
        if (next) capi.sqlite3_finalize(next);
        check(rc);
      }
      if (next) {
        if (statement) {
          capi.sqlite3_finalize(next);
          throw queryError();
        }
        statement = next;
      }
      const after = wasm.peekPtr(tail!);
      if (after <= cursor) throw queryError();
      cursor = after;
    }
    if (!statement) throw queryError();
    // EXPLAIN output is planner internals, not catalog data.
    if (capi.sqlite3_stmt_isexplain(statement)) throw queryError();
    if (!capi.sqlite3_stmt_readonly(statement)) {
      throw readError(
        "DB_SQLITE_READ_ONLY",
        "query",
        "Use a read-only SELECT query with this SQLite reader.",
      );
    }
    if (capi.sqlite3_bind_parameter_count(statement) !== parameters.length)
      throw queryError();
    for (const [index, value] of parameters.entries()) {
      const slot = index + 1;
      let rc: number;
      // Fresh statements already bind NULL; the pinned null binder discards its
      // C return code, so no unchecked binding call is necessary here.
      if (value === null) continue;
      else if (typeof value === "string" || value instanceof Uint8Array) {
        // Use the C pointer signatures: the pinned JS convenience text binder
        // references an undefined pMem variable, and empty blobs need a non-null
        // address to stay blobs rather than becoming SQL NULL.
        const data =
          typeof value === "string" ? new TextEncoder().encode(value) : value;
        const pointer = wasm.scopedAlloc(Math.max(1, data.byteLength));
        wasm.heap8u().set(data, pointer);
        rc =
          typeof value === "string"
            ? capi.sqlite3_bind_text(
                statement,
                slot,
                pointer,
                data.byteLength,
                capi.SQLITE_TRANSIENT,
              )
            : capi.sqlite3_bind_blob(
                statement,
                slot,
                pointer,
                data.byteLength,
                capi.SQLITE_TRANSIENT,
              );
      } else if (typeof value === "bigint")
        rc = capi.sqlite3_bind_int64(statement, slot, value);
      else rc = capi.sqlite3_bind_double(statement, slot, value);
      check(rc);
    }
    let resultBytes = 0;
    const reserve = (size: number): void => {
      if (size > limits.maxResultBytes - resultBytes)
        throw limitError("maxResultBytes", limits.maxResultBytes);
      resultBytes += size;
    };
    const columns: string[] = [];
    const count = capi.sqlite3_column_count(statement);
    for (let column = 0; column < count; column++) {
      const name = capi.sqlite3_column_name(statement, column);
      reserve(32 + 2 * utf8Bytes(name));
      columns.push(name);
    }
    const rows: SqliteReadValue[][] = [];
    for (;;) {
      const rc = capi.sqlite3_step(statement);
      if (rc === capi.SQLITE_DONE) break;
      if (rc !== capi.SQLITE_ROW) check(rc);
      if (rows.length === limits.maxRows)
        throw limitError("maxRows", limits.maxRows);
      reserve(32 + 16 * count);
      const row: SqliteReadValue[] = [];
      for (let column = 0; column < count; column++) {
        const type = capi.sqlite3_column_type(statement, column);
        if (type === capi.SQLITE_TEXT) {
          reserve(32 + 2 * capi.sqlite3_column_bytes(statement, column));
          row.push(capi.sqlite3_column_text(statement, column));
        } else if (type === capi.SQLITE_BLOB) {
          const length = capi.sqlite3_column_bytes(statement, column);
          reserve(32 + length);
          const pointer = capi.sqlite3_column_blob(statement, column);
          row.push(wasm.heap8u().slice(pointer, pointer + length));
        } else if (type === capi.SQLITE_INTEGER) {
          reserve(16);
          row.push(capi.sqlite3_column_int64(statement, column));
        } else if (type === capi.SQLITE_FLOAT) {
          reserve(8);
          row.push(capi.sqlite3_column_double(statement, column));
        } else row.push(null);
      }
      rows.push(row);
    }
    return { columns, rows, resultBytes };
  } finally {
    if (statement) capi.sqlite3_finalize(statement);
    wasm.scopedAllocPop(scope);
    capi.sqlite3_progress_handler(database, 0, 0, 0);
  }
}
