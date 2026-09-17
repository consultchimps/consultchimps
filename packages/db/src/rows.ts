import { engineOf, type Database } from "./database.js";
import { databaseError } from "./errors.js";
import { validateRegisteredTables } from "./internal/database-layout.js";
import type { EngineValue } from "./internal/engine.js";
import {
  identifierKey,
  quoteIdentifier,
  RECORD_ID_COLUMN,
  type ColumnType,
  type TableSchema,
} from "./schema.js";

/** A stored value as the row browser shows it: 64-bit integers become text. */
export type TableRowValue = string | number | boolean | null;

export interface TableRowsColumn {
  readonly name: string;
  readonly type: ColumnType;
}

export interface TableRowsPage {
  readonly table: string;
  readonly columns: readonly TableRowsColumn[];
  readonly rows: ReadonlyArray<readonly TableRowValue[]>;
  /** Present when more rows follow; pass it back as `page.cursor`. */
  readonly nextCursor?: string | undefined;
}

export interface TableRowsPageOptions {
  readonly limit?: number | undefined;
  readonly cursor?: string | undefined;
}

const DEFAULT_LIMIT = 50;
const MAX_LIMIT = 200;
const SOURCE_FILE_COLUMN = "_source_file_id";

// SQLite stores booleans as 0 and 1; the column type says what the reader
// meant, so the page shows true and false for both engines. Other engine
// values pass through as the engine returns them, except 64-bit integers,
// which become decimal strings so a page survives structured cloning.
function rowValue(
  value: EngineValue | undefined,
  type: ColumnType,
): TableRowValue {
  if (value === undefined || value === null) return null;
  if (
    type === "boolean" &&
    (typeof value === "number" || typeof value === "bigint")
  ) {
    return value !== 0 && value !== 0n;
  }
  if (typeof value === "bigint") return value.toString();
  if (value instanceof Uint8Array) {
    return `<binary ${String(value.byteLength)} bytes>`;
  }
  return value;
}

function cursorFor(table: TableSchema, recordId: string): string {
  return JSON.stringify([identifierKey(table.name), recordId]);
}

function parseCursor(
  cursor: string | undefined,
  table: TableSchema,
): string | undefined {
  if (cursor === undefined) return undefined;
  let fields: unknown;
  try {
    fields = JSON.parse(cursor);
  } catch {
    throw databaseError("DB_INVALID_CURSOR", "The row cursor is invalid.");
  }
  if (
    !Array.isArray(fields) ||
    fields.length !== 2 ||
    typeof fields[0] !== "string" ||
    typeof fields[1] !== "string"
  ) {
    throw databaseError("DB_INVALID_CURSOR", "The row cursor is invalid.");
  }
  if (fields[0] !== identifierKey(table.name)) {
    throw databaseError(
      "DB_INVALID_CURSOR",
      "The row cursor belongs to a different table.",
      { table: table.name },
    );
  }
  return fields[1];
}

/**
 * Read one page of a managed table's stored rows for a bounded, read-only
 * look at the data. Rows come in Record ID order: shorter IDs first, then
 * text order, the same order the batch history uses, so IDs past the
 * configured padding still follow the ones before them. Columns are the
 * Record ID, the table's declared columns in order, and the source file ID.
 * Nothing is written and no checkpoint is taken. A registered table whose
 * storage no longer matches its declaration is refused the way every other
 * managed read is, with `DB_SCHEMA_DRIFT`.
 */
export async function readTableRows(options: {
  readonly database: Database;
  readonly table: string;
  readonly page?: TableRowsPageOptions | undefined;
}): Promise<TableRowsPage> {
  const limit = options.page?.limit ?? DEFAULT_LIMIT;
  if (!Number.isInteger(limit) || limit < 1 || limit > MAX_LIMIT) {
    throw databaseError(
      "DB_INVALID_PAGE_SIZE",
      `Choose a row page size from 1 to ${String(MAX_LIMIT)}.`,
    );
  }
  const engine = engineOf(options.database);
  return engine.readTransaction(async (transaction) => {
    const registered = await validateRegisteredTables(
      transaction,
      options.database.format,
      { allowExtraColumns: true },
    );
    const table = registered.find(
      (schema) => identifierKey(schema.name) === identifierKey(options.table),
    );
    if (table === undefined) {
      throw databaseError(
        "DB_TABLE_NOT_FOUND",
        "The table does not exist in this database.",
        { table: options.table },
      );
    }
    const after = parseCursor(options.page?.cursor, table);
    const columns: TableRowsColumn[] = [
      { name: RECORD_ID_COLUMN, type: "text" },
      ...table.columns.map((column) => ({
        name: column.name,
        type: column.type,
      })),
      { name: SOURCE_FILE_COLUMN, type: "text" },
    ];
    const recordId = quoteIdentifier(RECORD_ID_COLUMN);
    const selected = columns.map((column) => quoteIdentifier(column.name));
    const rows = await transaction.query(
      `SELECT ${selected.join(", ")} FROM ${quoteIdentifier(table.name)}${
        after === undefined
          ? ""
          : ` WHERE LENGTH(${recordId}) > LENGTH(?) OR (LENGTH(${recordId}) = LENGTH(?) AND ${recordId} > ?)`
      } ORDER BY LENGTH(${recordId}), ${recordId} LIMIT ?`,
      after === undefined
        ? [BigInt(limit + 1)]
        : [after, after, after, BigInt(limit + 1)],
    );
    const page = rows.slice(0, limit);
    const last = page.at(-1)?.[RECORD_ID_COLUMN];
    return {
      table: table.name,
      columns,
      rows: page.map((row) =>
        columns.map((column) => rowValue(row[column.name], column.type)),
      ),
      ...(rows.length > limit && typeof last === "string"
        ? { nextCursor: cursorFor(table, last) }
        : {}),
    };
  });
}
