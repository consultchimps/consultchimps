import { ConsultChimpsError } from "@consultchimps/core";
import type { CellValue } from "@consultchimps/tabular";

import type { SqlValueType } from "./engine.js";

/**
 * The schema and stable-identifier model. It describes tables, their columns,
 * and the foreign keys between them, and it defines the per-table Record ID: a
 * human-readable, always-generated, immutable stable key that foreign keys
 * reference. The model, its per-table identifier configuration, and each table's
 * next-id counter are persisted inside the database file itself, so a reopened
 * file knows its own schema and where its identifiers left off.
 */

/** The supported column value kinds, each mapped to a SQLite storage class. */
export type ColumnType = "text" | "integer" | "real" | "boolean" | "date";

/** A user-defined column. The Record ID column is reserved and never declared here. */
export interface ColumnDefinition {
  name: string;
  type: ColumnType;
  /** Whether the column accepts null. Defaults to true. */
  nullable?: boolean;
}

/**
 * A foreign key: a text column in this table whose values are Record IDs of the
 * referenced table. Foreign keys always target the referenced table's Record ID
 * column, never an ordinary column.
 */
export interface ForeignKey {
  column: string;
  referencesTable: string;
}

/**
 * The per-table Record ID configuration. An identifier is `prefix`, then the
 * `separator`, then the counter zero-padded to `padding` digits, for example
 * `CUST-0001` with prefix `CUST`, separator `-`, and padding `4`.
 */
export interface RecordIdConfig {
  prefix: string;
  padding: number;
  /** Defaults to "-". */
  separator?: string;
}

/** A table's full schema, as declared and as read back from a database file. */
export interface TableSchema {
  name: string;
  columns: ColumnDefinition[];
  foreignKeys: ForeignKey[];
  recordId: RecordIdConfig;
}

/** The reserved column that holds every record's stable Record ID. */
export const RECORD_ID_COLUMN = "record_id";

/** Prefix reserved for this package's own metadata tables. */
export const RESERVED_TABLE_PREFIX = "_consultchimps";

/** Key/value metadata table (schema-format version and similar). */
export const METADATA_TABLE = "_consultchimps_meta";

/** Per-table registry: definition JSON and the next-id counter. */
export const TABLE_REGISTRY_TABLE = "_consultchimps_tables";

/** The metadata-format version stamped into a new database file. */
export const SCHEMA_FORMAT_VERSION = 1;

/** Default separator between a Record ID prefix and its number. */
export const DEFAULT_RECORD_ID_SEPARATOR = "-";

/**
 * Format a Record ID from its configuration and a counter value. The counter is
 * the record's position in the per-table sequence, starting at 1.
 */
export function formatRecordId(
  config: RecordIdConfig,
  counter: number,
): string {
  const separator = config.separator ?? DEFAULT_RECORD_ID_SEPARATOR;
  return `${config.prefix}${separator}${String(counter).padStart(config.padding, "0")}`;
}

/**
 * Guard an identifier used as a table or column name. Names are quoted before
 * they reach SQLite, so quoting handles ordinary text, but control characters
 * and embedded double quotes are rejected, and the reserved metadata prefix is
 * off limits so a table can never collide with the package's own bookkeeping.
 */
export function assertSafeIdentifier(
  name: string,
  role: "table" | "column",
): void {
  if (name.trim() === "") {
    throw new ConsultChimpsError(
      "DB_INVALID_IDENTIFIER",
      `A ${role} name cannot be empty.`,
      { details: { role } },
    );
  }
  if (name.length > 200) {
    throw new ConsultChimpsError(
      "DB_INVALID_IDENTIFIER",
      `The ${role} name "${name}" is too long (limit 200 characters).`,
      { details: { role, name } },
    );
  }
  // eslint-disable-next-line no-control-regex -- deliberately rejecting control characters and quotes in identifiers.
  if (/[\u0000-\u001f"]/.test(name)) {
    throw new ConsultChimpsError(
      "DB_INVALID_IDENTIFIER",
      `The ${role} name "${name}" contains a control character or a double quote, which are not allowed.`,
      { details: { role, name } },
    );
  }
  if (name.toLowerCase().startsWith(RESERVED_TABLE_PREFIX)) {
    throw new ConsultChimpsError(
      "DB_RESERVED_IDENTIFIER",
      `The ${role} name "${name}" uses the reserved prefix "${RESERVED_TABLE_PREFIX}".`,
      { details: { role, name } },
    );
  }
}

/**
 * Quote an identifier for interpolation into SQL. Callers must have run
 * `assertSafeIdentifier` first, which guarantees there is no embedded quote to
 * escape.
 */
export function quoteIdentifier(name: string): string {
  return `"${name}"`;
}

/** Map a column type to its SQLite storage class. */
export function sqlStorageClass(type: ColumnType): string {
  switch (type) {
    case "text":
    case "date":
      return "TEXT";
    case "integer":
    case "boolean":
      return "INTEGER";
    case "real":
      return "REAL";
    default: {
      const unexpected: never = type;
      throw new ConsultChimpsError(
        "DB_UNKNOWN_COLUMN_TYPE",
        `Unknown column type "${String(unexpected)}".`,
        { details: { type: unexpected } },
      );
    }
  }
}

/**
 * Convert a stored SQLite value to a tabular cell value, interpreting the
 * column's declared type (booleans come back from 0/1 integers).
 */
export function cellFromSqlValue(
  type: ColumnType,
  value: SqlValueType,
): CellValue {
  if (value === null) {
    return null;
  }
  switch (type) {
    case "boolean":
      return value !== 0 && value !== "0";
    case "integer":
    case "real":
      return typeof value === "number" ? value : Number(value);
    case "text":
    case "date":
      return typeof value === "string" ? value : String(value);
    default: {
      const unexpected: never = type;
      throw new ConsultChimpsError(
        "DB_UNKNOWN_COLUMN_TYPE",
        `Unknown column type "${String(unexpected)}".`,
        { details: { type: unexpected } },
      );
    }
  }
}

/**
 * Convert a tabular cell value to a value SQLite can store, interpreting the
 * column's declared type (booleans store as 0/1 integers).
 */
export function sqlValueFromCell(
  type: ColumnType,
  value: CellValue,
): SqlValueType {
  if (value === null || value === undefined) {
    return null;
  }
  switch (type) {
    case "boolean":
      return value ? 1 : 0;
    case "integer": {
      const numeric = typeof value === "number" ? value : Number(value);
      return Number.isFinite(numeric) ? Math.trunc(numeric) : null;
    }
    case "real": {
      const numeric = typeof value === "number" ? value : Number(value);
      return Number.isFinite(numeric) ? numeric : null;
    }
    case "text":
    case "date":
      return typeof value === "string" ? value : String(value);
    default: {
      const unexpected: never = type;
      throw new ConsultChimpsError(
        "DB_UNKNOWN_COLUMN_TYPE",
        `Unknown column type "${String(unexpected)}".`,
        { details: { type: unexpected } },
      );
    }
  }
}
