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
  if (identifierKey(name).startsWith(RESERVED_TABLE_PREFIX)) {
    throw new ConsultChimpsError(
      "DB_RESERVED_IDENTIFIER",
      `The ${role} name "${name}" uses the reserved prefix "${RESERVED_TABLE_PREFIX}".`,
      { details: { role, name } },
    );
  }
  // SQLite reserves the "sqlite_" table-name prefix for its own objects and
  // refuses CREATE TABLE on it; reject it here so the caller gets this stable
  // error rather than an unclassified engine exception.
  if (role === "table" && identifierKey(name).startsWith("sqlite_")) {
    throw new ConsultChimpsError(
      "DB_RESERVED_IDENTIFIER",
      `The table name "${name}" uses the "sqlite_" prefix, which SQLite reserves for internal use.`,
      { details: { role, name } },
    );
  }
  // "__proto__" cannot be carried as a plain-object data key: an object literal
  // treats it as the prototype setter and sql.js's getAsObject drops it, so it
  // could never round-trip. Reject it rather than lose data silently. Other
  // prototype names such as "constructor" are fine, because reads are guarded
  // with own-property checks where identifiers become object keys.
  if (name === "__proto__") {
    throw new ConsultChimpsError(
      "DB_RESERVED_IDENTIFIER",
      `The ${role} name "__proto__" is reserved and cannot be used, because it cannot be stored safely as a data key.`,
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

/**
 * The case-insensitive key for an identifier. SQLite treats table and column
 * names case-insensitively, so every identifier comparison in this package goes
 * through this one normalization (and its SQL twin, `... COLLATE NOCASE`), and
 * identifiers are stored and displayed with their original case. The fold is
 * ASCII A to Z only, exactly matching SQLite's `NOCASE` collation, so the
 * JavaScript comparisons here agree with the SQL lookups for every identifier:
 * a non-ASCII pair such as "Ä" and "ä" is the same distinct-name decision on
 * both sides rather than one deciding same and the other different.
 */
export function identifierKey(name: string): string {
  return name.replace(/[A-Z]/g, (letter) => letter.toLowerCase());
}

/**
 * Whether two identifiers denote the same table or column under SQLite's
 * case-insensitive rules.
 */
export function sameIdentifier(a: string, b: string): boolean {
  return identifierKey(a) === identifierKey(b);
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
 * column's declared type (booleans come back from 0/1 integers). SQLite's type
 * affinity does not stop a value that disagrees with the column type from being
 * stored (through the public `sql` handle or an externally edited file), so this
 * validates the stored representation and raises a structured error rather than
 * feed a corrupted value (a stray boolean 2, or a `NaN`) into a table
 * operation. It is the read-side twin of the strict write-side coercion.
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
      if (value === 0) {
        return false;
      }
      if (value === 1) {
        return true;
      }
      throw new ConsultChimpsError(
        "DB_CORRUPT_STORED_VALUE",
        "A boolean column holds a stored value that is not 0 or 1, so the database may be damaged.",
        { details: { type: "boolean" } },
      );
    case "integer":
      // Match the write side: an integer column must hold a whole number within
      // the range representable exactly, not a fractional or rounded value a raw
      // write or external edit could leave behind.
      if (typeof value === "number" && Number.isSafeInteger(value)) {
        return value;
      }
      throw new ConsultChimpsError(
        "DB_CORRUPT_STORED_VALUE",
        "An integer column holds a stored value that is not a whole number within the range that can be represented exactly, so the database may be damaged.",
        { details: { type } },
      );
    case "real":
      if (typeof value === "number" && Number.isFinite(value)) {
        return value;
      }
      throw new ConsultChimpsError(
        "DB_CORRUPT_STORED_VALUE",
        "A number column holds a stored value that is not a finite number, so the database may be damaged.",
        { details: { type } },
      );
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

// The text spellings accepted for a boolean column, so an imported "false" or
// "0" is not silently stored as true by JavaScript truthiness. An empty string
// is treated as an unset value (null); anything else is rejected rather than
// guessed.
const TRUE_TEXT = new Set(["true", "t", "yes", "y", "1"]);
const FALSE_TEXT = new Set(["false", "f", "no", "n", "0"]);

function booleanToSqlValue(value: Exclude<CellValue, null>): SqlValueType {
  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }
  if (typeof value === "number") {
    // Only 0 and 1 are unambiguous; a stray 2 or NaN is rejected rather than
    // coerced to true.
    if (value === 0) {
      return 0;
    }
    if (value === 1) {
      return 1;
    }
    throw new ConsultChimpsError(
      "DB_INVALID_BOOLEAN",
      "A boolean column received a number that is not 0 or 1.",
      { details: { type: "boolean" } },
    );
  }
  const normalized = value.trim().toLowerCase();
  if (normalized === "") {
    return null;
  }
  if (TRUE_TEXT.has(normalized)) {
    return 1;
  }
  if (FALSE_TEXT.has(normalized)) {
    return 0;
  }
  // The offending value is deliberately left out of the message and details:
  // it is imported cell content and may be confidential. The caller adds the
  // table and column context.
  throw new ConsultChimpsError(
    "DB_INVALID_BOOLEAN",
    "A boolean column received a value that is not one of true, false, yes, no, 1, or 0.",
    { details: { type: "boolean" } },
  );
}

// The offending value is always left out of these errors: it is imported cell
// content and may be confidential; the caller adds table and column context.
function invalidNumber(message: string, integer: boolean): ConsultChimpsError {
  return new ConsultChimpsError("DB_INVALID_NUMBER", message, {
    details: { type: integer ? "integer" : "real" },
  });
}

// Parse a numeric column value. A blank string is an intentionally empty cell
// (null); any other value that is not a valid number for the column type is
// rejected rather than silently discarded or rounded.
function numberToSqlValue(
  value: Exclude<CellValue, null>,
  integer: boolean,
): SqlValueType {
  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed === "") {
      return null;
    }
    if (integer) {
      // Validate the text itself before any numeric conversion: Number() would
      // round "1.00000000000000001" or "9007199254740993" to a safe integer and
      // hide the change. A digit string plus a range check on the exact value
      // (via BigInt) keeps the original meaning.
      if (!/^[+-]?\d+$/.test(trimmed)) {
        throw invalidNumber(
          "An integer column received a value that is not a whole number.",
          true,
        );
      }
      const exact = BigInt(trimmed);
      if (
        exact < BigInt(Number.MIN_SAFE_INTEGER) ||
        exact > BigInt(Number.MAX_SAFE_INTEGER)
      ) {
        throw invalidNumber(
          "An integer column received a whole number outside the range that can be represented exactly.",
          true,
        );
      }
      return Number(exact);
    }
    const numeric = Number(trimmed);
    if (!Number.isFinite(numeric)) {
      throw invalidNumber(
        "A number column received a value that is not a finite number.",
        false,
      );
    }
    return numeric;
  }

  // A numeric value. JavaScript has already rounded any literal beyond its safe
  // range before this point, so a safe-integer / finite check is the best the
  // number type allows.
  if (integer) {
    if (!Number.isSafeInteger(value)) {
      throw invalidNumber(
        Number.isFinite(value)
          ? "An integer column received a value that is not a whole number within the range that can be represented exactly."
          : "An integer column received a value that is not a finite number.",
        true,
      );
    }
    return value;
  }
  if (!Number.isFinite(value)) {
    throw invalidNumber(
      "A number column received a value that is not a finite number.",
      false,
    );
  }
  return value;
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
      return booleanToSqlValue(value);
    case "integer":
      return numberToSqlValue(value, true);
    case "real":
      return numberToSqlValue(value, false);
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
