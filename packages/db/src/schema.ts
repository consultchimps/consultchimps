import { ConsultChimpsError, isConsultChimpsError } from "@consultchimps/core";
import type { CellValue } from "@consultchimps/tabular";

import type { SqlValueType } from "./engine.js";

/**
 * The schema and stable-identifier model. It describes tables, their columns,
 * and the foreign keys between them, and it defines the per-table Record ID: a
 * human-readable, always-generated, immutable stable key that foreign keys
 * reference. The model, its per-table identifier configuration, and each table's
 * next-id counter are persisted inside the database file itself, so a reopened
 * file knows its own schema and where its identifiers left off.
 *
 * This module holds no engine and is published on its own as
 * `@consultchimps/db/schema`. A browser page that only has to decide whether
 * two table names are the same name, or what a padding may be, would otherwise
 * download a WebAssembly database to ask, and it must not: the identifier rules
 * belong wherever a name is typed, and the engine belongs only where rows are.
 */

/**
 * Every supported column value kind, in one list.
 *
 * The type below is derived from it, and the check that reads a stored schema
 * back reads the same list, so a kind added here is a kind both of them know
 * about. Kept as two lists, one of them would eventually stop matching the
 * other, and a column type this package supports would be rejected as unknown
 * on the way back in.
 */
export const COLUMN_TYPES = [
  "text",
  "integer",
  "real",
  "boolean",
  "date",
] as const;

/**
 * The supported column value kinds, each mapped to a SQLite storage class.
 *
 * A `date` column holds ISO 8601 text and nothing else: a calendar date
 * (`2026-01-31`) or a date and time (`2026-01-31T09:30:00Z`), stored as written
 * once its surrounding spaces are removed. Anything else is refused rather than
 * stored, because a column that claims a spelling and holds another is a column
 * nothing can read back reliably. `isIsoDateText` is the rule.
 */
export type ColumnType = (typeof COLUMN_TYPES)[number];

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

/**
 * The detail every failure to convert one cell value carries.
 *
 * The wrapper that adds the table and column to such a failure used to name the
 * codes it knew about, so a conversion added later fell straight through it and
 * reached the caller saying only which type had complained. A list of codes kept
 * in step by hand is a list that stops being in step. This is the one question
 * that wrapper asks instead, and `valueConversionError` is the only way to
 * answer yes to it.
 */
const VALUE_CONVERSION_DETAIL = "valueConversion";

/**
 * Raise a failure to read one cell value into a column. The public code stays
 * whatever the caller passes; the marker is what makes it recognisable as a
 * conversion failure rather than a programming error.
 */
function valueConversionError(
  code: string,
  message: string,
  details: Record<string, unknown>,
): ConsultChimpsError {
  return new ConsultChimpsError(code, message, {
    details: { ...details, [VALUE_CONVERSION_DETAIL]: true },
  });
}

/**
 * Whether a failure came from reading a cell value into a column, and so has a
 * table and a column that would tell a person where to look.
 */
export function isValueConversionError(
  error: unknown,
): error is ConsultChimpsError {
  return (
    isConsultChimpsError(error) &&
    error.details?.[VALUE_CONVERSION_DETAIL] === true
  );
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
 * The longest a table or column name may be, measured the way JavaScript
 * measures a string: in UTF-16 code units, so a character outside the basic
 * plane counts as the two units it occupies. The unit matters, because anything
 * that shortens a name to fit has to shorten it in the same unit this is
 * expressed in, which is why `truncateIdentifier` lives beside it.
 */
export const MAX_IDENTIFIER_LENGTH = 200;

/**
 * Shorten text so it fits the identifier limit, cutting only between whole
 * characters.
 *
 * Slicing a string at a code-unit index can land in the middle of a surrogate
 * pair and leave half a character behind. That half is not text: SQLite stores
 * it as U+FFFD, so the name that comes back is not the name that went in, and
 * a suggestion a person accepted turns into something else on the way to the
 * database. Iterating a string yields whole code points, so the cut can only
 * fall between characters.
 */
export function truncateIdentifier(
  text: string,
  limit: number = MAX_IDENTIFIER_LENGTH,
): string {
  if (text.length <= limit) {
    return text;
  }
  let result = "";
  for (const character of text) {
    if (result.length + character.length > limit) {
      break;
    }
    result += character;
  }
  return result;
}

/**
 * A generous but bounded cap on Record ID zero-padding, so a mistaken
 * configuration cannot drive `String.padStart` into an enormous allocation.
 */
export const MAX_RECORD_ID_PADDING = 64;

/**
 * Guard a Record ID configuration before it is stored or used. Table creation
 * and the import path both call this, so a prefix or padding is judged by one
 * rule wherever it arrives from.
 */
export function assertRecordIdConfig(
  table: string,
  config: RecordIdConfig,
): void {
  if (config.prefix.trim() === "") {
    throw new ConsultChimpsError(
      "DB_INVALID_RECORD_ID_CONFIG",
      `The Record ID prefix for "${table}" cannot be empty.`,
      { details: { table } },
    );
  }
  if (
    !Number.isInteger(config.padding) ||
    config.padding < 0 ||
    config.padding > MAX_RECORD_ID_PADDING
  ) {
    throw new ConsultChimpsError(
      "DB_INVALID_RECORD_ID_CONFIG",
      `The Record ID padding for "${table}" must be a whole number from 0 to ${MAX_RECORD_ID_PADDING}.`,
      { details: { table, padding: config.padding } },
    );
  }
}

/**
 * A calendar date, and a date with a time, in the one spelling ISO 8601 gives
 * them. A date written any other way ("01/02/2024") stays text, because there
 * is no way to tell a January date from a February one without guessing a
 * convention the file never stated.
 *
 * The grammar only says which characters may appear where. What the parts mean
 * together is decided in `isIsoTimestamp`, because a timestamp is one value:
 * checking its hour, minute, second, and offset against their own ranges in
 * isolation accepts "24:30" and an offset of "+99:99", neither of which is a
 * time of day.
 */
const ISO_DATE = /^(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})$/u;
const ISO_DATE_TIME =
  /^(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})T(?<hour>\d{2}):(?<minute>\d{2})(?::(?<second>\d{2})(?:\.(?<fraction>\d{1,9}))?)?(?:(?<zulu>Z)|(?<offsetSign>[+-])(?<offsetHour>\d{2}):(?<offsetMinute>\d{2}))?$/u;

function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

function isCalendarDate(year: number, month: number, day: number): boolean {
  if (month < 1 || month > 12 || day < 1) {
    return false;
  }
  const lengths = [
    31,
    isLeapYear(year) ? 29 : 28,
    31,
    30,
    31,
    30,
    31,
    31,
    30,
    31,
    30,
    31,
  ];
  return day <= (lengths[month - 1] as number);
}

/**
 * Whether a matched date and time is a real instant, judged as a whole.
 *
 * Every rule here that could be stated about one part of the value is instead
 * stated about the parts together, because that is where the mistakes are: a
 * value whose hour, minute, second, and offset are each individually in range
 * can still be an instant that does not exist.
 *
 * - Hour 24 is the end of a day, so it is accepted only as exactly `24:00`,
 *   with a zero second and a zero fraction if either is written. `24:30` is
 *   nothing.
 * - Second 60 is a leap second. One is inserted at `23:59:60` UTC, so it is
 *   accepted only at that time and only when the value says it is UTC, written
 *   as `Z` or as the explicit `+00:00`. A local offset and an absent offset
 *   both stay text: `23:59:60+05:00` is a different instant from the leap
 *   second, and a value with no offset does not say which instant it is. This
 *   costs nothing, because a date column stores the characters either way.
 * - An offset is a real offset: hours 00 to 23, minutes 00 to 59, judged as one
 *   offset rather than three independent numbers.
 *
 * The date a leap second falls on is deliberately not constrained. Which day
 * carries one is a decision announced by IERS, not a rule this package can
 * state, so it is left to the value.
 */
function isIsoTimestamp(stamp: RegExpExecArray): boolean {
  // Named rather than numbered: adding a capture to the grammar above used to
  // renumber every read below it, which is its own way of pairing the wrong
  // things together.
  const parts = stamp.groups as Record<string, string | undefined>;
  if (
    !isCalendarDate(
      Number(parts["year"]),
      Number(parts["month"]),
      Number(parts["day"]),
    )
  ) {
    return false;
  }
  const hour = Number(parts["hour"]);
  const minute = Number(parts["minute"]);
  const second = parts["second"] === undefined ? 0 : Number(parts["second"]);
  const fraction = parts["fraction"];
  // Scanned rather than matched with `^0+$`, which backtracks once per digit on
  // a fraction that is not all zeros for no gain over reading it straight
  // through.
  const zeroFraction =
    fraction === undefined || [...fraction].every((digit) => digit === "0");

  // The offset, judged first because the leap-second rule depends on it.
  const hasOffset = parts["offsetSign"] !== undefined;
  if (hasOffset) {
    if (
      Number(parts["offsetHour"]) > 23 ||
      Number(parts["offsetMinute"]) > 59
    ) {
      return false;
    }
  }
  // "+00:00" is UTC stated the long way. "-00:00" is not: RFC 3339 gives it the
  // separate meaning of an unknown local offset, so it does not assert UTC.
  const isUtc =
    parts["zulu"] !== undefined ||
    (parts["offsetSign"] === "+" &&
      Number(parts["offsetHour"]) === 0 &&
      Number(parts["offsetMinute"]) === 0);

  if (hour === 24) {
    if (minute !== 0 || second !== 0 || !zeroFraction) {
      return false;
    }
  } else if (hour > 23) {
    return false;
  }
  if (minute > 59) {
    return false;
  }
  if (second === 60) {
    if (hour !== 23 || minute !== 59 || !isUtc) {
      return false;
    }
  } else if (second > 59) {
    return false;
  }
  return true;
}

/**
 * Whether text is a date this package will store in a `date` column.
 *
 * One definition, used by the import's type inference to decide that a column
 * is a date and by `sqlValueFromCell` to decide that a value may be stored in
 * one. Two definitions would mean a column inferred as a date could reject the
 * very values that made it a date, or accept ones that never would have.
 */
export function isIsoDateText(value: string): boolean {
  const date = ISO_DATE.exec(value);
  if (date !== null) {
    const parts = date.groups as Record<string, string>;
    return isCalendarDate(
      Number(parts["year"]),
      Number(parts["month"]),
      Number(parts["day"]),
    );
  }
  const stamp = ISO_DATE_TIME.exec(value);
  return stamp !== null && isIsoTimestamp(stamp);
}

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
  if (name.length > MAX_IDENTIFIER_LENGTH) {
    throw new ConsultChimpsError(
      "DB_INVALID_IDENTIFIER",
      `The ${role} name "${name}" is too long (limit ${MAX_IDENTIFIER_LENGTH} characters).`,
      { details: { role, name } },
    );
  }
  // An identifier has to be text before it can be anything else. A lone
  // surrogate is half a character: it survives every check below, and SQLite
  // then stores it as U+FFFD, so the name in the database is not the name that
  // was asked for. Refusing it here means every caller inherits the rule,
  // rather than each one having to remember to produce well-formed input.
  if (!name.isWellFormed()) {
    throw new ConsultChimpsError(
      "DB_INVALID_IDENTIFIER",
      `The ${role} name contains an incomplete character, so it is not valid text. Check the encoding of the file it came from.`,
      { details: { role } },
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
    case "date":
      // The write side refuses anything that is not an ISO 8601 date, so a
      // stored value that is not one arrived through the raw `sql` handle or an
      // external edit. Reading it back as a date would carry that through to
      // every caller, so it is reported here instead.
      if (typeof value === "string" && isIsoDateText(value)) {
        return value;
      }
      throw new ConsultChimpsError(
        "DB_CORRUPT_STORED_VALUE",
        "A date column holds a stored value that is not an ISO 8601 date, so the database may be damaged.",
        { details: { type } },
      );
    case "text":
      // A raw write or external edit could leave a BLOB here, which sql.js
      // returns as a Uint8Array; converting it to "65,66" would corrupt the
      // value, so reject a non-string stored value instead.
      if (typeof value === "string") {
        return value;
      }
      throw new ConsultChimpsError(
        "DB_CORRUPT_STORED_VALUE",
        "A text column holds a stored value that is not text, so the database may be damaged.",
        { details: { type } },
      );
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
    throw valueConversionError(
      "DB_INVALID_BOOLEAN",
      "A boolean column received a number that is not 0 or 1.",
      { type: "boolean" },
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
  throw valueConversionError(
    "DB_INVALID_BOOLEAN",
    "A boolean column received a value that is not one of true, false, yes, no, 1, or 0.",
    { type: "boolean" },
  );
}

// The offending value is always left out of these errors: it is imported cell
// content and may be confidential; the caller adds table and column context.
function invalidNumber(message: string, integer: boolean): ConsultChimpsError {
  return valueConversionError("DB_INVALID_NUMBER", message, {
    type: integer ? "integer" : "real",
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
 * Parse a value for a date column.
 *
 * The column claims ISO 8601, so this is where that claim is kept. Surrounding
 * spaces are removed first, because " 2026-01-31 " is the same date written
 * with a spreadsheet's padding and storing the padding would put a value in the
 * column that does not match the spelling the column promises. A blank is an
 * unset cell, like every other typed column here. Anything else is refused: a
 * date column that quietly holds "hello" is a column every later read has to
 * second-guess.
 */
function dateToSqlValue(value: Exclude<CellValue, null>): SqlValueType {
  const text = typeof value === "string" ? value.trim() : String(value);
  if (text === "") {
    return null;
  }
  if (!isIsoDateText(text)) {
    // The offending value is deliberately left out: it is imported cell content
    // and may be confidential. The caller adds the table and column context.
    throw valueConversionError(
      "DB_INVALID_DATE",
      "A date column received a value that is not an ISO 8601 date such as 2026-01-31 or 2026-01-31T09:30:00Z.",
      { type: "date" },
    );
  }
  return text;
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
    case "date":
      return dateToSqlValue(value);
    case "text":
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
