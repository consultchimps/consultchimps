import { ConsultChimpsError } from "@consultchimps/core";
import type { CellValue, Table } from "@consultchimps/tabular";

import { insertRecordsFromTable } from "./bridge.js";
import type { Database } from "./database.js";
import {
  assertRecordIdConfig,
  assertSafeIdentifier,
  identifierKey,
  RECORD_ID_COLUMN,
  type ColumnDefinition,
  type ColumnType,
  type RecordIdConfig,
  type TableSchema,
} from "./schema.js";

/**
 * Turning a table of values into a database table: what type each column holds,
 * what the table and its Record ID should be called, and the import itself.
 *
 * The logic here is pure and runtime-neutral on purpose. The browser workspace
 * drives it from a Web Worker today and a `db import` command will drive it
 * from a terminal later, and both have to reach the same tables from the same
 * file, so nothing in this module knows about files, workers, or a user
 * interface: a caller hands it a `Table` (from `@consultchimps/xlsx`, from
 * `parseCsvTable`, or built by hand) and it does the rest.
 */

/** How a column's type was decided, alongside the type itself. */
export interface InferredColumn extends ColumnDefinition {
  /** Values that were not blank, and so took part in the decision. */
  readonly valueCount: number;
}

/** What to call an imported table and how to number its records. */
export interface ImportTableOptions {
  /** The table name. Leading and trailing spaces are trimmed. */
  readonly name: string;
  /** The Record ID prefix and padding. The prefix is trimmed. */
  readonly recordId: RecordIdConfig;
}

/** One table to create in a single import. */
export interface ImportTableRequest extends ImportTableOptions {
  readonly table: Table;
}

/** What an import created, enough to tell a person what happened. */
export interface ImportedTable {
  /** The table name as it was stored. */
  readonly name: string;
  /** The columns that were created, with the type inferred for each. */
  readonly columns: InferredColumn[];
  /** The Record ID configuration the table was created with. */
  readonly recordId: RecordIdConfig;
  /** How many records were inserted. */
  readonly rowCount: number;
  /** The first generated Record ID, or null when the table came in empty. */
  readonly firstRecordId: string | null;
  /** The last generated Record ID, or null when the table came in empty. */
  readonly lastRecordId: string | null;
  /**
   * Columns of the input that were not created. Only the reserved Record ID
   * column can appear here: identifiers are always generated, so an input that
   * already carries one (a table read back out of a workspace) is imported
   * without it rather than refused.
   */
  readonly ignoredColumns: string[];
}

/** The Record ID column's matching key, folded once rather than at each use. */
const RECORD_ID_KEY = identifierKey(RECORD_ID_COLUMN);

/**
 * A whole number with no exponent, no leading zeros, and no grouping. The
 * leading-zero rule is what keeps a reference code such as "007" text: read as
 * a number it would come back as 7, and a padded code that loses its padding is
 * a changed value, not a converted one. An exponent form ("1e5") is excluded
 * for the same reason: the file says one thing and the column would say
 * another.
 */
const INTEGER_TEXT = /^-?(?:0|[1-9]\d*)$/u;

/** The same grammar with a required fractional part. */
const DECIMAL_TEXT = /^-?(?:0|[1-9]\d*)\.\d+$/u;

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
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/u;
const ISO_DATE_TIME =
  /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,9}))?)?(?:Z|([+-])(\d{2}):(\d{2}))?$/u;

/**
 * The text spellings read as a boolean. "1" and "0" are deliberately absent:
 * they are numbers in every file that also holds counts, and a column of ones
 * and zeros is far more often a quantity than a flag. A column that really is a
 * flag written that way imports as a number and can be changed afterwards,
 * which is recoverable; guessing the other way is not.
 */
const BOOLEAN_TEXT = new Set(["true", "false", "yes", "no"]);

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
 * The two spellings ISO 8601 allows outside the ordinary ranges are each tied
 * to the rest of the value rather than waved through on their own:
 *
 * - Hour 24 is the end of a day, so it is accepted only as exactly `24:00`,
 *   with a zero second and a zero fraction if either is written. `24:30` is
 *   nothing.
 * - Second 60 is a leap second, which is inserted at `23:59:60` UTC, so it is
 *   accepted only there. A leap second written against a local offset stays
 *   text, which costs nothing: a date column stores the characters either way.
 *
 * An offset is a real offset: hours 00 to 23, minutes 00 to 59.
 */
function isIsoTimestamp(stamp: RegExpExecArray): boolean {
  if (!isCalendarDate(Number(stamp[1]), Number(stamp[2]), Number(stamp[3]))) {
    return false;
  }
  const hour = Number(stamp[4]);
  const minute = Number(stamp[5]);
  const second = stamp[6] === undefined ? 0 : Number(stamp[6]);
  const fraction = stamp[7];
  // Scanned rather than matched with `^0+$`, which backtracks once per digit on
  // a fraction that is not all zeros for no gain over reading it straight
  // through.
  const zeroFraction =
    fraction === undefined || [...fraction].every((digit) => digit === "0");

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
    // A leap second, and only where one is inserted.
    if (hour !== 23 || minute !== 59) {
      return false;
    }
  } else if (second > 59) {
    return false;
  }

  // The offset is present as a whole or not at all: the grammar captures its
  // sign, hours, and minutes together, so one part cannot be checked without
  // the others.
  if (stamp[8] !== undefined) {
    if (Number(stamp[9]) > 23 || Number(stamp[10]) > 59) {
      return false;
    }
  }
  return true;
}

function isIsoDateText(value: string): boolean {
  const date = ISO_DATE.exec(value);
  if (date !== null) {
    return isCalendarDate(Number(date[1]), Number(date[2]), Number(date[3]));
  }
  const stamp = ISO_DATE_TIME.exec(value);
  return stamp !== null && isIsoTimestamp(stamp);
}

/** What a single value could be. A blank value votes for nothing. */
interface ValueVote {
  readonly integer: boolean;
  readonly real: boolean;
  readonly boolean: boolean;
  readonly date: boolean;
}

const NOTHING: ValueVote = {
  integer: false,
  real: false,
  boolean: false,
  date: false,
};

/** A plain or exponent-form decimal, the two shapes `String(number)` produces. */
const DECIMAL_PARTS = /^([+-]?)(\d+)(?:\.(\d+))?(?:[eE]([+-]?\d+))?$/u;

/** Trim leading zeros with a scan, so no regular expression can be quadratic. */
function withoutLeadingZeros(digits: string): string {
  let start = 0;
  while (start < digits.length - 1 && digits[start] === "0") {
    start += 1;
  }
  return digits.slice(start);
}

/** Trim trailing zeros with a scan, for the same reason. */
function withoutTrailingZeros(digits: string): string {
  let end = digits.length;
  while (end > 0 && digits[end - 1] === "0") {
    end -= 1;
  }
  return digits.slice(0, end);
}

/**
 * One canonical spelling for the exact value a decimal denotes, so two
 * spellings of the same number compare equal and two spellings of different
 * numbers do not. "12.50" and "12.5" both give "12.5", and "1e+21" gives its
 * twenty-two digits. Negative zero is zero, which is the one place the value
 * and the characters part company on purpose.
 */
function canonicalDecimal(text: string): string | null {
  const parts = DECIMAL_PARTS.exec(text);
  if (parts === null) {
    return null;
  }
  const exponent = parts[4] === undefined ? 0 : Number(parts[4]);
  const whole = parts[2] as string;
  let digits = whole + (parts[3] ?? "");
  // How many digits sit left of the point once the exponent has moved it.
  let point = whole.length + exponent;
  if (point < 0) {
    digits = "0".repeat(-point) + digits;
    point = 0;
  }
  if (point > digits.length) {
    digits += "0".repeat(point - digits.length);
  }
  const integerPart = withoutLeadingZeros(digits.slice(0, point) || "0");
  const fractionPart = withoutTrailingZeros(digits.slice(point));
  const magnitude =
    fractionPart === "" ? integerPart : `${integerPart}.${fractionPart}`;
  return magnitude === "0" ? "0" : `${parts[1] === "-" ? "-" : ""}${magnitude}`;
}

/**
 * Whether reading this text as a number gives back exactly the value written.
 *
 * Both numeric candidates go through this one predicate, so neither can end up
 * with a weaker rule than the other. Parsing alone is not enough: `Number` is
 * happy to return a finite number for "0.12345678901234567890", having quietly
 * dropped the digits it cannot hold, and a value the column would read back
 * differently has been changed rather than converted. Comparing the canonical
 * form of the text with the canonical form of the parsed number catches every
 * such case, overflow to Infinity and underflow to zero included.
 */
function roundTripsAsNumber(text: string): boolean {
  const parsed = Number(text);
  if (!Number.isFinite(parsed)) {
    return false;
  }
  const written = canonicalDecimal(text);
  return written !== null && written === canonicalDecimal(String(parsed));
}

function voteForText(value: string): ValueVote {
  const trimmed = value.trim();
  if (BOOLEAN_TEXT.has(trimmed.toLowerCase())) {
    return { ...NOTHING, boolean: true };
  }
  if (isIsoDateText(trimmed)) {
    return { ...NOTHING, date: true };
  }
  if (INTEGER_TEXT.test(trimmed)) {
    // A whole number has to round-trip and to land inside the range an integer
    // column stores exactly, which is the narrower of the two: 2^53 survives
    // the conversion but is past the range `sqlValueFromCell` accepts, so
    // inference must not offer it as an integer either.
    return roundTripsAsNumber(trimmed) && Number.isSafeInteger(Number(trimmed))
      ? { ...NOTHING, integer: true, real: true }
      : NOTHING;
  }
  if (DECIMAL_TEXT.test(trimmed)) {
    return roundTripsAsNumber(trimmed) ? { ...NOTHING, real: true } : NOTHING;
  }
  return NOTHING;
}

function voteFor(value: CellValue): ValueVote {
  if (typeof value === "boolean") {
    return { ...NOTHING, boolean: true };
  }
  if (typeof value === "number") {
    // A value that arrives as a number has already been through the conversion
    // `roundTripsAsNumber` guards, so there is nothing left for it to lose. The
    // range check an integer column needs still applies.
    if (!Number.isFinite(value)) {
      return NOTHING;
    }
    return {
      ...NOTHING,
      integer: Number.isSafeInteger(value),
      real: true,
    };
  }
  return voteForText(value ?? "");
}

function isBlank(value: CellValue): boolean {
  return (
    value === null ||
    value === undefined ||
    (typeof value === "string" && value.trim() === "")
  );
}

/**
 * Infer the column types of a table from its values, conservatively.
 *
 * A column takes a type only when every value that is not blank fits it, and
 * the candidates never overlap, so the answer does not depend on the order the
 * checks run in:
 *
 * - `boolean` when every value is true, false, yes, or no, in any case, or an
 *   actual boolean
 * - `date` when every value is an ISO 8601 date (`2026-01-31`) or date and time
 *   (`2026-01-31T09:30:00Z`), with the calendar checked
 * - `integer` when every value is a whole number with no leading zeros, no
 *   exponent, and no digits past the exactly representable range
 * - `real` when every value is a number under the same spelling rules, with a
 *   fractional part allowed
 * - `text` in every other case, including a column that is entirely blank and
 *   any column whose values disagree
 *
 * Nothing is guessed: a column of "1" and "yes" is text, and so is a column of
 * "2026-01-31" and "31/01/2026", because either could be read two ways and text
 * is the only reading that changes nothing.
 */
export function inferColumnTypes(table: Table): InferredColumn[] {
  return table.columns.map((column) => {
    let valueCount = 0;
    let couldBeInteger = true;
    let couldBeReal = true;
    let couldBeBoolean = true;
    let couldBeDate = true;

    for (const row of table.rows) {
      // An own-property read, so a column named like an Object.prototype member
      // is judged on the row's own value rather than an inherited one.
      const value = Object.prototype.hasOwnProperty.call(row, column)
        ? (row[column] ?? null)
        : null;
      if (isBlank(value)) {
        continue;
      }
      valueCount += 1;
      const vote = voteFor(value);
      couldBeInteger &&= vote.integer;
      couldBeReal &&= vote.real;
      couldBeBoolean &&= vote.boolean;
      couldBeDate &&= vote.date;
    }

    const type: ColumnType =
      valueCount === 0
        ? "text"
        : couldBeBoolean
          ? "boolean"
          : couldBeDate
            ? "date"
            : couldBeInteger
              ? "integer"
              : couldBeReal
                ? "real"
                : "text";
    return { name: column, type, valueCount };
  });
}

/**
 * A safe table name derived from a worksheet or file name, as a starting point
 * a person can edit. Letters and digits in any script are kept, every other run
 * of characters becomes one underscore, and a name the schema would refuse is
 * prefixed rather than silently changed into something unrecognisable.
 */
export function suggestTableName(source: string): string {
  // The first replace collapses every run of other characters to one "_", so
  // the edge trims below never face repeated underscores and stay linear on any
  // input. A trailing "_+$" would not: it can start matching anywhere inside a
  // run, which is quadratic on a name that ends in something else. This is the
  // same shape, and the same fix, as `normalizedColumnKey` in
  // `@consultchimps/tabular`.
  const cleaned = source
    .replace(/[^\p{L}\p{N}]+/gu, "_")
    .replace(/^_/u, "")
    .replace(/_$/u, "")
    .slice(0, 180);
  const candidate = cleaned === "" ? "table" : cleaned;
  try {
    assertSafeIdentifier(candidate, "table");
    return candidate;
  } catch {
    // A reserved prefix ("sqlite_", the package's own) is the only way a
    // cleaned name still fails, and prefixing keeps the original readable.
    return `table_${candidate}`;
  }
}

/**
 * A Record ID prefix derived from a table name, as a starting point a person
 * can edit: the first four characters of a single word ("Customers" gives
 * "CUST"), or the initials of a name with several words ("Sales Orders" gives
 * "SO").
 */
export function suggestRecordIdPrefix(tableName: string): string {
  const words = tableName
    .split(/[^\p{L}\p{N}]+/u)
    .filter((word) => word !== "");
  if (words.length === 0) {
    return "REC";
  }
  const letters =
    words.length === 1
      ? [...(words[0] as string)].slice(0, 4).join("")
      : words
          .slice(0, 4)
          .map((word) => [...word][0] as string)
          .join("");
  return letters.toUpperCase();
}

/**
 * Build the schema an import would create, without creating it. The Record ID
 * column is never declared: it is generated, so an input that already carries
 * one contributes no column here.
 */
export function importedTableSchema(
  table: Table,
  options: ImportTableOptions,
): TableSchema {
  return schemaFromColumns(inferColumnTypes(table), options);
}

// The schema half of an import, taking the inference as given so a table is
// scanned once even though the schema and the report both describe its columns.
function schemaFromColumns(
  inferred: readonly InferredColumn[],
  options: ImportTableOptions,
): TableSchema {
  return {
    name: options.name.trim(),
    columns: inferred
      .filter((column) => identifierKey(column.name) !== RECORD_ID_KEY)
      .map((column) => ({ name: column.name, type: column.type })),
    foreignKeys: [],
    recordId: { ...options.recordId, prefix: options.recordId.prefix.trim() },
  };
}

/** Validate one request's column list before anything is written. */
function assertImportColumns(
  tableName: string,
  columns: readonly ColumnDefinition[],
): void {
  if (columns.length === 0) {
    throw new ConsultChimpsError(
      "DB_IMPORT_NO_COLUMNS",
      `There are no columns to import into "${tableName}", so there is nothing to create. Check that the file has a header row.`,
      { details: { table: tableName } },
    );
  }
  const seen = new Set<string>();
  for (const column of columns) {
    // The same identifier rules table creation applies, run early so a bad
    // header is reported before any table is created.
    assertSafeIdentifier(column.name, "column");
    const key = identifierKey(column.name);
    if (seen.has(key)) {
      throw new ConsultChimpsError(
        "DB_DUPLICATE_COLUMN",
        `The import for "${tableName}" would create the column "${column.name}" more than once.`,
        { details: { table: tableName, column: column.name } },
      );
    }
    seen.add(key);
  }
}

/**
 * Import several tables into a workspace as one unit.
 *
 * Every name, Record ID configuration, and column list is checked first, so the
 * usual failures (a name that is already taken, an empty prefix) are reported
 * before anything is written. What follows runs inside a single transaction:
 * either every table in the call is created and filled, or the workspace is
 * left exactly as it was. A half-finished import, where a first sheet landed
 * and a second failed, is the one outcome a person cannot easily undo, so it is
 * the one outcome this rules out.
 */
export function importTables(
  database: Database,
  requests: readonly ImportTableRequest[],
): ImportedTable[] {
  if (requests.length === 0) {
    throw new ConsultChimpsError(
      "DB_IMPORT_NOTHING_SELECTED",
      "No tables were chosen, so there was nothing to import. Choose at least one and try again.",
      { details: {} },
    );
  }

  const taken = new Map<string, string>();
  for (const schema of database.getSchema()) {
    taken.set(identifierKey(schema.name), schema.name);
  }

  const planned: Array<{
    inferred: InferredColumn[];
    request: ImportTableRequest;
    schema: TableSchema;
  }> = [];
  for (const request of requests) {
    const inferred = inferColumnTypes(request.table);
    const schema = schemaFromColumns(inferred, request);
    // Table and column names are matched the way SQLite matches them, never by
    // a fresh case-insensitive comparison, so "Customer" and "customer" are one
    // name here exactly as they are in the engine.
    assertSafeIdentifier(schema.name, "table");
    assertRecordIdConfig(schema.name, schema.recordId);
    const key = identifierKey(schema.name);
    const existing = taken.get(key);
    if (existing !== undefined) {
      throw new ConsultChimpsError(
        "DB_TABLE_EXISTS",
        `The workspace already holds a table called "${existing}", so "${schema.name}" cannot be created. Choose a different name.`,
        { details: { table: schema.name, existingTable: existing } },
      );
    }
    taken.set(key, schema.name);
    assertImportColumns(schema.name, schema.columns);
    planned.push({ inferred, request, schema });
  }

  return database.sql.transaction(() =>
    planned.map(({ inferred, request, schema }) => {
      database.createTable(schema);
      const inserted = insertRecordsFromTable(
        database,
        schema.name,
        request.table,
      );
      return {
        name: schema.name,
        columns: inferred.filter(
          (column) => identifierKey(column.name) !== RECORD_ID_KEY,
        ),
        recordId: schema.recordId,
        rowCount: inserted.length,
        firstRecordId: inserted[0]?.recordId ?? null,
        lastRecordId: inserted[inserted.length - 1]?.recordId ?? null,
        ignoredColumns: request.table.columns.filter(
          (column) => identifierKey(column) === RECORD_ID_KEY,
        ),
      };
    }),
  );
}

/** Import one table into a workspace. The single-table form of `importTables`. */
export function importTable(
  database: Database,
  table: Table,
  options: ImportTableOptions,
): ImportedTable {
  return importTables(database, [{ ...options, table }])[0] as ImportedTable;
}
