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
 */
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/u;
const ISO_DATE_TIME =
  /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2})(?:\.\d{1,9})?)?(?:Z|[+-]\d{2}:\d{2})?$/u;

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

function isIsoDateText(value: string): boolean {
  const date = ISO_DATE.exec(value);
  if (date !== null) {
    return isCalendarDate(Number(date[1]), Number(date[2]), Number(date[3]));
  }
  const stamp = ISO_DATE_TIME.exec(value);
  if (stamp === null) {
    return false;
  }
  if (!isCalendarDate(Number(stamp[1]), Number(stamp[2]), Number(stamp[3]))) {
    return false;
  }
  const hour = Number(stamp[4]);
  const minute = Number(stamp[5]);
  const second = stamp[6] === undefined ? 0 : Number(stamp[6]);
  // 24:00 and a leap second are legal ISO spellings, so the bounds are the
  // written ones rather than a clock's.
  return hour <= 24 && minute <= 59 && second <= 60;
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

function voteForText(value: string): ValueVote {
  const trimmed = value.trim();
  if (BOOLEAN_TEXT.has(trimmed.toLowerCase())) {
    return { ...NOTHING, boolean: true };
  }
  if (isIsoDateText(trimmed)) {
    return { ...NOTHING, date: true };
  }
  if (INTEGER_TEXT.test(trimmed)) {
    // A whole number past the exactly representable range would be rounded on
    // the way in, so it stays text and keeps its digits.
    const exact = BigInt(trimmed);
    const representable =
      exact >= BigInt(Number.MIN_SAFE_INTEGER) &&
      exact <= BigInt(Number.MAX_SAFE_INTEGER);
    return representable ? { ...NOTHING, integer: true, real: true } : NOTHING;
  }
  if (DECIMAL_TEXT.test(trimmed) && Number.isFinite(Number(trimmed))) {
    return { ...NOTHING, real: true };
  }
  return NOTHING;
}

function voteFor(value: CellValue): ValueVote {
  if (typeof value === "boolean") {
    return { ...NOTHING, boolean: true };
  }
  if (typeof value === "number") {
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
