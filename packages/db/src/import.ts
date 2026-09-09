import { ConsultChimpsError } from "@consultchimps/core";
import {
  uniqueHeaders,
  type CellValue,
  type Table,
  type TableRow,
} from "@consultchimps/tabular";

import { insertRecordsFromTable } from "./bridge.js";
import type { Database } from "./database.js";
import {
  assertRecordIdConfig,
  assertSafeIdentifier,
  identifierKey,
  isIsoDateText,
  truncateIdentifier,
  MAX_IDENTIFIER_LENGTH,
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
  /**
   * Columns stored under a different name from the one the file wrote, because
   * the header was longer than a name may be or collided with another once it
   * had been shortened. Every value is still there; only the name changed.
   */
  readonly renamedColumns: RenamedColumn[];
}

/** A column of the source, and the name the table carries it under. */
export interface RenamedColumn {
  /** The header as the file wrote it. */
  readonly from: string;
  /** The name it is stored as, once it had to fit. */
  readonly to: string;
}

/** The source table as both inference and storage will see it. */
interface PreparedTable {
  readonly table: Table;
  readonly renamedColumns: RenamedColumn[];
}

/**
 * How much room a name has to leave for the number that may be added to it.
 *
 * `uniqueHeaders` numbers a repeated name `_2`, `_3`, and so on, taking the
 * lowest number free. What blocks a number is a name already taken or one the
 * header row carries further along, and there is at most one of each per column,
 * so the number can never pass twice the column count plus two. Room for that
 * many digits, and the underscore, is what makes the finished name fit by
 * construction rather than by hoping it does.
 */
function suffixHeadroom(columnCount: number): number {
  return 1 + String(2 * columnCount + 2).length;
}

/**
 * Read a table the way both inference and storage will see it: blanks resolved,
 * and every column under the name it will actually be stored as.
 *
 * A cell holding nothing but whitespace holds nothing. Inference already treats
 * one as blank, so without this the column's type is decided as if the cell were
 * empty while the cell itself is stored as spaces: a value that is neither
 * absent nor present, that `IS NULL` does not find.
 *
 * Only blankness is normalized in the values. Surrounding spaces on a value that
 * is not blank are removed by the conversion for a typed column, which is the
 * one place that knows the column's type; a text column keeps its content
 * exactly as the file wrote it.
 *
 * The names are the other half of the same idea. A header may be any length in a
 * `Table`, which is a runtime-neutral model that knows nothing about a database,
 * while a column name here has a limit. Numbering a repeated header and then
 * checking that limit is two steps in two layers, and it fails between them: two
 * copies of a 200-character header become one name of 202, and the whole import
 * is refused over a name the person had no way to change, for a duplicate the
 * import was supposed to number for them. So the limit is applied first, with
 * room left for the number, and the numbering runs over names that already fit.
 * Headers that differed only past the cut collide and are numbered,
 * deterministically, and no value is lost: every column is still stored, under a
 * name the result reports.
 */
function prepareForImport(table: Table): PreparedTable {
  const budget = Math.max(
    1,
    MAX_IDENTIFIER_LENGTH - suffixHeadroom(table.columns.length),
  );
  // A header cut down to nothing, which takes a name wider than the budget,
  // becomes a blank for `uniqueHeaders` to fill the way it fills any other.
  const columns = uniqueHeaders(
    table.columns.map((column) => truncateIdentifier(column, budget) || null),
  );
  const renamedColumns: RenamedColumn[] = [];
  table.columns.forEach((from, index) => {
    const to = columns[index] as string;
    if (to !== from) {
      renamedColumns.push({ from, to });
    }
  });

  return {
    renamedColumns,
    table: {
      ...table,
      columns,
      rows: table.rows.map((row) => {
        // A prototype-free destination, and own-property reads, for the reason
        // `addRecordsFromTable` gives.
        const cleaned: TableRow = Object.create(null);
        table.columns.forEach((from, index) => {
          const value = Object.prototype.hasOwnProperty.call(row, from)
            ? (row[from] ?? null)
            : null;
          cleaned[columns[index] as string] =
            typeof value === "string" && value.trim() === "" ? null : value;
        });
        return cleaned;
      }),
    },
  };
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
 * The text spellings read as a boolean. "1" and "0" are deliberately absent:
 * they are numbers in every file that also holds counts, and a column of ones
 * and zeros is far more often a quantity than a flag. A column that really is a
 * flag written that way imports as a number and can be changed afterwards,
 * which is recoverable; guessing the other way is not.
 */
const BOOLEAN_TEXT = new Set(["true", "false", "yes", "no"]);

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
 *
 * This judges the table it is handed, exactly as handed: it is the inference
 * primitive, not the import. It does not resolve blank cells or shorten a header
 * to a name a database will take, because it has no column names to give and no
 * limit to give them against. An import prepares a table first and infers from
 * that, which is why `importedTableSchema` and `importTables` agree; a caller
 * running this on raw headers is asking a narrower question and gets an answer
 * to that question.
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

/** Put in front of a cleaned name the schema would otherwise refuse. */
const FALLBACK_TABLE_PREFIX = "table_";

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
  //
  // The budget leaves room for the fallback prefix below, so a name that needs
  // both shortening and prefixing still fits, and it is spent in the unit the
  // limit is written in rather than a number picked to be safely under it.
  const cleaned = truncateIdentifier(
    source
      .replace(/[^\p{L}\p{N}]+/gu, "_")
      .replace(/^_/u, "")
      .replace(/_$/u, ""),
    MAX_IDENTIFIER_LENGTH - FALLBACK_TABLE_PREFIX.length,
  );
  const candidate = cleaned === "" ? "table" : cleaned;
  try {
    assertSafeIdentifier(candidate, "table");
    return candidate;
  } catch {
    // A reserved prefix ("sqlite_", the package's own) is the only way a
    // cleaned name still fails, and prefixing keeps the original readable.
    return `${FALLBACK_TABLE_PREFIX}${candidate}`;
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
  return planImport(table, options).schema;
}

// The schema half of an import, taking the inference as given so a table is
// scanned once even though the schema and the report both describe its columns.
/** What one table of an import will be, worked out once. */
interface PlannedTable {
  /** The source, prepared: blanks resolved and columns under storable names. */
  readonly table: Table;
  /** The types inferred from that prepared table. */
  readonly inferred: InferredColumn[];
  /** Headers stored under a different name from the one the file wrote. */
  readonly renamedColumns: RenamedColumn[];
  /** The schema the import will create. */
  readonly schema: TableSchema;
}

/**
 * Work out what importing this table would create.
 *
 * The one derivation. A preview that repeated any part of it would eventually
 * describe a table the import does not build: `importedTableSchema` used to
 * infer straight from the raw headers, so it advertised a column name of 201
 * characters that the import itself would never have created, and a duplicate
 * under a different name from the one that would be stored.
 */
function planImport(table: Table, options: ImportTableOptions): PlannedTable {
  const { renamedColumns, table: prepared } = prepareForImport(table);
  const inferred = inferColumnTypes(prepared);
  return {
    inferred,
    renamedColumns,
    schema: schemaFromColumns(inferred, options),
    table: prepared,
  };
}

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
    renamedColumns: RenamedColumn[];
    request: ImportTableRequest;
    schema: TableSchema;
  }> = [];
  for (const request of requests) {
    // The same derivation `importedTableSchema` answers with, so what a caller
    // was shown is what gets built.
    const { inferred, renamedColumns, schema, table } = planImport(
      request.table,
      request,
    );
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
    planned.push({
      inferred,
      renamedColumns,
      request: { ...request, table },
      schema,
    });
  }

  return database.sql.transaction(() =>
    planned.map(({ inferred, renamedColumns, request, schema }) => {
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
        renamedColumns,
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
