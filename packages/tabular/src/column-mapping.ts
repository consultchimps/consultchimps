import { ConsultChimpsError } from "@consultchimps/core";

import {
  normalizedColumnKey,
  type CellValue,
  type Table,
  type TableRow,
} from "./table.js";

/**
 * The only column mapping document version this engine understands. The
 * mapping document is a public file format: a new shape gets a new version
 * rather than a silent reinterpretation of this one.
 */
export const COLUMN_MAPPING_VERSION = 1;

/**
 * Parse a source column's text as a date written in `format` and write it out
 * as an ISO 8601 date string ("2024-03-09"). Supported format tokens are
 * `YYYY` (four digits), `MM`/`DD` (exactly two digits), and `M`/`D` (one or
 * two digits); every other character in the format is a literal separator.
 */
export interface DateColumnCoercion {
  type: "date";
  format: string;
}

/**
 * Parse a source column's text as a number written with the declared
 * separators. Both default to the English convention: "." for decimals and
 * "," for thousands. An empty `thousandsSeparator` means the source writes no
 * thousands separator at all.
 */
export interface NumberColumnCoercion {
  type: "number";
  decimalSeparator?: string;
  thousandsSeparator?: string;
}

export type ColumnCoercion = DateColumnCoercion | NumberColumnCoercion;

/**
 * One canonical column of a mapping: the output name, the source header
 * spellings that fold into it, and an optional coercion applied to its values.
 *
 * Aliases match by normalized column key, so one entry catches every spacing,
 * punctuation, and case variant of that spelling. A canonical column always
 * matches its own name too - it never has to repeat itself in `aliases`.
 */
export interface CanonicalColumn {
  name: string;
  aliases: string[];
  coercion?: ColumnCoercion;
}

/**
 * The versioned column mapping document, applied during consolidation.
 * Global scope: one set of canonical columns for every input.
 */
export interface ColumnMapping {
  version: typeof COLUMN_MAPPING_VERSION;
  columns: CanonicalColumn[];
  /** Constant columns appended to every row, in declaration order. */
  constants?: Record<string, CellValue>;
}

/** A mapped table plus the unmapped columns that passed through unchanged. */
export interface ColumnMappingResult {
  table: Table;
  unmappedColumns: string[];
}

/**
 * The many-table form of {@link ColumnMappingResult}. `unmappedColumns` lists
 * every distinct unmapped spelling across the inputs, in first-seen order, so
 * one warning can name them all.
 */
export interface ColumnMappingTablesResult {
  tables: Table[];
  unmappedColumns: string[];
}

/** The header list of one input, with whatever provenance the caller knows. */
export interface ColumnHeaderSource {
  columns: string[];
  file?: string | undefined;
  sheet?: string | undefined;
}

/** Where one spelling of a suggested group was seen. */
export interface ColumnSpellingOccurrence {
  column: string;
  file?: string;
  sheet?: string;
}

/**
 * Columns whose normalized keys already match, differing only in spelling.
 * `canonical` is the first spelling seen, which is what the drafted mapping
 * proposes as the output name.
 */
export interface ColumnEquivalenceGroup {
  canonical: string;
  key: string;
  spellings: string[];
  occurrences: ColumnSpellingOccurrence[];
}

/** A drafted mapping plus the evidence each of its entries rests on. */
export interface ColumnMappingSuggestion {
  mapping: ColumnMapping;
  groups: ColumnEquivalenceGroup[];
}

const MAPPING_INVALID = "TABLE_MAPPING_INVALID";
const MAPPING_COLUMN_COLLISION = "TABLE_MAPPING_COLUMN_COLLISION";
const MAPPING_CONSTANT_COLLISION = "TABLE_MAPPING_CONSTANT_COLLISION";
const MAPPING_COERCION_FAILED = "TABLE_MAPPING_COERCION_FAILED";

const DEFAULT_DECIMAL_SEPARATOR = ".";
const DEFAULT_THOUSANDS_SEPARATOR = ",";

const DAYS_IN_MONTH = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31] as const;

type DateFieldName = "year" | "month" | "day";

interface DateFormatPlan {
  fields: DateFieldName[];
  pattern: RegExp;
}

interface CompiledCoercion {
  coercion: ColumnCoercion;
  date?: DateFormatPlan;
  number?: RegExp;
}

interface CompiledColumn {
  name: string;
  coercion?: CompiledCoercion;
}

interface CompiledMapping {
  columnByKey: Map<string, CompiledColumn>;
  constants: Array<{ name: string; value: CellValue }>;
  mapping: ColumnMapping;
}

interface CoercionOutcome {
  ok: boolean;
  value: CellValue;
}

function invalidMapping(
  problem: string,
  message: string,
  details: Record<string, unknown> = {},
): ConsultChimpsError {
  return new ConsultChimpsError(
    MAPPING_INVALID,
    `The column mapping is not usable. ${message}`,
    { details: { problem, ...details } },
  );
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isCellValue(value: unknown): value is CellValue {
  return (
    value === null ||
    typeof value === "string" ||
    typeof value === "boolean" ||
    (typeof value === "number" && Number.isFinite(value))
  );
}

function describeIndexedColumn(index: number, name?: string): string {
  return name === undefined
    ? `Column entry ${index + 1}`
    : `Canonical column "${name}"`;
}

/**
 * Describe the input a table came from, for error messages: " in sheet
 * \"South\" of \"records.xlsx\"", or "" when the table carries no provenance.
 */
function describeTableSource(table: Table): string {
  const sheet = table.source?.sheet;
  const file = table.source?.file;
  if (sheet && file) {
    return ` in sheet "${sheet}" of "${file}"`;
  }
  if (sheet) {
    return ` in sheet "${sheet}"`;
  }
  if (file) {
    return ` in "${file}"`;
  }
  return "";
}

function sourceDetails(table: Table): Record<string, unknown> {
  const details: Record<string, unknown> = {};
  if (table.source?.file !== undefined) {
    details["file"] = table.source.file;
  }
  if (table.source?.sheet !== undefined) {
    details["sheet"] = table.source.sheet;
  }
  return details;
}

function sourceRowNumber(table: Table, index: number): number {
  return table.sourceRows?.[index] ?? (table.source?.firstDataRow ?? 2) + index;
}

function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

function daysInMonth(year: number, month: number): number {
  if (month === 2 && isLeapYear(year)) {
    return 29;
  }
  return DAYS_IN_MONTH[month - 1] ?? 0;
}

function escapeRegExp(literal: string): string {
  return literal.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
}

const DATE_FORMAT_TOKENS: ReadonlyArray<
  readonly [token: string, pattern: string, field: DateFieldName]
> = [
  ["YYYY", "(\\d{4})", "year"],
  ["MM", "(\\d{2})", "month"],
  ["DD", "(\\d{2})", "day"],
  ["M", "(\\d{1,2})", "month"],
  ["D", "(\\d{1,2})", "day"],
];

/**
 * Turn a declared date format into an anchored pattern plus the order its
 * capture groups appear in. Returns undefined when the format does not name
 * each of year, month, and day exactly once, which would make parsing
 * ambiguous or impossible.
 */
function planDateFormat(format: string): DateFormatPlan | undefined {
  const fields: DateFieldName[] = [];
  let pattern = "";
  let literal = "";
  let position = 0;

  while (position < format.length) {
    const match = DATE_FORMAT_TOKENS.find(([token]) =>
      format.startsWith(token, position),
    );
    if (!match) {
      literal += format[position];
      position += 1;
      continue;
    }
    pattern += escapeRegExp(literal);
    literal = "";
    pattern += match[1];
    fields.push(match[2]);
    position += match[0].length;
  }
  pattern += escapeRegExp(literal);

  for (const field of ["year", "month", "day"] as const) {
    if (fields.filter((candidate) => candidate === field).length !== 1) {
      return undefined;
    }
  }

  return { fields, pattern: new RegExp(`^${pattern}$`, "u") };
}

/**
 * The shape a number written with these separators is allowed to take, used
 * to check a value before any separator is removed. Thousands separators may
 * only group the integer part into runs of three, so a malformed grouping
 * such as "1,23.4" is refused instead of quietly reading as 1.234 once the
 * separators are stripped. The fractional part never carries a grouping
 * separator.
 */
function planNumberFormat(thousands: string, decimal: string): RegExp {
  const groupSeparator = escapeRegExp(thousands);
  const decimalSeparator = escapeRegExp(decimal);
  const integer =
    thousands === "" ? "\\d+" : `\\d{1,3}(?:${groupSeparator}\\d{3})+|\\d+`;
  const fraction = `${decimalSeparator}\\d+`;
  return new RegExp(
    `^[+-]?(?:(?:${integer})(?:${fraction})?|${fraction})$`,
    "u",
  );
}

function compileCoercion(
  value: unknown,
  describe: string,
): CompiledCoercion | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (!isPlainObject(value)) {
    throw invalidMapping(
      "invalid_coercion",
      `${describe} declares a coercion that is not an object.`,
    );
  }

  const type = value["type"];
  if (type === "date") {
    const format = value["format"];
    if (typeof format !== "string" || format.trim() === "") {
      throw invalidMapping(
        "invalid_date_format",
        `${describe} declares a date coercion without a source format.`,
      );
    }
    const date = planDateFormat(format);
    if (!date) {
      throw invalidMapping(
        "invalid_date_format",
        `${describe} declares the date format "${format}", which does not name a year, a month, and a day exactly once. Use tokens YYYY, MM or M, and DD or D.`,
        { format },
      );
    }
    return { coercion: { type: "date", format }, date };
  }

  if (type === "number") {
    const decimalSeparator =
      value["decimalSeparator"] ?? DEFAULT_DECIMAL_SEPARATOR;
    const thousandsSeparator =
      value["thousandsSeparator"] ?? DEFAULT_THOUSANDS_SEPARATOR;
    if (
      typeof decimalSeparator !== "string" ||
      typeof thousandsSeparator !== "string" ||
      decimalSeparator.length !== 1 ||
      thousandsSeparator.length > 1 ||
      /[\d+-]/u.test(decimalSeparator) ||
      /[\d+-]/u.test(thousandsSeparator) ||
      decimalSeparator === thousandsSeparator
    ) {
      throw invalidMapping(
        "invalid_number_separators",
        `${describe} declares number separators that cannot be told apart. The decimal separator must be one non-digit character, and the thousands separator must be empty or one different non-digit character.`,
        { decimalSeparator, thousandsSeparator },
      );
    }
    return {
      coercion: { type: "number", decimalSeparator, thousandsSeparator },
      number: planNumberFormat(thousandsSeparator, decimalSeparator),
    };
  }

  throw invalidMapping(
    "unknown_coercion_type",
    `${describe} declares the unknown coercion type ${JSON.stringify(type)}. Supported types are "date" and "number".`,
    { type },
  );
}

/**
 * Validate a parsed mapping document and build its lookup tables.
 *
 * Because matching is normalized, two entries whose normalized keys are equal
 * would leave the document unable to say where a source header belongs. So
 * canonical names must be distinct after normalization, no two aliases may
 * normalize alike - within one canonical column or across two - and a
 * constant column may not take a mapped column's key. An alias that merely
 * restates its own canonical column's name is allowed: it resolves to the
 * same column either way.
 */
function compileColumnMapping(value: unknown): CompiledMapping {
  if (!isPlainObject(value)) {
    throw invalidMapping(
      "not_an_object",
      "A mapping must be a JSON object with a version and a list of canonical columns.",
    );
  }

  const version = value["version"];
  if (version !== COLUMN_MAPPING_VERSION) {
    throw invalidMapping(
      "unsupported_version",
      `Version ${JSON.stringify(version)} is not supported. This toolkit reads version ${COLUMN_MAPPING_VERSION} mappings.`,
      { supportedVersion: COLUMN_MAPPING_VERSION, version },
    );
  }

  const rawColumns = value["columns"];
  if (!Array.isArray(rawColumns)) {
    throw invalidMapping(
      "columns_not_a_list",
      'The mapping\'s "columns" property must be a list of canonical columns.',
    );
  }

  const columnByKey = new Map<string, CompiledColumn>();
  const canonicalKeys = new Map<string, string>();
  const aliasKeys = new Map<string, string>();
  const columns: CanonicalColumn[] = [];
  const declaredAliases: Array<{
    aliases: unknown[];
    compiled: CompiledColumn;
    declared: CanonicalColumn;
    index: number;
  }> = [];

  // Pass one claims every canonical name, so an alias colliding with a
  // canonical declared later in the document is still reported as an alias
  // problem rather than depending on declaration order.
  rawColumns.forEach((entry, index) => {
    if (!isPlainObject(entry)) {
      throw invalidMapping(
        "invalid_column_entry",
        `${describeIndexedColumn(index)} is not an object.`,
        { index },
      );
    }
    const name = entry["name"];
    if (typeof name !== "string" || name.trim() === "") {
      throw invalidMapping(
        "missing_canonical_name",
        `${describeIndexedColumn(index)} has no canonical column name.`,
        { index },
      );
    }
    const aliases = entry["aliases"];
    if (!Array.isArray(aliases)) {
      throw invalidMapping(
        "aliases_not_a_list",
        `${describeIndexedColumn(index, name)} must list its aliases, even if the list is empty.`,
        { canonicalColumn: name },
      );
    }

    const key = normalizedColumnKey(name);
    const owner = canonicalKeys.get(key);
    if (owner !== undefined) {
      throw invalidMapping(
        "duplicate_canonical_column",
        `Canonical columns "${owner}" and "${name}" match the same source headers. Keep one of them.`,
        { canonicalColumn: name, conflictsWith: owner, key },
      );
    }
    canonicalKeys.set(key, name);

    const compiled: CompiledColumn = { name };
    const coercion = compileCoercion(
      entry["coercion"],
      describeIndexedColumn(index, name),
    );
    if (coercion) {
      compiled.coercion = coercion;
    }
    columnByKey.set(key, compiled);

    const declared: CanonicalColumn = { name, aliases: [] };
    if (coercion) {
      declared.coercion = coercion.coercion;
    }
    columns.push(declared);
    declaredAliases.push({ aliases, compiled, declared, index });
  });

  declaredAliases.forEach(({ aliases, compiled, declared, index }) => {
    const name = compiled.name;

    for (const alias of aliases) {
      if (typeof alias !== "string" || alias.trim() === "") {
        throw invalidMapping(
          "invalid_alias",
          `${describeIndexedColumn(index, name)} lists an alias that is not a non-empty header spelling.`,
          { canonicalColumn: name },
        );
      }
      const key = normalizedColumnKey(alias);
      // An alias may restate its own canonical column's name - that resolves
      // to the same column - but two aliases that normalize alike, or an
      // alias that reaches into another canonical column, leave the document
      // unable to say where a source header belongs.
      const owner = aliasKeys.get(key) ?? canonicalKeys.get(key);
      if (owner !== undefined && !(owner === name && !aliasKeys.has(key))) {
        throw invalidMapping(
          "duplicate_alias",
          owner === name
            ? `Canonical column "${name}" lists the alias "${alias}" more than once, counting spelling variants.`
            : `The alias "${alias}" is claimed by both "${owner}" and "${name}". An alias can belong to only one canonical column.`,
          { alias, canonicalColumn: name, conflictsWith: owner, key },
        );
      }
      aliasKeys.set(key, name);
      columnByKey.set(key, compiled);
      declared.aliases.push(alias);
    }
  });

  const rawConstants = value["constants"];
  const constants: Array<{ name: string; value: CellValue }> = [];
  if (rawConstants !== undefined) {
    if (!isPlainObject(rawConstants)) {
      throw invalidMapping(
        "constants_not_an_object",
        'The mapping\'s "constants" property must be an object of column names to values.',
      );
    }
    for (const [name, constantValue] of Object.entries(rawConstants)) {
      if (name.trim() === "") {
        throw invalidMapping(
          "invalid_constant_column",
          "A constant column needs a name.",
        );
      }
      if (!isCellValue(constantValue)) {
        throw invalidMapping(
          "invalid_constant_value",
          `Constant column "${name}" must hold a text, number, true/false, or null value.`,
          { constantColumn: name },
        );
      }
      const key = normalizedColumnKey(name);
      const owner = canonicalKeys.get(key) ?? aliasKeys.get(key);
      if (owner !== undefined) {
        throw invalidMapping(
          "constant_column_collision",
          `Constant column "${name}" collides with the mapped column "${owner}". Rename one of them.`,
          { conflictsWith: owner, constantColumn: name, key },
        );
      }
      canonicalKeys.set(key, name);
      constants.push({ name, value: constantValue });
    }
  }

  const mapping: ColumnMapping = {
    version: COLUMN_MAPPING_VERSION,
    columns,
  };
  if (rawConstants !== undefined) {
    mapping.constants = Object.fromEntries(
      constants.map(({ name, value: constantValue }) => [name, constantValue]),
    );
  }

  return { columnByKey, constants, mapping };
}

/**
 * Check that a parsed value is a well-formed version 1 mapping and return it
 * typed. Reading and parsing the mapping file is the caller's job; this
 * package never touches the filesystem.
 *
 * Throws `TABLE_MAPPING_INVALID` naming the first problem found, with a
 * machine-readable `problem` in the error details.
 */
export function validateColumnMapping(value: unknown): ColumnMapping {
  return compileColumnMapping(value).mapping;
}

function coerceDate(value: CellValue, plan: DateFormatPlan): CoercionOutcome {
  if (typeof value !== "string") {
    return { ok: false, value };
  }
  const text = value.trim();
  if (text === "") {
    return { ok: true, value: null };
  }
  const match = plan.pattern.exec(text);
  if (!match) {
    return { ok: false, value };
  }

  let year = 0;
  let month = 0;
  let day = 0;
  plan.fields.forEach((field, index) => {
    const digits = Number(match[index + 1]);
    if (field === "year") {
      year = digits;
    } else if (field === "month") {
      month = digits;
    } else {
      day = digits;
    }
  });

  if (month < 1 || month > 12 || day < 1 || day > daysInMonth(year, month)) {
    return { ok: false, value };
  }

  const isoYear = String(year).padStart(4, "0");
  const isoMonth = String(month).padStart(2, "0");
  const isoDay = String(day).padStart(2, "0");
  return { ok: true, value: `${isoYear}-${isoMonth}-${isoDay}` };
}

function coerceNumber(
  value: CellValue,
  coercion: NumberColumnCoercion,
  pattern: RegExp,
): CoercionOutcome {
  if (typeof value === "number") {
    return Number.isFinite(value)
      ? { ok: true, value }
      : { ok: false, value: null };
  }
  if (typeof value !== "string") {
    return { ok: false, value };
  }

  const text = value.trim();
  if (text === "") {
    return { ok: true, value: null };
  }

  // The written form is checked first, so a separator can only be removed
  // once it has been shown to sit where the declared format allows.
  if (!pattern.test(text)) {
    return { ok: false, value };
  }

  const decimalSeparator =
    coercion.decimalSeparator ?? DEFAULT_DECIMAL_SEPARATOR;
  const thousandsSeparator =
    coercion.thousandsSeparator ?? DEFAULT_THOUSANDS_SEPARATOR;

  let candidate = text;
  if (thousandsSeparator !== "") {
    candidate = candidate.split(thousandsSeparator).join("");
  }
  if (decimalSeparator !== ".") {
    candidate = candidate.split(decimalSeparator).join(".");
  }

  const parsed = Number(candidate);
  return Number.isFinite(parsed)
    ? { ok: true, value: parsed }
    : { ok: false, value };
}

function applyCoercion(
  value: CellValue,
  compiled: CompiledCoercion,
): CoercionOutcome {
  if (value === null) {
    return { ok: true, value: null };
  }
  // The plans are built during validation, so they always exist here.
  if (compiled.coercion.type === "date") {
    return compiled.date
      ? coerceDate(value, compiled.date)
      : { ok: false, value };
  }
  return compiled.number
    ? coerceNumber(value, compiled.coercion, compiled.number)
    : { ok: false, value };
}

function describeCoercion(coercion: ColumnCoercion): string {
  if (coercion.type === "date") {
    return `a date in the declared format "${coercion.format}"`;
  }
  const thousands = coercion.thousandsSeparator ?? DEFAULT_THOUSANDS_SEPARATOR;
  const decimal = coercion.decimalSeparator ?? DEFAULT_DECIMAL_SEPARATOR;
  return thousands === ""
    ? `a number with "${decimal}" as the decimal separator`
    : `a number with "${thousands}" for thousands and "${decimal}" for decimals`;
}

/**
 * A row that treats every column name as ordinary data. Assigning to a plain
 * object's "__proto__" would run the inherited setter instead of storing the
 * value, so a column with that name would be listed in `columns` and missing
 * from every row - a silent hole in the promised lossless passthrough.
 */
function createRow(): TableRow {
  return Object.create(null) as TableRow;
}

/**
 * Read one cell, ignoring anything a row inherits. Without the own-property
 * check, reading the column "__proto__" from a plain object row would hand
 * back `Object.prototype` instead of a missing value.
 */
function readCell(row: TableRow, column: string): CellValue {
  return Object.hasOwn(row, column) ? (row[column] ?? null) : null;
}

interface MappedColumn {
  coercion?: CompiledCoercion;
  input: string;
  output: string;
}

function applyCompiledMapping(
  table: Table,
  compiled: CompiledMapping,
): ColumnMappingResult {
  const mapped: MappedColumn[] = [];
  const unmappedColumns: string[] = [];
  const canonicalOwner = new Map<string, string>();

  for (const column of table.columns) {
    const canonical = compiled.columnByKey.get(normalizedColumnKey(column));
    if (!canonical) {
      unmappedColumns.push(column);
      mapped.push({ input: column, output: column });
      continue;
    }

    const owner = canonicalOwner.get(canonical.name);
    if (owner !== undefined) {
      throw new ConsultChimpsError(
        MAPPING_COLUMN_COLLISION,
        `Columns "${owner}" and "${column}"${describeTableSource(table)} both map to the canonical column "${canonical.name}". Combining them would silently drop one of the two values, so nothing was written. Map only one of them, or rename one column in the source.`,
        {
          details: {
            canonicalColumn: canonical.name,
            columns: [owner, column],
            ...sourceDetails(table),
          },
        },
      );
    }
    canonicalOwner.set(canonical.name, column);

    const entry: MappedColumn = { input: column, output: canonical.name };
    if (canonical.coercion) {
      entry.coercion = canonical.coercion;
    }
    mapped.push(entry);
  }

  const outputKeys = new Set(
    mapped.map(({ output }) => normalizedColumnKey(output)),
  );
  for (const constant of compiled.constants) {
    const key = normalizedColumnKey(constant.name);
    if (outputKeys.has(key)) {
      const existing = mapped.find(
        ({ output }) => normalizedColumnKey(output) === key,
      );
      throw new ConsultChimpsError(
        MAPPING_CONSTANT_COLLISION,
        `The constant column "${constant.name}" collides with the column "${existing?.output ?? constant.name}"${describeTableSource(table)}. Rename the constant column, or map that source column to a different canonical column.`,
        {
          details: {
            column: existing?.output ?? constant.name,
            constantColumn: constant.name,
            ...sourceDetails(table),
          },
        },
      );
    }
    outputKeys.add(key);
  }

  const rows: TableRow[] = table.rows.map((inputRow, index) => {
    const outputRow = createRow();
    for (const { coercion, input, output } of mapped) {
      const value = readCell(inputRow, input);
      if (!coercion) {
        outputRow[output] = value;
        continue;
      }
      const outcome = applyCoercion(value, coercion);
      if (!outcome.ok) {
        throw new ConsultChimpsError(
          MAPPING_COERCION_FAILED,
          `Row ${sourceRowNumber(table, index)} of column "${input}"${describeTableSource(table)} does not hold ${describeCoercion(coercion.coercion)}, so the canonical column "${output}" could not be filled in and nothing was written. Correct that cell, or remove the coercion from the mapping.`,
          {
            details: {
              canonicalColumn: output,
              coercion: coercion.coercion.type,
              column: input,
              row: sourceRowNumber(table, index),
              ...sourceDetails(table),
            },
          },
        );
      }
      outputRow[output] = outcome.value;
    }
    for (const constant of compiled.constants) {
      outputRow[constant.name] = constant.value;
    }
    return outputRow;
  });

  const result: Table = {
    columns: [
      ...mapped.map(({ output }) => output),
      ...compiled.constants.map(({ name }) => name),
    ],
    rows,
  };
  if (table.source) {
    result.source = { ...table.source };
  }
  if (table.sourceRows) {
    result.sourceRows = [...table.sourceRows];
  }

  return { table: result, unmappedColumns };
}

/**
 * Apply a column mapping to one table: aliases fold into their canonical
 * column, unmapped columns pass through under their own names, declared
 * coercions run, and constant columns are appended last.
 *
 * Matching is by normalized column key, so it is independent of the
 * `normalizeHeaders` union option; canonical names are written verbatim.
 * Apply this to each source table before `unionTables`, so that two columns
 * of the same sheet folding into one canonical column is caught per sheet.
 */
export function applyColumnMapping(
  table: Table,
  mapping: ColumnMapping,
): ColumnMappingResult {
  return applyCompiledMapping(table, compileColumnMapping(mapping));
}

/**
 * Apply one mapping to many tables, validating it once. The reported unmapped
 * columns are the distinct spellings across every input, in first-seen order.
 */
export function applyColumnMappingToTables(
  tables: Table[],
  mapping: ColumnMapping,
): ColumnMappingTablesResult {
  const compiled = compileColumnMapping(mapping);
  const mappedTables: Table[] = [];
  const unmappedColumns: string[] = [];
  const seen = new Set<string>();

  for (const table of tables) {
    const result = applyCompiledMapping(table, compiled);
    mappedTables.push(result.table);
    for (const column of result.unmappedColumns) {
      if (!seen.has(column)) {
        seen.add(column);
        unmappedColumns.push(column);
      }
    }
  }

  return { tables: mappedTables, unmappedColumns };
}

/**
 * Draft a mapping from the inputs' header lists by grouping columns whose
 * normalized keys already match - the spelling variants of one field. The
 * first spelling seen becomes the proposed canonical name.
 *
 * Deliberately no similarity matching and no value-based signals: a group
 * here is an exact normalization equivalence. That still conflates headers
 * that differ only in punctuation ("A+B" and "A-B" share the key a_b), so
 * the draft is evidence for a human review, never a guarantee. Synonyms
 * that share no spelling stay a manual mapping entry; nothing applies the
 * draft automatically.
 */
export function suggestColumnMapping(
  sources: ColumnHeaderSource[],
): ColumnMappingSuggestion {
  const groupsByKey = new Map<string, ColumnEquivalenceGroup>();

  for (const source of sources) {
    for (const column of source.columns) {
      const key = normalizedColumnKey(column);
      let group = groupsByKey.get(key);
      if (!group) {
        group = { canonical: column, key, occurrences: [], spellings: [] };
        groupsByKey.set(key, group);
      }
      if (!group.spellings.includes(column)) {
        group.spellings.push(column);
      }
      const occurrence: ColumnSpellingOccurrence = { column };
      if (source.file !== undefined) {
        occurrence.file = source.file;
      }
      if (source.sheet !== undefined) {
        occurrence.sheet = source.sheet;
      }
      group.occurrences.push(occurrence);
    }
  }

  // A key spelled the same way everywhere needs no mapping entry; only the
  // groups that actually disagree about spelling are worth reviewing.
  const groups = [...groupsByKey.values()].filter(
    (group) => group.spellings.length > 1,
  );

  return {
    groups,
    mapping: {
      version: COLUMN_MAPPING_VERSION,
      // Every spelling in a group already shares one normalized key, and a
      // canonical column matches its own key, so the entry needs no aliases:
      // listing them would repeat the same key and fail validation. The
      // spellings each entry folds are reported as group evidence instead.
      // Aliases are what a reviewer adds by hand for synonyms.
      columns: groups.map((group) => ({ name: group.canonical, aliases: [] })),
    },
  };
}
