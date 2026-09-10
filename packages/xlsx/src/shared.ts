/**
 * Platform-neutral internals shared by the path-based and byte-based workbook
 * operations. This module must stay free of node:fs and node:path imports so
 * the byte entry point can run in browsers.
 */
import { ConsultChimpsError } from "@consultchimps/core";
import {
  applyColumnMappingToTables,
  type CellValue,
  type ColumnHeaderSource,
  type ColumnMapping,
  type ColumnMappingSuggestion,
  groupTableByColumn,
  normalizedColumnKey,
  suggestColumnMapping,
  type Table,
  type TableRow,
  uniqueHeaders,
  unionTables,
  validateColumnMapping,
} from "@consultchimps/tabular";
import * as XLSX from "xlsx";

import {
  type ExcelTableDefinition,
  readExcelTableDefinitions,
} from "./excel-tables.js";
import { XLSX_ERRORS } from "./errors.js";
import { preserveWorkbookWithFilteredExcelTable } from "./preserve-table-split.js";
import type { AllWorksheetSplitMetric } from "./split/all-worksheet.js";
import { splitOutputFilenames } from "./split/names.js";
import { stripPivotParts } from "./tier1/pivot.js";
import { convertWorkbookToValues } from "./values-only.js";

export const WORKBOOK_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
export const MACRO_WORKBOOK_MEDIA_TYPE =
  "application/vnd.ms-excel.sheet.macroEnabled.12";
export const CONSOLIDATE_OPERATION = "sheets.consolidate";
/** The worksheet a consolidation writes into unless the caller names another. */
export const CONSOLIDATED_SHEET_NAME = "Consolidated";
/**
 * Workbook inspection: the glossary's single verb for describing an input's
 * structure without producing files, matching `pptx.inspect-template`. The
 * operation itself lives in `src/operations/describe.ts`; only its name is
 * here, beside the other operation constants.
 */
export const INSPECT_OPERATION = "sheets.inspect";
export const MERGE_OPERATION = "sheets.merge";
export const SPLIT_OPERATION = "sheets.split-by-column";
export const WORKBOOK_EXTENSION = ".xlsx";
export const MACRO_WORKBOOK_EXTENSION = ".xlsm";

/** Whether an output name asks for a macro-enabled workbook. */
export function isMacroWorkbookName(name: string): boolean {
  return name.toLocaleLowerCase().endsWith(MACRO_WORKBOOK_EXTENSION);
}

// Identical inputs must produce byte-identical outputs, so generated workbooks
// carry fixed document timestamps instead of the current time.
const FIXED_WORKBOOK_DATE = new Date(0);
const WORKBOOK_CREATOR = "ConsultChimps";

export type ConsolidateWorkbooksMetric =
  | "inputFiles"
  | "inputTables"
  | "outputColumns"
  | "outputRows"
  | "suggestedColumns"
  | "unmappedColumns";
export type ConsolidateWorkbooksPlanMetric = "inputFiles" | "outputFiles";
export type MergeWorkbooksMetric =
  "hiddenSheets" | "inputFiles" | "outputSheets";
/**
 * One metric vocabulary for every split, whichever engine ran and whichever
 * surface asked. A single-source split reports zero for the work only the
 * all-worksheet engine does, rather than omitting the key, so a caller can read
 * a metric without first asking which mode produced the result.
 */
export type SplitWorkbookByColumnMetric = AllWorksheetSplitMetric;
export type SplitWorkbookByColumnPlanMetric = Exclude<
  SplitWorkbookByColumnMetric,
  "outputRows"
>;

export interface ReadWorkbookOptions {
  headerRow?: number | undefined;
  includeHiddenSheets?: boolean | undefined;
  sheets?: string[] | undefined;
}

/**
 * Hand the event loop back for one full turn.
 *
 * Cancellation reaches a Web Worker as a posted message, and a message is a
 * macrotask: an operation that runs to completion without ever yielding one -
 * however many `await`s it contains, because awaiting an already-resolved
 * value only drains microtasks - has already posted its output before the
 * worker dequeues the `cancel`. So an operation whose expensive steps are
 * synchronous must yield a macrotask between them, and check its signal after
 * each yield, or its Cancel button does nothing.
 *
 * Operations built on JSZip get this for free, since loading and generating a
 * package yield on their own; the ones that build a workbook synchronously do
 * not.
 */
export function yieldToEventLoop(): Promise<void> {
  return new Promise<void>((resolve) => {
    setTimeout(resolve, 0);
  });
}

export interface ConsolidateTablesOptions {
  addSourceColumns?: boolean | undefined;
  /**
   * Fold the source headers into canonical columns before the union. The
   * mapping is matched by normalized column key whatever `normalizeHeaders`
   * says, so the two options never compete: see {@link consolidateTables}.
   */
  mapping?: ColumnMapping | undefined;
  /**
   * Match columns whose headers differ only in case, spacing, or punctuation
   * (for example "Failed Checks" and "Failed_Checks") instead of requiring
   * the exact same header in every worksheet.
   */
  normalizeHeaders?: boolean | undefined;
}

/** One consolidated table, plus the columns no mapping entry claimed. */
export interface ConsolidateTablesResult {
  table: Table;
  /**
   * Every distinct unmapped spelling across the inputs, in first-seen order.
   * Empty when no mapping was applied.
   */
  unmappedColumns: string[];
}

/** The portable filename a byte-level mapping draft is offered under. */
export const SUGGESTED_MAPPING_FILE_NAME = "mapping-draft.json";
/** The media type a written column mapping carries as an artifact. */
export const MAPPING_MEDIA_TYPE = "application/json";

/**
 * Describe the worksheet a table came from, for error messages: " in sheet
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

function tableSourceDetails(table: Table): Record<string, unknown> {
  const details: Record<string, unknown> = {};
  if (table.source?.file !== undefined) {
    details["file"] = table.source.file;
  }
  if (table.source?.sheet !== undefined) {
    details["sheet"] = table.source.sheet;
  }
  return details;
}

function tableSourceRowNumber(table: Table, index: number): number {
  return table.sourceRows?.[index] ?? (table.source?.firstDataRow ?? 2) + index;
}

/**
 * The exact shape `cellToPrimitive` writes for a cell the workbook stores as a
 * real date, and nothing a person types into a cell: `workbookDateText` writes
 * the full ISO 8601 timestamp for every such cell, whether or not it carries a
 * time. Matching it is how a date coercion can tell "the workbook already holds
 * this as a date" from ordinary text it should try to parse, which is also why
 * the shape stays one shape.
 */
const WORKBOOK_DATE_TEXT = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/u;

/**
 * Refuse a declared date coercion the worksheet cannot honestly satisfy.
 *
 * A mapping's date coercion parses a column's *text*, written in a format the
 * mapping declares. A worksheet hands that column two other shapes for the
 * same idea, and neither one is that text:
 *
 * - A number. Excel stores a date as a count of days from an epoch, and which
 *   day the count starts from is a property of the workbook, not of the value:
 *   the 1900 and 1904 date systems sit 1462 days apart, and the 1900 system
 *   deliberately reproduces Lotus 1-2-3's non-existent 29 February 1900, so
 *   serial 60 names no calendar day at all. The table model a mapping sees
 *   carries no epoch, and a bare number is indistinguishable from a case
 *   number, a quantity, or an amount. Converting one would be a guess dressed
 *   as a conversion, so the run stops instead.
 * - A date the workbook already stores as a date, which reaches the table as a
 *   full ISO 8601 instant. There is no text in the declared format to parse,
 *   and the value is already unambiguous, so the honest answer is to drop the
 *   coercion rather than to reinterpret the instant.
 *
 * Either refuses before anything is written, naming the file, the worksheet,
 * the column, and the row, so the message says which cell to look at.
 */
function refuseNonTextDateColumns(
  tables: Table[],
  mapping: ColumnMapping,
): void {
  const dateColumnByKey = new Map<string, { format: string; name: string }>();
  for (const column of mapping.columns) {
    if (column.coercion?.type !== "date") {
      continue;
    }
    const entry = { format: column.coercion.format, name: column.name };
    for (const spelling of [column.name, ...column.aliases]) {
      dateColumnByKey.set(normalizedColumnKey(spelling), entry);
    }
  }
  if (dateColumnByKey.size === 0) {
    return;
  }

  for (const table of tables) {
    const dateColumns = table.columns
      .map((column) => ({
        canonical: dateColumnByKey.get(normalizedColumnKey(column)),
        column,
      }))
      .filter(
        (
          entry,
        ): entry is {
          canonical: { format: string; name: string };
          column: string;
        } => entry.canonical !== undefined,
      );
    if (dateColumns.length === 0) {
      continue;
    }

    table.rows.forEach((row, index) => {
      for (const { canonical, column } of dateColumns) {
        const value = Object.hasOwn(row, column) ? (row[column] ?? null) : null;
        // A blank cell stays blank, and ordinary text is exactly what the
        // coercion is for; everything else is refused below.
        const isSourceText =
          typeof value === "string" && !WORKBOOK_DATE_TEXT.test(value);
        if (value === null || isSourceText) {
          continue;
        }

        const location = `Row ${tableSourceRowNumber(table, index)} of column "${column}"${describeTableSource(table)}`;
        const declared = `the canonical column "${canonical.name}" declares a date coercion, which reads text written in the format "${canonical.format}"`;
        const problem =
          typeof value === "number"
            ? {
                message: `${location} holds the number ${value}, and ${declared}. Excel counts a date as a number of days, and the day that count starts from belongs to the workbook rather than to the cell, so a bare number cannot be read as a date without guessing. Format that column as text in the source workbook, or remove the date coercion for "${canonical.name}". Nothing was written.`,
                valueType: "number",
              }
            : typeof value === "string"
              ? {
                  message: `${location} holds a value the workbook stores as a date rather than as text, and ${declared}. A cell the workbook already holds as a date needs no coercion. Remove the date coercion for "${canonical.name}" and the value carries through. Nothing was written.`,
                  valueType: "workbook-date",
                }
              : {
                  message: `${location} does not hold text, and ${declared}. Correct that cell, or remove the date coercion for "${canonical.name}". Nothing was written.`,
                  valueType: typeof value,
                };

        throw new ConsultChimpsError(
          XLSX_ERRORS.XLSX_MAPPING_DATE_NOT_TEXT,
          problem.message,
          {
            details: {
              canonicalColumn: canonical.name,
              column,
              format: canonical.format,
              row: tableSourceRowNumber(table, index),
              valueType: problem.valueType,
              ...tableSourceDetails(table),
            },
          },
        );
      }
    });
  }
}

/**
 * The consolidation core both surfaces call: fold the source headers into
 * canonical columns when a mapping is supplied, then stack every worksheet
 * table read from the inputs into one union table. Keeping the refusal, the
 * mapping, and the union in one place is what makes the file API and the byte
 * API produce the same columns, the same row order, and the same bytes for the
 * same workbooks.
 *
 * The mapping is applied per source table, before the union, so that two
 * columns of one worksheet folding into a single canonical column is caught as
 * that worksheet's ambiguity rather than silently merged across the inputs.
 *
 * `mapping` and `normalizeHeaders` do not compete. A mapping always matches by
 * normalized column key, so the flag changes nothing about which source header
 * reaches which canonical column; canonical names are written verbatim and are
 * identical in every mapped table, so they union exactly either way.
 * `normalizeHeaders` continues to govern only how the columns the mapping did
 * not claim are matched against each other.
 */
export function consolidateTables(
  tables: Table[],
  options: ConsolidateTablesOptions = {},
): ConsolidateTablesResult {
  if (tables.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_TABLES,
      "No visible, non-empty worksheets were found in the input workbooks.",
    );
  }

  const unionOptions = {
    addSourceColumns: options.addSourceColumns,
    normalizeHeaders: options.normalizeHeaders,
  };

  if (!options.mapping) {
    return { table: unionTables(tables, unionOptions), unmappedColumns: [] };
  }

  const mapping = validateColumnMapping(options.mapping);
  refuseNonTextDateColumns(tables, mapping);
  const mapped = applyColumnMappingToTables(tables, mapping);
  return {
    table: unionTables(mapped.tables, unionOptions),
    unmappedColumns: mapped.unmappedColumns,
  };
}

/**
 * Draft a mapping from the headers of the tables a consolidation just read,
 * carrying each table's file and worksheet through as the evidence a reviewer
 * needs to judge a proposed group.
 */
export function suggestMappingForTables(
  tables: Table[],
): ColumnMappingSuggestion {
  return suggestColumnMapping(
    tables.map((table) => {
      const source: ColumnHeaderSource = { columns: table.columns };
      if (table.source?.file !== undefined) {
        source.file = table.source.file;
      }
      if (table.source?.sheet !== undefined) {
        source.sheet = table.source.sheet;
      }
      return source;
    }),
  );
}

/**
 * Render a mapping as the JSON document both surfaces hand back, indented so a
 * reviewer can edit it and ending in a newline so it appends cleanly. Identical
 * inputs produce identical bytes, as every other output of this package does.
 */
export function serializeColumnMapping(mapping: ColumnMapping): string {
  return `${JSON.stringify(mapping, null, 2)}\n`;
}

/**
 * The warning that keeps an unmapped column loud but lossless: it passed
 * through under its own name, and here is its name.
 */
export function unmappedColumnsWarning(columns: string[]): string {
  const one = columns.length === 1;
  return `${columns.length} column${one ? "" : "s"} did not match the column mapping and ${
    one ? "kept its own name" : "kept their own names"
  }: ${columns.map((column) => `"${column}"`).join(", ")}. Add ${
    one ? "it" : "them"
  } to the mapping if ${one ? "it belongs" : "they belong"} in a canonical column.`;
}

/**
 * Refuse a run that both applies a mapping and drafts one.
 *
 * ADR 0002 fixed neither reading: drafting from the source headers ignores the
 * mapping that is being applied, and drafting from the mapped headers proposes
 * a second review of the first one's output. Both are defensible, so the
 * toolkit declines to pick one on the caller's behalf. A refusal stays
 * reversible; a guess would become a compatibility promise.
 */
export function refuseMappingWithSuggestion(): void {
  throw new ConsultChimpsError(
    XLSX_ERRORS.XLSX_MAPPING_SUGGEST_CONFLICT,
    "Choose either a column mapping to apply or a drafted mapping to review, not both in one run. Applying a mapping and drafting one describe two different reviews of the same headers, and ConsultChimps will not decide which was meant. Run the consolidation twice if both are wanted.",
    { details: { problem: "mapping_and_suggestion" } },
  );
}

export interface ReadWorkbookExcelTablesOptions {
  includeHiddenSheets?: boolean | undefined;
  sheets?: string[] | undefined;
  tables?: string[] | undefined;
}

export interface ReadWorkbookNamedRangesOptions {
  includeHiddenSheets?: boolean | undefined;
  names?: string[] | undefined;
  sheets?: string[] | undefined;
}

export interface ReadWorksheetRecordsOptions {
  headerRow?: number | undefined;
  worksheet?: string | undefined;
}

export interface WorksheetRecords {
  columns: string[];
  rows: Array<Record<string, string>>;
  skippedEmptyRows: number;
  sourceRows: number[];
  worksheet: string;
}

export interface WorkbookExcelTable extends Table {
  excelTableName: string;
  excelTableRange: string;
}

/**
 * The rectangle a worksheet read actually covered: the header row it keyed on
 * and the rows and columns it took values from, all one-based except the
 * columns, which are zero-based as everywhere else in this package.
 *
 * It travels with the table so that anything asking a further question about
 * the same read asks it about the same rectangle. The alternative, resolving
 * the region a second time somewhere else, is how two answers about one
 * worksheet start disagreeing.
 */
export interface WorksheetRegion {
  readonly headerRow: number;
  readonly lastRow: number;
  readonly startColumn: number;
  readonly endColumn: number;
}

/** One worksheet as the table reader saw it, whether or not it yielded a table. */
export interface WorksheetTableReport {
  /** The worksheet this came from. */
  sheet: string;
  /**
   * The table, or undefined when the worksheet holds no header row with rows of
   * values under it, which is the condition `workbookTables` filters on.
   */
  table: Table | undefined;
  /** The rectangle the table was read from, absent when there is no table. */
  region: WorksheetRegion | undefined;
}

export interface WorkbookNamedRange extends Table {
  rangeName: string;
  rangeRef: string;
}

export interface ParseWorkbookOptions {
  /** Machine-readable context added to a read failure. */
  details?: Record<string, unknown> | undefined;
  /** Keep cached display text, which worksheet records report verbatim. */
  cellText?: boolean | undefined;
}

/**
 * Parse workbook bytes, reporting an unreadable workbook as a stable error.
 * The source label appears in the message, so callers pass a file path or an
 * in-memory input name.
 */
export function parseWorkbookBytes(
  workbookBytes: Uint8Array,
  source: string,
  options: ParseWorkbookOptions = {},
): XLSX.WorkBook {
  try {
    return XLSX.read(workbookBytes, {
      // The serial is kept rather than turned into a `Date`, and the number
      // format is kept so a date-formatted serial can be told from a quantity.
      // See `workbookDateText`: a `Date` has a local face as well as a UTC one,
      // and which of them is the workbook's depends on how the engine composed
      // it, so text built from one is text built on somebody else's convention.
      cellDates: false,
      cellNF: true,
      cellText: options.cellText ?? false,
      dense: false,
      type: "array",
    });
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${source}`,
      {
        cause: error,
        details: options.details ?? { source },
      },
    );
  }
}

/**
 * Read the Excel Table definitions from a workbook package, reporting an
 * unreadable package with the same stable error as a failed workbook parse.
 */
export async function parseExcelTableDefinitions(
  workbookBytes: Uint8Array,
  source: string,
  details?: Record<string, unknown>,
): Promise<ExcelTableDefinition[]> {
  try {
    return await readExcelTableDefinitions(workbookBytes);
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${source}`,
      {
        cause: error,
        details: details ?? { source },
      },
    );
  }
}

/** The calendar components a serial decodes to. */
interface SerialDateParts {
  y: number;
  m: number;
  d: number;
  H: number;
  M: number;
  S: number;
  /** Fraction of a second, from 0 up to but not including 1. */
  u: number;
}

/**
 * The number-format functions this reader calls directly.
 *
 * They are the engine's own serial decoding, and they are arithmetic:
 * `parse_date_code` turns a serial into calendar components with no `Date`
 * anywhere in the middle. The package's types do not declare them, so the
 * shape this reader depends on is written out here rather than assumed.
 */
const SSF = XLSX.SSF as unknown as {
  is_date(format: string): boolean;
  parse_date_code(
    serial: number,
    options?: { date1904?: boolean },
  ): SerialDateParts | null | undefined;
};

/** Whether the workbook counts its dates from 1904 rather than from 1900. */
function workbookDateSystem(workbook: XLSX.WorkBook): boolean {
  return workbook.Workbook?.WBProps?.date1904 === true;
}

function pad(value: number, width: number): string {
  return String(value).padStart(width, "0");
}

/**
 * The ISO 8601 text for a date-formatted serial.
 *
 * A workbook stores a date as a count of days from an epoch, and which day the
 * count starts from is a property of the workbook rather than of the cell. So
 * the text is computed from those two things and nothing else: the serial and
 * the workbook's date system, decoded by the engine's own `parse_date_code`,
 * with the components written out directly.
 *
 * No `Date` takes part. A `Date` has a local face as well as a UTC one, and
 * which of them carries the calendar date depends on how whoever built it
 * chose to compose it. Text built from a `Date` is therefore text built on
 * that choice, and where the choice was the local face the same workbook read
 * as a different calendar day in every time zone. This is a function of the
 * workbook, so every reader in every zone gets the same characters.
 *
 * One shape, always the full timestamp, whether or not the cell carries a
 * time. A date column then reads the same way down its whole length rather
 * than changing shape at the first cell that happens to carry an hour, and the
 * `WORKBOOK_DATE_TEXT` rule above can still tell a value the workbook holds as
 * a date from ordinary text somebody typed, which a bare `2024-01-01` could
 * not. `@consultchimps/db` accepts both spellings in a `date` column, so the
 * choice costs a consumer nothing.
 *
 * A serial the engine cannot decode into a calendar day, such as the day-zero
 * serial 0, comes back undefined and the caller keeps the number. Serial 60 is
 * decoded rather than refused: the 1900 system deliberately reproduces a
 * spreadsheet-era bug in which 29 February 1900 exists, and that is the day
 * Excel shows for it.
 */
function workbookDateText(
  serial: number,
  date1904: boolean,
): string | undefined {
  const parts = SSF.parse_date_code(serial, { date1904 });
  if (!parts || parts.m < 1 || parts.m > 12 || parts.d < 1 || parts.d > 31) {
    return undefined;
  }
  // Clamped rather than allowed to carry: a fraction that rounds to a whole
  // second would otherwise write ".1000" where a millisecond field belongs.
  const milliseconds = Math.min(999, Math.round((parts.u || 0) * 1000));
  return `${pad(parts.y, 4)}-${pad(parts.m, 2)}-${pad(parts.d, 2)}T${pad(
    parts.H,
    2,
  )}:${pad(parts.M, 2)}:${pad(parts.S, 2)}.${pad(milliseconds, 3)}Z`;
}

/**
 * The date text for a cell, or undefined when the cell is not one the workbook
 * stores as a date.
 *
 * A stored date is a number wearing a date number format, and the format is
 * the only thing that separates it from a case number or a quantity, which is
 * why this reader keeps the format rather than only the value.
 */
function cellDateText(
  cell: XLSX.CellObject,
  date1904: boolean,
): string | undefined {
  if (cell.t !== "n" || typeof cell.v !== "number") {
    return undefined;
  }
  const format = cell.z;
  return typeof format === "string" && SSF.is_date(format)
    ? workbookDateText(cell.v, date1904)
    : undefined;
}

function cellToPrimitive(
  cell: XLSX.CellObject | undefined,
  date1904: boolean,
): CellValue {
  if (!cell || cell.v === null || cell.v === undefined) {
    return null;
  }

  const dateText = cellDateText(cell, date1904);
  if (dateText !== undefined) {
    return dateText;
  }

  if (cell.v instanceof Date) {
    // This reader keeps serials, so it produces none. The branch is here so a
    // cell from a workbook somebody else parsed with `cellDates` cannot fall
    // through to `String(v)` and arrive as a platform date string.
    return cell.v.toISOString();
  }

  if (
    typeof cell.v === "string" ||
    typeof cell.v === "number" ||
    typeof cell.v === "boolean"
  ) {
    return cell.v;
  }

  return String(cell.w ?? cell.v);
}

function cellToDisplayText(
  cell: XLSX.CellObject | undefined,
  date1904: boolean,
): string {
  if (!cell || cell.v === null || cell.v === undefined) {
    return "";
  }

  if (typeof cell.w === "string") {
    return cell.w;
  }

  // No cached display text: a stored date would otherwise report as the raw
  // serial, which is a number nobody wrote into the cell.
  const dateText = cellDateText(cell, date1904);
  if (dateText !== undefined) {
    return dateText;
  }

  if (cell.v instanceof Date) {
    return cell.v.toISOString();
  }

  if (typeof cell.v === "boolean") {
    return cell.v ? "TRUE" : "FALSE";
  }

  return String(cell.v);
}

function getCell(
  worksheet: XLSX.WorkSheet,
  rowIndex: number,
  columnIndex: number,
): XLSX.CellObject | undefined {
  return worksheet[XLSX.utils.encode_cell({ c: columnIndex, r: rowIndex })] as
    XLSX.CellObject | undefined;
}

function findHeaderRow(
  worksheet: XLSX.WorkSheet,
  range: XLSX.Range,
  date1904: boolean,
  configuredRow?: number,
): number | undefined {
  if (configuredRow !== undefined) {
    if (!Number.isInteger(configuredRow) || configuredRow < 1) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_INVALID_HEADER_ROW,
        "The header row must be a positive integer.",
        { details: { configuredRow } },
      );
    }
    return configuredRow - 1;
  }

  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    for (
      let columnIndex = range.s.c;
      columnIndex <= range.e.c;
      columnIndex += 1
    ) {
      if (
        cellToPrimitive(getCell(worksheet, rowIndex, columnIndex), date1904) !==
        null
      ) {
        return rowIndex;
      }
    }
  }

  return undefined;
}

function isVisibleSheet(workbook: XLSX.WorkBook, sheetName: string): boolean {
  const metadata = workbook.Workbook?.Sheets?.find(
    (sheet) => sheet.name === sheetName,
  );
  return (metadata?.Hidden ?? 0) === 0;
}

function worksheetToTable(
  sourceFile: string,
  sheetName: string,
  worksheet: XLSX.WorkSheet,
  date1904: boolean,
  configuredHeaderRow?: number,
): { table: Table; region: WorksheetRegion } | undefined {
  const reference = worksheet["!ref"];
  if (!reference) {
    return undefined;
  }

  const range = XLSX.utils.decode_range(reference);
  const headerRowIndex = findHeaderRow(
    worksheet,
    range,
    date1904,
    configuredHeaderRow,
  );
  if (
    headerRowIndex === undefined ||
    headerRowIndex < range.s.r ||
    headerRowIndex > range.e.r
  ) {
    return undefined;
  }

  const rawHeaders: Array<string | null> = [];
  for (
    let columnIndex = range.s.c;
    columnIndex <= range.e.c;
    columnIndex += 1
  ) {
    const value = cellToPrimitive(
      getCell(worksheet, headerRowIndex, columnIndex),
      date1904,
    );
    rawHeaders.push(value === null ? null : String(value));
  }

  const columns = uniqueHeaders(rawHeaders);
  const rows: TableRow[] = [];
  const sourceRows: number[] = [];

  for (
    let rowIndex = headerRowIndex + 1;
    rowIndex <= range.e.r;
    rowIndex += 1
  ) {
    const values = columns.map((_, index) =>
      cellToPrimitive(
        getCell(worksheet, rowIndex, range.s.c + index),
        date1904,
      ),
    );

    if (values.every((value) => value === null || value === "")) {
      continue;
    }

    const row: TableRow = {};
    columns.forEach((column, index) => {
      row[column] = values[index] ?? null;
    });
    rows.push(row);
    sourceRows.push(rowIndex + 1);
  }

  if (rows.length === 0) {
    return undefined;
  }

  return {
    // The rectangle is exactly the one the rows above were read from, reported
    // rather than recomputed, so a later question about this read cannot be
    // asked of a different region.
    region: {
      headerRow: headerRowIndex + 1,
      lastRow: range.e.r + 1,
      startColumn: range.s.c,
      endColumn: range.e.c,
    },
    table: {
      columns,
      rows,
      sourceRows,
      source: {
        file: sourceFile,
        firstDataRow: headerRowIndex + 2,
        sheet: sheetName,
      },
    },
  };
}

function excelTableToTable(
  sourceFile: string,
  definition: ExcelTableDefinition,
  worksheet: XLSX.WorkSheet,
  date1904: boolean,
): WorkbookExcelTable | undefined {
  let range: XLSX.Range;
  try {
    range = XLSX.utils.decode_range(definition.range);
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_EXCEL_TABLE,
      `Excel Table "${definition.name}" has an invalid range.`,
      {
        cause: error,
        details: {
          range: definition.range,
          sheet: definition.sheet,
          table: definition.name,
        },
      },
    );
  }

  const rangeColumnCount = range.e.c - range.s.c + 1;
  if (rangeColumnCount !== definition.columns.length) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_EXCEL_TABLE,
      `Excel Table "${definition.name}" has inconsistent column metadata.`,
      {
        details: {
          columnCount: definition.columns.length,
          range: definition.range,
          rangeColumnCount,
          sheet: definition.sheet,
          table: definition.name,
        },
      },
    );
  }

  const columns = uniqueHeaders(
    definition.columns.map((column) => column || null),
  );
  const firstDataRowIndex = range.s.r + (definition.headerRow ? 1 : 0);
  const lastDataRowIndex = range.e.r - (definition.totalsRow ? 1 : 0);
  const rows: TableRow[] = [];
  const sourceRows: number[] = [];

  for (
    let rowIndex = firstDataRowIndex;
    rowIndex <= lastDataRowIndex;
    rowIndex += 1
  ) {
    const values = columns.map((_, index) =>
      cellToPrimitive(
        getCell(worksheet, rowIndex, range.s.c + index),
        date1904,
      ),
    );

    if (values.every((value) => value === null || value === "")) {
      continue;
    }

    const row: TableRow = {};
    columns.forEach((column, index) => {
      row[column] = values[index] ?? null;
    });
    rows.push(row);
    sourceRows.push(rowIndex + 1);
  }

  if (rows.length === 0) {
    return undefined;
  }

  return {
    columns,
    excelTableName: definition.name,
    excelTableRange: definition.range,
    rows,
    sourceRows,
    source: {
      file: sourceFile,
      firstDataRow: firstDataRowIndex + 1,
      sheet: definition.sheet,
    },
  };
}

const BUILTIN_DEFINED_NAME_PREFIX = "_xlnm.";
const NAMED_RANGE_REF_PATTERN =
  /^(?:'(?<quotedSheet>(?:[^']|'')+)'|(?<sheet>[^'!,:]+))!(?<range>\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$/u;

function parseNamedRangeRef(
  ref: string,
): { range: string; sheet: string } | undefined {
  const match = NAMED_RANGE_REF_PATTERN.exec(ref.trim());
  const sheet =
    match?.groups?.quotedSheet?.replaceAll("''", "'") ?? match?.groups?.sheet;
  const range = match?.groups?.range;
  if (!sheet || !range) {
    return undefined;
  }
  return { range: range.replaceAll("$", ""), sheet };
}

function namedRangeToTable(
  sourceFile: string,
  name: string,
  sheetName: string,
  rangeRef: string,
  worksheet: XLSX.WorkSheet,
  date1904: boolean,
): WorkbookNamedRange | undefined {
  let range: XLSX.Range;
  try {
    range = XLSX.utils.decode_range(rangeRef);
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_NAMED_RANGE,
      `Named range "${name}" has an invalid cell reference.`,
      {
        cause: error,
        details: { name, range: rangeRef, sheet: sheetName },
      },
    );
  }

  const rawHeaders: Array<string | null> = [];
  for (
    let columnIndex = range.s.c;
    columnIndex <= range.e.c;
    columnIndex += 1
  ) {
    const value = cellToPrimitive(
      getCell(worksheet, range.s.r, columnIndex),
      date1904,
    );
    rawHeaders.push(value === null ? null : String(value));
  }
  const columns = uniqueHeaders(rawHeaders);

  const rows: TableRow[] = [];
  const sourceRows: number[] = [];
  for (let rowIndex = range.s.r + 1; rowIndex <= range.e.r; rowIndex += 1) {
    const values = columns.map((_, index) =>
      cellToPrimitive(
        getCell(worksheet, rowIndex, range.s.c + index),
        date1904,
      ),
    );
    if (values.every((value) => value === null || value === "")) {
      continue;
    }

    const row: TableRow = {};
    columns.forEach((column, index) => {
      row[column] = values[index] ?? null;
    });
    rows.push(row);
    sourceRows.push(rowIndex + 1);
  }

  if (rows.length === 0) {
    return undefined;
  }

  return {
    columns,
    rows,
    sourceRows,
    rangeName: name,
    rangeRef,
    source: {
      file: sourceFile,
      firstDataRow: range.s.r + 2,
      sheet: sheetName,
    },
  };
}

function lowercaseSet(values: string[] | undefined): Set<string> | undefined {
  return values
    ? new Set(values.map((value) => value.toLocaleLowerCase()))
    : undefined;
}

/**
 * Read every selected worksheet, reporting each one whether or not it yielded a
 * table, with the rectangle each read covered. `workbookTables` is this list
 * with the tables taken out of it, so the two can never describe the same
 * worksheet differently.
 */
export function workbookWorksheetReports(
  workbook: XLSX.WorkBook,
  sourceFile: string,
  options: ReadWorkbookOptions = {},
): WorksheetTableReport[] {
  const date1904 = workbookDateSystem(workbook);
  const selectedSheets = lowercaseSet(options.sheets);
  const reports: WorksheetTableReport[] = [];

  for (const sheetName of workbook.SheetNames) {
    if (!options.includeHiddenSheets && !isVisibleSheet(workbook, sheetName)) {
      continue;
    }
    if (selectedSheets && !selectedSheets.has(sheetName.toLocaleLowerCase())) {
      continue;
    }

    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      continue;
    }
    const read = worksheetToTable(
      sourceFile,
      sheetName,
      worksheet,
      date1904,
      options.headerRow,
    );
    reports.push({
      sheet: sheetName,
      table: read?.table,
      region: read?.region,
    });
  }

  return reports;
}

export function workbookTables(
  workbook: XLSX.WorkBook,
  sourceFile: string,
  options: ReadWorkbookOptions = {},
): Table[] {
  const tables: Table[] = [];
  for (const report of workbookWorksheetReports(
    workbook,
    sourceFile,
    options,
  )) {
    if (report.table) {
      tables.push(report.table);
    }
  }
  return tables;
}

export function workbookExcelTables(
  workbook: XLSX.WorkBook,
  definitions: ExcelTableDefinition[],
  sourceFile: string,
  options: ReadWorkbookExcelTablesOptions = {},
): WorkbookExcelTable[] {
  const date1904 = workbookDateSystem(workbook);
  const selectedSheets = lowercaseSet(options.sheets);
  const selectedTables = lowercaseSet(options.tables);
  const tables: WorkbookExcelTable[] = [];

  for (const definition of definitions) {
    if (
      !options.includeHiddenSheets &&
      !isVisibleSheet(workbook, definition.sheet)
    ) {
      continue;
    }
    if (
      selectedSheets &&
      !selectedSheets.has(definition.sheet.toLocaleLowerCase())
    ) {
      continue;
    }
    if (
      selectedTables &&
      !selectedTables.has(definition.name.toLocaleLowerCase())
    ) {
      continue;
    }

    const worksheet = workbook.Sheets[definition.sheet];
    if (!worksheet) {
      continue;
    }
    const table = excelTableToTable(
      sourceFile,
      definition,
      worksheet,
      date1904,
    );
    if (table) {
      tables.push(table);
    }
  }

  return tables;
}

export function workbookNamedRanges(
  workbook: XLSX.WorkBook,
  sourceFile: string,
  options: ReadWorkbookNamedRangesOptions = {},
): WorkbookNamedRange[] {
  const date1904 = workbookDateSystem(workbook);
  const selectedSheets = lowercaseSet(options.sheets);
  const selectedNames = lowercaseSet(options.names);
  const ranges: WorkbookNamedRange[] = [];

  for (const definedName of workbook.Workbook?.Names ?? []) {
    if (
      !definedName.Name ||
      definedName.Name.startsWith(BUILTIN_DEFINED_NAME_PREFIX)
    ) {
      continue;
    }
    const parsed = parseNamedRangeRef(definedName.Ref ?? "");
    if (!parsed) {
      continue;
    }
    if (
      !options.includeHiddenSheets &&
      !isVisibleSheet(workbook, parsed.sheet)
    ) {
      continue;
    }
    if (
      selectedSheets &&
      !selectedSheets.has(parsed.sheet.toLocaleLowerCase())
    ) {
      continue;
    }
    if (
      selectedNames &&
      !selectedNames.has(definedName.Name.toLocaleLowerCase())
    ) {
      continue;
    }

    const worksheet = workbook.Sheets[parsed.sheet];
    if (!worksheet) {
      continue;
    }
    const table = namedRangeToTable(
      sourceFile,
      definedName.Name,
      parsed.sheet,
      parsed.range,
      worksheet,
      date1904,
    );
    if (table) {
      ranges.push(table);
    }
  }

  return ranges;
}

export function workbookWorksheetRecords(
  workbook: XLSX.WorkBook,
  options: ReadWorksheetRecordsOptions,
): WorksheetRecords {
  const date1904 = workbookDateSystem(workbook);
  const requestedWorksheet = options.worksheet?.trim();
  const worksheetName = requestedWorksheet
    ? workbook.SheetNames.find(
        (candidate) =>
          candidate.toLocaleLowerCase() ===
          requestedWorksheet.toLocaleLowerCase(),
      )
    : workbook.SheetNames[0];
  if (!worksheetName) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_WORKSHEET_NOT_FOUND,
      requestedWorksheet
        ? `Worksheet "${options.worksheet}" was not found in the workbook.`
        : "The workbook does not contain a worksheet.",
      {
        details: {
          availableWorksheets: workbook.SheetNames,
          worksheet: options.worksheet,
        },
      },
    );
  }

  const worksheet = workbook.Sheets[worksheetName];
  const reference = worksheet?.["!ref"];
  if (!worksheet || !reference) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_HEADER_ROW,
      `Worksheet "${worksheetName}" does not contain a header row.`,
      {
        details: {
          headerRow: options.headerRow,
          worksheet: worksheetName,
        },
      },
    );
  }

  const range = XLSX.utils.decode_range(reference);
  const headerRowIndex = findHeaderRow(
    worksheet,
    range,
    date1904,
    options.headerRow,
  );
  if (headerRowIndex === undefined || headerRowIndex > range.e.r) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_HEADER_ROW,
      `Worksheet "${worksheetName}" does not contain the selected header row.`,
      {
        details: {
          headerRow: options.headerRow,
          worksheet: worksheetName,
        },
      },
    );
  }

  const columns: string[] = [];
  for (
    let columnIndex = range.s.c;
    columnIndex <= range.e.c;
    columnIndex += 1
  ) {
    const header = cellToDisplayText(
      getCell(worksheet, headerRowIndex, columnIndex),
      date1904,
    ).trim();
    if (!header) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_EMPTY_HEADER,
        `Worksheet "${worksheetName}" contains an empty column header.`,
        {
          details: {
            column: columnIndex + 1,
            headerRow: headerRowIndex + 1,
            worksheet: worksheetName,
          },
        },
      );
    }
    columns.push(header);
  }

  const duplicateHeaders = columns.filter(
    (column, index) => columns.indexOf(column) !== index,
  );
  if (duplicateHeaders.length > 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_DUPLICATE_HEADER,
      `Worksheet "${worksheetName}" contains duplicate column headers.`,
      {
        details: {
          duplicateHeaders: [...new Set(duplicateHeaders)],
          headerRow: headerRowIndex + 1,
          worksheet: worksheetName,
        },
      },
    );
  }

  const rows: Array<Record<string, string>> = [];
  const sourceRows: number[] = [];
  let skippedEmptyRows = 0;

  for (
    let rowIndex = headerRowIndex + 1;
    rowIndex <= range.e.r;
    rowIndex += 1
  ) {
    const cells = columns.map((_, columnOffset) =>
      getCell(worksheet, rowIndex, range.s.c + columnOffset),
    );
    const isEmpty = cells.every((cell) => {
      const value = cellToPrimitive(cell, date1904);
      return value === null || value === "";
    });
    if (isEmpty) {
      skippedEmptyRows += 1;
      continue;
    }

    const row: Record<string, string> = {};
    columns.forEach((column, columnOffset) => {
      row[column] = cellToDisplayText(cells[columnOffset], date1904);
    });
    rows.push(row);
    sourceRows.push(rowIndex + 1);
  }

  return {
    columns,
    rows,
    skippedEmptyRows,
    sourceRows,
    worksheet: worksheetName,
  };
}

/** Pin document metadata so identical inputs serialize to identical bytes. */
function applyDeterministicProperties(workbook: XLSX.WorkBook): void {
  workbook.Props = {
    Author: WORKBOOK_CREATOR,
    CreatedDate: FIXED_WORKBOOK_DATE,
    LastAuthor: WORKBOOK_CREATOR,
    ModifiedDate: FIXED_WORKBOOK_DATE,
  };
}

function serializeWorkbook(workbook: XLSX.WorkBook): Uint8Array {
  applyDeterministicProperties(workbook);
  return new Uint8Array(
    XLSX.write(workbook, {
      bookType: "xlsx",
      // Deduplicate repeated text through the workbook's shared-strings table,
      // as Excel itself does. Without it every cell carries its own text, so
      // repetitive tables serialize considerably larger than their inputs. The
      // table is built in first-encounter order, which keeps identical inputs
      // producing byte-identical outputs.
      bookSST: true,
      compression: true,
      type: "array",
    }) as ArrayBuffer,
  );
}

/** Build a single-worksheet workbook holding one table's values. */
export function buildTableWorkbookBytes(
  table: Table,
  sheetName: string,
): Uint8Array {
  if (table.columns.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_COLUMNS,
      "Cannot write a table with no columns.",
    );
  }

  const data: CellValue[][] = [
    table.columns,
    ...table.rows.map((row) =>
      table.columns.map((column) => row[column] ?? null),
    ),
  ];
  const worksheet = XLSX.utils.aoa_to_sheet(data);
  worksheet["!autofilter"] = {
    ref: XLSX.utils.encode_range({
      s: { c: 0, r: 0 },
      e: {
        c: table.columns.length - 1,
        r: table.rows.length,
      },
    }),
  };
  worksheet["!cols"] = table.columns.map((column) => {
    const longest = table.rows.reduce(
      (length, row) => Math.max(length, String(row[column] ?? "").length),
      column.length,
    );
    return { wch: Math.min(Math.max(longest + 2, 10), 60) };
  });

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  return serializeWorkbook(workbook);
}

// The worksheet merge lives in `src/merge/`: it is a part-level transplant on
// the L0 package rather than a rebuild through a spreadsheet library, so it
// shares nothing with the readers above beyond the operation constants.
export {
  appendWorkbookSheets,
  createMergeState,
  finishMergedWorkbook,
  type MergedWorkbook,
  type MergeWorkbooksBuildOptions,
  type MergeWorkbooksState,
} from "./merge/transplant.js";

export interface SplitSelectionOptions {
  column: string;
  headerRow?: number | undefined;
  includeBlank?: boolean | undefined;
  includeHiddenSheets?: boolean | undefined;
  preserveWorkbook?: boolean | undefined;
  range?: string | undefined;
  sheet?: string | undefined;
  table?: string | undefined;
}

/**
 * Reject option combinations that cannot be satisfied and report whether the
 * split keeps the complete source workbook.
 */
export function resolvePreserveWorkbook(
  options: SplitSelectionOptions,
): boolean {
  if (options.table && options.range) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_TABLE_RANGE_CONFLICT,
      "Choose either an Excel Table or a named range as the data source, not both.",
      {
        details: {
          range: options.range,
          table: options.table,
        },
      },
    );
  }
  if (options.table && options.headerRow !== undefined) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_TABLE_HEADER_ROW,
      "The headerRow option cannot be used with an Excel Table; the table defines its own headers.",
      {
        details: {
          headerRow: options.headerRow,
          table: options.table,
        },
      },
    );
  }
  if (options.range && options.headerRow !== undefined) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_RANGE_HEADER_ROW,
      "The headerRow option cannot be used with a named range; the range's first row provides the headers.",
      {
        details: {
          headerRow: options.headerRow,
          range: options.range,
        },
      },
    );
  }
  if (options.preserveWorkbook === true && !options.table) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_PRESERVE_REQUIRES_TABLE,
      "The preserveWorkbook option requires a named Excel Table; it is not available for named ranges or plain worksheet splits.",
      {
        details: {
          preserveWorkbook: options.preserveWorkbook,
          range: options.range,
          table: options.table,
        },
      },
    );
  }

  // Table splits preserve the complete workbook unless explicitly disabled.
  return options.preserveWorkbook ?? options.table !== undefined;
}

export interface SplitSourceContext {
  /** Human-readable source used in messages: a file path or an input name. */
  label: string;
  /** The name recorded as a table's source file. */
  file: string;
  /** Machine-readable source context added to error details. */
  details: Record<string, unknown>;
}

export interface ResolvedSplitSource {
  grouped: ReturnType<typeof groupTableByColumn>;
  preservedTableDefinition: ExcelTableDefinition | undefined;
  preserveWorkbook: boolean;
  table: Table;
}

/**
 * Select the single table a split reads from, group its rows, and locate the
 * package definition a preserved split rewrites.
 */
export async function resolveSplitSource(
  workbookBytes: Uint8Array,
  context: SplitSourceContext,
  options: SplitSelectionOptions,
): Promise<ResolvedSplitSource> {
  const preserveWorkbook = resolvePreserveWorkbook(options);
  const workbook = parseWorkbookBytes(workbookBytes, context.label, {
    cellText: options.range !== undefined,
    details: context.details,
  });
  const sheets = options.sheet ? [options.sheet] : undefined;

  let definitions: ExcelTableDefinition[] = [];
  let availableExcelTables: WorkbookExcelTable[] = [];
  let availableNamedRanges: WorkbookNamedRange[] = [];
  let tables: Table[];

  if (options.table) {
    definitions = await parseExcelTableDefinitions(
      workbookBytes,
      context.label,
      context.details,
    );
    availableExcelTables = workbookExcelTables(
      workbook,
      definitions,
      context.file,
      { includeHiddenSheets: options.includeHiddenSheets, sheets },
    );
    tables = availableExcelTables.filter(
      (table) =>
        table.excelTableName.toLocaleLowerCase() ===
        options.table?.toLocaleLowerCase(),
    );
  } else if (options.range) {
    availableNamedRanges = workbookNamedRanges(workbook, context.file, {
      includeHiddenSheets: options.includeHiddenSheets,
      sheets,
    });
    tables = availableNamedRanges.filter(
      (namedRange) =>
        namedRange.rangeName.toLocaleLowerCase() ===
        options.range?.toLocaleLowerCase(),
    );
  } else {
    tables = workbookTables(workbook, context.file, {
      headerRow: options.headerRow,
      includeHiddenSheets: options.includeHiddenSheets,
      sheets,
    });
  }

  if (tables.length === 0) {
    const selectedSource = options.table
      ? `Excel Table "${options.table}"`
      : options.range
        ? `Named range "${options.range}"`
        : options.sheet
          ? `Worksheet "${options.sheet}"`
          : undefined;
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_NO_TABLE,
      selectedSource
        ? `${selectedSource} was not found or has no data rows.`
        : "No visible, non-empty worksheet was found in the input workbook.",
      {
        details: {
          availableRanges: availableNamedRanges.map((namedRange) => ({
            name: namedRange.rangeName,
            sheet: namedRange.source?.sheet,
          })),
          availableTables: availableExcelTables.map((table) => ({
            name: table.excelTableName,
            sheet: table.source?.sheet,
          })),
          ...context.details,
          range: options.range,
          sheet: options.sheet,
          table: options.table,
        },
      },
    );
  }

  if (tables.length > 1) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_MULTIPLE_TABLES,
      options.table
        ? `Excel Table "${options.table}" was found on multiple worksheets; choose one with the sheet option.`
        : options.range
          ? `Named range "${options.range}" is defined more than once; choose a worksheet with the sheet option.`
          : "The workbook contains multiple non-empty worksheets; choose one with the sheet option.",
      {
        details: {
          availableSheets: tables
            .map((table) => table.source?.sheet)
            .filter((sheet) => sheet !== undefined),
          ...context.details,
        },
      },
    );
  }

  const table = tables[0];
  if (!table) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_NO_TABLE,
      "No worksheet table was available to split.",
      { details: { ...context.details } },
    );
  }

  const grouped = groupTableByColumn(table, options.column, {
    includeBlank: options.includeBlank,
  });
  if (grouped.groups.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_NO_GROUPS,
      `No output groups remain for column "${grouped.column}".`,
      {
        details: {
          column: grouped.column,
          includeBlank: options.includeBlank ?? true,
          ...context.details,
        },
      },
    );
  }

  const preservedTableDefinition = preserveWorkbook
    ? definitions.find(
        (definition) =>
          definition.name.toLocaleLowerCase() ===
            options.table?.toLocaleLowerCase() &&
          definition.sheet.toLocaleLowerCase() ===
            table.source?.sheet?.toLocaleLowerCase(),
      )
    : undefined;
  if (preserveWorkbook && !preservedTableDefinition) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_PRESERVE_TABLE_NOT_FOUND,
      `Excel Table "${options.table}" could not be located in the workbook package.`,
      {
        details: {
          ...context.details,
          sheet: table.source?.sheet,
          table: options.table,
        },
      },
    );
  }

  return { grouped, preservedTableDefinition, preserveWorkbook, table };
}

/**
 * Prepare the workbook bytes a preserved split rewrites for every group,
 * optionally replacing formulas with their cached values first.
 */
export async function preservedSplitTemplateBytes(
  workbookBytes: Uint8Array,
  values: boolean | undefined,
): Promise<Uint8Array> {
  return values ? convertWorkbookToValues(workbookBytes) : workbookBytes;
}

/**
 * Produce one group's workbook, preserving the source package when asked.
 *
 * Both branches are deliberately still off the layered engine after Phase 1.
 * The preserved branch's contract is a refusal, not a repair (see
 * `preserve-table-split.ts`). The rebuilding branch does not edit a workbook at
 * all: it writes a fresh single-worksheet package from parsed cell values, so
 * every structure the corpus tracks is absent by construction rather than lost
 * by accident, and there is no row to relocate. Migrating it would mean
 * changing what a compact split *produces*, which is a decision for the phase
 * that makes it, not a side effect of moving the split engine.
 */
export async function buildSplitGroupBytes(
  group: { table: Table },
  context: {
    /** Called with the pivot tables removed from this group's output. */
    onPivotTablesRemoved?: ((removed: number) => void) | undefined;
    preservedTableDefinition: ExcelTableDefinition | undefined;
    sheetName: string;
    templateBytes: Uint8Array | undefined;
  },
): Promise<Uint8Array> {
  if (context.templateBytes && context.preservedTableDefinition) {
    const preserved = await preserveWorkbookWithFilteredExcelTable(
      context.templateBytes,
      {
        definition: context.preservedTableDefinition,
        sourceRows: group.table.sourceRows ?? [],
      },
    );
    // Tier-1 wiring: the preserved path copies the source package, pivot caches
    // included, so this group's recipient would receive every other group's
    // rows inside the cache. The rebuilding path below cannot leak them because
    // it writes a fresh package from parsed cell values.
    const stripped = await stripPivotParts(preserved);
    context.onPivotTablesRemoved?.(stripped.removedPivotTables);
    return stripped.bytes;
  }
  return buildTableWorkbookBytes(group.table, context.sheetName);
}

export function skippedRowsWarning(
  grouped: ReturnType<typeof groupTableByColumn>,
): string {
  return `Skipped ${grouped.skippedRows} row${
    grouped.skippedRows === 1 ? "" : "s"
  } with blank values in "${grouped.column}".`;
}

/**
 * Derive one portable output filename per group value, disambiguating values
 * that sanitize to the same name.
 *
 * The naming rules themselves live in `split/names.ts`, because a byte split
 * returns these filenames and a file split joins the same names onto a
 * directory; only the two surfaces' default prefixes differ.
 */
export function splitOutputFileNames(
  filenamePrefix: string,
  values: CellValue[],
  extension: string = WORKBOOK_EXTENSION,
): string[] {
  return splitOutputFilenames(filenamePrefix, values, extension);
}

export function withoutWorkbookExtension(name: string): string {
  return name.replace(/\.xls[xm]$/iu, "");
}

export { safeNameFragment } from "@consultchimps/core";
