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
import {
  calendarIsoText,
  isComponentsInRange,
  utcCalendarParts,
} from "./model/calendar.js";
import type { WorkbookModel } from "./model/index.js";
import type { WorksheetModel } from "./model/types.js";
import { WorkbookRead } from "./operations/read-model.js";
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
      // The serial is kept rather than turned into a `Date`. A `Date` has a
      // local face as well as a UTC one, and which of them is the workbook's
      // depends on how the engine composed it, so text built from one is text
      // built on somebody else's convention. Which cells are dates, and what
      // they hold, is the document model's answer: see `WorkbookDates`.
      cellDates: false,
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

/**
 * What a worksheet holds in the cells it declares or formats as dates, read
 * from the document model.
 *
 * The engine cannot answer this. A worksheet may say a cell is a date in two
 * ways: by wearing a date number format, or by declaring `t="d"` and writing
 * ISO 8601 text. Reading with `cellDates` off, which is what keeps a serial a
 * serial, turns the second kind into a plain number and drops the declaration
 * entirely: measured on this engine, `<c t="d"><v>2024-01-01</v></c>` arrives
 * as the number 45292 with no field left saying what it was. So the reader
 * asked the style alone, and an unstyled date cell became an integer, in the
 * table and in the schema an import inferred from it. Reading with `cellDates`
 * on keeps the declaration but hands back the engine's own parse of the text,
 * which remaps a year of 0099 to 1999 and normalises a month of 13 into
 * January of the next year: two defects the model refuses.
 *
 * So the model answers, for both kinds. It is the reader that sees the
 * declared type, it owns the style table, and its values come through the one
 * calendar route. The engine keeps the work it is good at: the used range, the
 * header row, and every cell that is not a date.
 */
export interface WorkbookDates {
  /** The dates one worksheet holds. */
  forSheet(sheet: string): SheetDates;
}

/**
 * What the model read in a date cell, and how the cell said it was one.
 *
 * The two ways differ in what the engine makes of the cell, so they differ in
 * how far this has to override it. A cell wearing a date format is a number the
 * engine formats correctly for display and reports as a serial: its stored
 * value comes from here, and its displayed text is the engine's, which is the
 * text the worksheet shows. A cell that declares `t="d"` is one the engine has
 * already turned into a serial, so both its value and its displayed text come
 * from here; the engine's display for it is that serial, which is a number
 * nobody wrote.
 */
interface CellDate {
  /**
   * One of the two answers a date cell has: the canonical timestamp, or the
   * cell's own text where that names no moment. Never blank - a cell holding
   * nothing is not a date cell, and `WorksheetModel.cellValue` reads it as no
   * value at all, so it never reaches this map and the reader's ordinary blank
   * rule answers for it.
   */
  readonly value: string;
  readonly declared: boolean;
}

/** A worksheet's date cells, by row and column. */
type WorksheetDates = ReadonlyMap<string, CellDate>;

/**
 * One worksheet's dates, asked for the way the engine indexes its cells: a
 * zero-based row, as `getCell` takes, rather than the one-based row a worksheet
 * writes. Converting here means no reader has to remember which of the two it
 * is holding.
 */
export type SheetDates = (
  rowIndex: number,
  columnIndex: number,
) => CellDate | undefined;

function dateKey(row: number, column: number): string {
  return `${row},${column}`;
}

/**
 * Collect a worksheet's date cells in one pass.
 *
 * A cell is a date when it declares itself one or when its style says so; both
 * questions are the model's, and the value is the model's too, so the reader
 * has one rule and one spelling rather than a second set of its own. A cell
 * the model reads as text - a `t="d"` cell whose text names no moment - is
 * kept as that text, because that is what the file says it is.
 */
function collectWorksheetDates(
  model: WorkbookModel,
  worksheet: WorksheetModel | undefined,
): WorksheetDates {
  const dates = new Map<string, CellDate>();
  if (!worksheet) {
    return dates;
  }
  for (const row of worksheet.rows()) {
    for (const cell of row.cells) {
      const declared = cell.type === DECLARED_DATE_TYPE;
      if (!declared && !model.isDateStyle(cell.styleIndex)) {
        continue;
      }
      const value = worksheet.cellValue(cell.ref);
      if (value instanceof Date) {
        dates.set(dateKey(cell.ref.row, cell.ref.column), {
          declared,
          value: calendarIsoText(utcCalendarParts(value)),
        });
      } else if (declared && typeof value === "string") {
        dates.set(dateKey(cell.ref.row, cell.ref.column), {
          declared,
          value,
        });
      }
    }
  }
  return dates;
}

/** The OOXML cell type a worksheet declares a date with. */
const DECLARED_DATE_TYPE = "d";

/**
 * The dates a workbook's bytes hold, for a reader that has bytes rather than a
 * loaded model. The worksheets are parsed on demand, so a reader that touches
 * one worksheet pays for one.
 */
export async function readWorkbookDates(
  bytes: Uint8Array,
  source: string,
  details: Record<string, unknown>,
): Promise<WorkbookDates> {
  return workbookDatesFrom(await WorkbookRead.load(bytes, { source, details }));
}

/**
 * The dates a workbook holds, read once per worksheet and only when a reader
 * asks for that worksheet.
 */
export function workbookDatesFrom(read: WorkbookRead): WorkbookDates {
  const bySheet = new Map<string, WorksheetDates>();
  return {
    forSheet(sheet) {
      return (rowIndex, columnIndex) => {
        let dates = bySheet.get(sheet);
        if (dates === undefined) {
          dates = collectWorksheetDates(read.workbook, read.worksheet(sheet));
          bySheet.set(sheet, dates);
        }
        return dates.get(dateKey(rowIndex + 1, columnIndex));
      };
    },
  };
}

/**
 * The same spelling for a moment that arrived already parsed.
 *
 * Unreachable from this package's own reading: `parseWorkbookBytes` keeps
 * serials, so no cell it produces carries a `Date`. It is here so a cell from a
 * workbook somebody else parsed with `cellDates` cannot fall through to
 * `String(v)` and arrive as a platform date string, and it goes through
 * `calendarIsoText` like everything else rather than through `toISOString`, so
 * there is one spelling of a date in this package and not two.
 *
 * The UTC face is the one read, because a moment composed from a workbook's
 * epoch and whole days wears the calendar date on that face and the local face
 * is that shifted by wherever the reader is sitting. A face outside the years
 * that can be written has no spelling, so it comes back undefined and the
 * caller falls through, the same rule the two paths above follow.
 */
function parsedDateText(value: Date): string | undefined {
  const parts = utcCalendarParts(value);
  return isComponentsInRange(parts) ? calendarIsoText(parts) : undefined;
}

function cellToPrimitive(
  cell: XLSX.CellObject | undefined,
  date: CellDate | undefined,
): CellValue {
  // A date the model read is the stored value, whether or not the engine saw a
  // cell here at all: a `t="d"` cell whose text names no moment reaches the
  // engine as nothing, and the text is what the file holds.
  if (date !== undefined) {
    return date.value;
  }
  if (!cell || cell.v === null || cell.v === undefined) {
    return null;
  }

  if (cell.v instanceof Date) {
    const parsed = parsedDateText(cell.v);
    if (parsed !== undefined) {
      return parsed;
    }
  }

  if (typeof cell.v === "number") {
    // A cell cannot hold one of these; the engine makes them out of text it
    // could not read, such as the spaces in a declared date cell. Blank is what
    // the cell holds, and blank is what every other reader here calls it.
    return Number.isFinite(cell.v) ? cell.v : null;
  }

  if (typeof cell.v === "string" || typeof cell.v === "boolean") {
    return cell.v;
  }

  return String(cell.w ?? cell.v);
}

function cellToDisplayText(
  cell: XLSX.CellObject | undefined,
  date: CellDate | undefined,
): string {
  // A cell that declares itself a date goes before the cached display text,
  // because that text is the serial the engine made of it rather than anything
  // the worksheet shows. A cell that only wears a date format does not: the
  // engine formats it the way the worksheet does, and displayed text is what
  // this reader promises.
  if (date?.declared === true) {
    return String(date.value);
  }
  if (!cell || cell.v === null || cell.v === undefined) {
    return "";
  }

  if (typeof cell.w === "string") {
    return cell.w;
  }

  if (cell.v instanceof Date) {
    const parsed = parsedDateText(cell.v);
    if (parsed !== undefined) {
      return parsed;
    }
  }

  if (date !== undefined) {
    // No cached display text: the stored date beats the raw serial, which is a
    // number nobody wrote into the cell.
    return String(date.value);
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
  dates: SheetDates,
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
        cellToPrimitive(
          getCell(worksheet, rowIndex, columnIndex),
          dates(rowIndex, columnIndex),
        ) !== null
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
  dates: SheetDates,
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
    dates,
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
      dates(headerRowIndex, columnIndex),
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
        dates(rowIndex, range.s.c + index),
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
  dates: SheetDates,
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
        dates(rowIndex, range.s.c + index),
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
  dates: SheetDates,
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
      dates(range.s.r, columnIndex),
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
        dates(rowIndex, range.s.c + index),
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
  workbookDates: WorkbookDates,
  sourceFile: string,
  options: ReadWorkbookOptions = {},
): WorksheetTableReport[] {
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
      // The workbook lists this worksheet and the engine produced nothing for
      // it, which it does silently: a cell it cannot parse, such as a declared
      // date holding no text at all, takes the whole part with it. Skipping
      // would drop a worksheet from every list this feeds - the tables, the
      // sheets an import offers - and say nothing, which is the one outcome a
      // reader must never produce. The document model reads such a worksheet;
      // until this reader takes its cells from there, the honest answer is to
      // stop and name it.
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_READ_FAILED,
        `Worksheet "${sheetName}" is listed in ${sourceFile} but could not be read from it, so what it holds is unknown.`,
        { details: { source: sourceFile, worksheet: sheetName } },
      );
    }
    const read = worksheetToTable(
      sourceFile,
      sheetName,
      worksheet,
      workbookDates.forSheet(sheetName),
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
  workbookDates: WorkbookDates,
  sourceFile: string,
  options: ReadWorkbookOptions = {},
): Table[] {
  const tables: Table[] = [];
  for (const report of workbookWorksheetReports(
    workbook,
    workbookDates,
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
  workbookDates: WorkbookDates,
  definitions: ExcelTableDefinition[],
  sourceFile: string,
  options: ReadWorkbookExcelTablesOptions = {},
): WorkbookExcelTable[] {
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
      workbookDates.forSheet(definition.sheet),
    );
    if (table) {
      tables.push(table);
    }
  }

  return tables;
}

export function workbookNamedRanges(
  workbook: XLSX.WorkBook,
  workbookDates: WorkbookDates,
  sourceFile: string,
  options: ReadWorkbookNamedRangesOptions = {},
): WorkbookNamedRange[] {
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
      workbookDates.forSheet(parsed.sheet),
    );
    if (table) {
      ranges.push(table);
    }
  }

  return ranges;
}

export function workbookWorksheetRecords(
  workbook: XLSX.WorkBook,
  workbookDates: WorkbookDates,
  options: ReadWorksheetRecordsOptions,
): WorksheetRecords {
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
  const dates = workbookDates.forSheet(worksheetName);
  const headerRowIndex = findHeaderRow(
    worksheet,
    range,
    dates,
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
      dates(headerRowIndex, columnIndex),
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
    const isEmpty = cells.every((cell, columnOffset) => {
      const value = cellToPrimitive(
        cell,
        dates(rowIndex, range.s.c + columnOffset),
      );
      return value === null || value === "";
    });
    if (isEmpty) {
      skippedEmptyRows += 1;
      continue;
    }

    const row: Record<string, string> = {};
    columns.forEach((column, columnOffset) => {
      row[column] = cellToDisplayText(
        cells[columnOffset],
        dates(rowIndex, range.s.c + columnOffset),
      );
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
  // The dates come from the document model, which is the only reader that sees
  // a cell declare itself a date rather than only wear a date format.
  const dates = workbookDatesFrom(
    await WorkbookRead.load(workbookBytes, {
      source: context.label,
      details: context.details ?? { source: context.label },
    }),
  );
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
      dates,
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
    availableNamedRanges = workbookNamedRanges(workbook, dates, context.file, {
      includeHiddenSheets: options.includeHiddenSheets,
      sheets,
    });
    tables = availableNamedRanges.filter(
      (namedRange) =>
        namedRange.rangeName.toLocaleLowerCase() ===
        options.range?.toLocaleLowerCase(),
    );
  } else {
    tables = workbookTables(workbook, dates, context.file, {
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
