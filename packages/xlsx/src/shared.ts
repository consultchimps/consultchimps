/**
 * Platform-neutral internals shared by the path-based and byte-based workbook
 * operations. This module must stay free of node:fs and node:path imports so
 * the byte entry point can run in browsers.
 */
import { ConsultChimpsError } from "@consultchimps/core";
import {
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
} from "@consultchimps/tabular";

import type { ExcelTableDefinition } from "./excel-tables.js";
import {
  decodeRange,
  type CellRectangle,
  type StreamedWorkbook,
} from "./operations/consolidate/reader.js";
import {
  openWorkbookBytes,
  readSheetGrid,
  SheetGrids,
  type SheetGrid,
} from "./operations/sheet-grid.js";
import { XLSX_ERRORS } from "./errors.js";
import { preserveWorkbookWithFilteredExcelTable } from "./preserve-table-split.js";
import {
  cellKey,
  countTitleRows,
  detectHeaderRow,
  isBlankValue,
  type RowValueCount,
} from "./region/header-detection.js";
import type { AllWorksheetSplitMetric } from "./split/all-worksheet.js";
import { splitOutputFilenames } from "./split/names.js";
import { stripPivotParts } from "./tier1/pivot.js";
import { CellError } from "./package/cell-error.js";
import { convertWorkbookToValues } from "./values-only.js";
import {
  cellWidthLength,
  tableColumnWidth,
  TableWorkbookWriter,
} from "./package/table-writer.js";

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

export type ConsolidateWorkbooksMetric =
  | "inputFiles"
  | "inputTables"
  | "outputColumns"
  | "outputRows"
  | "skippedSpacerColumns"
  | "skippedTitleRows"
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
export const WORKBOOK_DATE_TEXT: RegExp =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/u;

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
export function refuseNonTextDateColumns(
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
                  // An error cell reaches here as the error it holds.
                  valueType:
                    (value as unknown) instanceof CellError
                      ? "error"
                      : typeof value,
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
 * and the rows and columns it examined for values, all one-based except the
 * columns, which are zero-based as everywhere else in this package.
 *
 * It travels with the table so that anything asking a further question about
 * the same read asks it about the same rectangle. The alternative, resolving
 * the region a second time somewhere else, is how two answers about one
 * worksheet start disagreeing.
 *
 * The rectangle spans the worksheet's used columns, spacer columns included:
 * a spacer is a column the read examined and found nothing in, so it is
 * inside what the read covered even though the table carries no column for
 * it. `skippedSpacerColumns` on the report says how many there were.
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
  /**
   * Rows above the header row that held a value and were therefore left out
   * of the table: a report title, a "Prepared by" line, a merged banner. Zero
   * when the worksheet yielded no table.
   */
  skippedTitleRows: number;
  /**
   * Columns of the used range that held nothing in the header row or in any
   * row under it, and were therefore left out of the table. Zero when the
   * worksheet yielded no table.
   */
  skippedSpacerColumns: number;
}

export interface WorkbookNamedRange extends Table {
  rangeName: string;
  rangeRef: string;
}

/** What a cell holds, in the engine's zero-based numbering. */
type CellLookup = (rowIndex: number, columnIndex: number) => CellValue;

/**
 * One walk over the used range, in the engine's zero-based numbering: what
 * each row holds, for the header rule, and the last row each column holds a
 * value on, for the spacer rule. The header rule in
 * `src/region/header-detection.ts` reads the counts; it is the same rule the
 * region resolver applies to the document model, so the header row an
 * inspection reports is the one this reader keys on.
 */
interface RangeProfile {
  /**
   * One entry per row of the range holding a value, top to bottom. Blank rows
   * are left out: the rule reads them from the gaps in row numbers, and a
   * `!ref` padded to the bottom of the sheet by formatting would otherwise
   * cost an entry per empty row.
   */
  readonly counts: readonly RowValueCount[];
  /**
   * Per column offset from `range.startColumn`, the last row index holding a
   * value, or -1 when the column holds none. A column is a spacer for a header
   * row when its last value sits above that row.
   */
  readonly lastValueRow: readonly number[];
}

function profileRange(
  value: CellLookup,
  range: CellRectangle,
  merges: readonly CellRectangle[],
): RangeProfile {
  const width = range.endColumn - range.startColumn + 1;
  const counts: RowValueCount[] = [];
  const lastValueRow: number[] = new Array<number>(width).fill(-1);
  // The top-left cells of the merges spanning columns, in the engine's
  // zero-based numbering: a value there is a banner value.
  const banners = new Set<string>();
  for (const merge of merges) {
    if (merge.endColumn > merge.startColumn) {
      banners.add(cellKey(merge.startRow, merge.startColumn));
    }
  }
  for (let rowIndex = range.startRow; rowIndex <= range.endRow; rowIndex += 1) {
    let values = 0;
    let bannerValues = 0;
    for (let offset = 0; offset < width; offset += 1) {
      if (isBlankValue(value(rowIndex, range.startColumn + offset))) {
        continue;
      }
      values += 1;
      if (banners.has(cellKey(rowIndex, range.startColumn + offset))) {
        bannerValues += 1;
      }
      lastValueRow[offset] = rowIndex;
    }
    if (values > 0) {
      counts.push({ bannerValues, row: rowIndex, values });
    }
  }
  return { counts, lastValueRow };
}

/**
 * The name a header cell gives its column, or null for a blank cell. The
 * region layer's `headerCellName` spells the same cell the same way from the
 * document model, which is what lets an inspection promise the columns a
 * consolidation produces. An error cell is named by its text, `#REF!`.
 */
function headerCellName(value: CellValue): string | null {
  return isBlankValue(value) ? null : String(value);
}

/**
 * The column offsets a read from `headerRowIndex` keeps: every column holding
 * a value in the header row or under it. The rest are spacers.
 */
function keptColumnOffsets(
  profile: RangeProfile,
  headerRowIndex: number,
): number[] {
  const kept: number[] = [];
  profile.lastValueRow.forEach((lastRow, offset) => {
    if (lastRow >= headerRowIndex) {
      kept.push(offset);
    }
  });
  return kept;
}

/**
 * The zero-based index of a declared header row, validated and never
 * second-guessed, or undefined when none was declared. Decided before the
 * sheet is profiled, so a declared row the used range does not reach costs
 * nothing: a `!ref` padded to the bottom of the sheet by formatting is not
 * walked to refuse an option that was wrong on its face.
 */
export function declaredHeaderRowIndex(
  configuredRow?: number,
): number | undefined {
  if (configuredRow === undefined) {
    return undefined;
  }
  if (!Number.isInteger(configuredRow) || configuredRow < 1) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_HEADER_ROW,
      "The header row must be a positive integer.",
      { details: { configuredRow } },
    );
  }
  return configuredRow - 1;
}

/** Whether the worksheet named exactly `sheetName` is visible; one the workbook does not list counts as visible. */
function isVisibleSheet(
  workbook: StreamedWorkbook,
  sheetName: string,
): boolean {
  return (
    workbook.sheets.find((sheet) => sheet.name === sheetName)?.visible ?? true
  );
}

/** One worksheet read as a table, with what the read left out on the way. */
interface WorksheetTableRead {
  table: Table;
  region: WorksheetRegion;
  skippedTitleRows: number;
  skippedSpacerColumns: number;
}

function worksheetToTable(
  sourceFile: string,
  sheetName: string,
  grid: SheetGrid,
  configuredHeaderRow?: number,
): WorksheetTableRead | undefined {
  const range = grid.range;
  if (!range) {
    return undefined;
  }

  const declared = declaredHeaderRowIndex(configuredHeaderRow);
  if (
    declared !== undefined &&
    (declared < range.startRow || declared > range.endRow)
  ) {
    return undefined;
  }
  const profile = profileRange(grid.value, range, grid.merges);
  const headerRowIndex = declared ?? detectHeaderRow(profile.counts);
  if (headerRowIndex === undefined) {
    return undefined;
  }

  // Spacer columns are left out; the rest are named by their position among
  // the columns that were kept, which is what `column_3` means to a reader of
  // the output: the third column of this table.
  const width = range.endColumn - range.startColumn + 1;
  const keptOffsets = keptColumnOffsets(profile, headerRowIndex);
  const columns = uniqueHeaders(
    keptOffsets.map((offset) =>
      headerCellName(grid.value(headerRowIndex, range.startColumn + offset)),
    ),
  );

  const rows: TableRow[] = [];
  const sourceRows: number[] = [];
  for (
    let rowIndex = headerRowIndex + 1;
    rowIndex <= range.endRow;
    rowIndex += 1
  ) {
    // A row holding values only in spacer columns cannot exist: a column with
    // a value under the header is not a spacer. So a blank row here is blank
    // across the whole used width, as it was before spacers were left out.
    const values = keptOffsets.map((offset) =>
      grid.value(rowIndex, range.startColumn + offset),
    );
    if (values.every(isBlankValue)) {
      continue;
    }
    const row: TableRow = {};
    columns.forEach((column, position) => {
      row[column] = values[position] ?? null;
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
      lastRow: range.endRow + 1,
      startColumn: range.startColumn,
      endColumn: range.endColumn,
    },
    skippedSpacerColumns: width - keptOffsets.length,
    skippedTitleRows: countTitleRows(profile.counts, headerRowIndex),
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
  grid: SheetGrid,
): WorkbookExcelTable | undefined {
  // A lowercase reference names the same cells; the decoder reads capitals.
  const range = decodeRange(definition.range.toUpperCase());
  const rangeColumnCount = range.endColumn - range.startColumn + 1;
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
  const firstDataRowIndex = range.startRow + (definition.headerRow ? 1 : 0);
  const lastDataRowIndex = range.endRow - (definition.totalsRow ? 1 : 0);
  const rows: TableRow[] = [];
  const sourceRows: number[] = [];

  for (
    let rowIndex = firstDataRowIndex;
    rowIndex <= lastDataRowIndex;
    rowIndex += 1
  ) {
    const values = columns.map((_, index) =>
      grid.value(rowIndex, range.startColumn + index),
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
  // A lowercase reference names the same cells; the decoder reads capitals.
  return { range: range.replaceAll("$", "").toUpperCase(), sheet };
}

function namedRangeToTable(
  sourceFile: string,
  name: string,
  sheetName: string,
  rangeRef: string,
  grid: SheetGrid,
): WorkbookNamedRange | undefined {
  const range = decodeRange(rangeRef);
  const rawHeaders: Array<string | null> = [];
  for (
    let columnIndex = range.startColumn;
    columnIndex <= range.endColumn;
    columnIndex += 1
  ) {
    const value = grid.value(range.startRow, columnIndex);
    rawHeaders.push(value === null ? null : String(value));
  }
  const columns = uniqueHeaders(rawHeaders);

  const rows: TableRow[] = [];
  const sourceRows: number[] = [];
  for (
    let rowIndex = range.startRow + 1;
    rowIndex <= range.endRow;
    rowIndex += 1
  ) {
    const values = columns.map((_, index) =>
      grid.value(rowIndex, range.startColumn + index),
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
      firstDataRow: range.startRow + 2,
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
export async function workbookWorksheetReports(
  workbook: StreamedWorkbook,
  sourceFile: string,
  options: ReadWorkbookOptions = {},
): Promise<WorksheetTableReport[]> {
  const selectedSheets = lowercaseSet(options.sheets);
  const reports: WorksheetTableReport[] = [];

  for (const sheet of workbook.sheets) {
    if (!options.includeHiddenSheets && !sheet.visible) {
      continue;
    }
    if (selectedSheets && !selectedSheets.has(sheet.name.toLocaleLowerCase())) {
      continue;
    }

    // A worksheet the workbook lists but whose part cannot be found is
    // refused by the read, never skipped: skipping would drop it from every
    // list this feeds and say nothing.
    const read = worksheetToTable(
      sourceFile,
      sheet.name,
      await readSheetGrid(workbook, sheet),
      options.headerRow,
    );
    reports.push({
      sheet: sheet.name,
      table: read?.table,
      region: read?.region,
      skippedTitleRows: read?.skippedTitleRows ?? 0,
      skippedSpacerColumns: read?.skippedSpacerColumns ?? 0,
    });
  }

  return reports;
}

/** What a consolidation takes from a workbook's reports: the tables, and what the reads left out. */
export interface ConsolidationInputs {
  tables: Table[];
  skippedTitleRows: number;
  skippedSpacerColumns: number;
}

/**
 * The tables among `reports`, in worksheet order, with the title rows and
 * spacer columns their reads left out summed up, so both consolidation
 * surfaces report the same counts from the same reads.
 */
export function consolidationInputs(
  reports: readonly WorksheetTableReport[],
): ConsolidationInputs {
  const inputs: ConsolidationInputs = {
    tables: [],
    skippedTitleRows: 0,
    skippedSpacerColumns: 0,
  };
  for (const report of reports) {
    if (report.table) {
      inputs.tables.push(report.table);
    }
    inputs.skippedTitleRows += report.skippedTitleRows;
    inputs.skippedSpacerColumns += report.skippedSpacerColumns;
  }
  return inputs;
}

export async function workbookTables(
  workbook: StreamedWorkbook,
  sourceFile: string,
  options: ReadWorkbookOptions = {},
): Promise<Table[]> {
  return consolidationInputs(
    await workbookWorksheetReports(workbook, sourceFile, options),
  ).tables;
}

export async function workbookExcelTables(
  workbook: StreamedWorkbook,
  sourceFile: string,
  options: ReadWorkbookExcelTablesOptions = {},
): Promise<WorkbookExcelTable[]> {
  const selectedSheets = lowercaseSet(options.sheets);
  const selectedTables = lowercaseSet(options.tables);
  const grids = new SheetGrids(workbook);
  const tables: WorkbookExcelTable[] = [];

  for (const definition of workbook.tables) {
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

    const grid = await grids.named(definition.sheet);
    if (!grid) {
      continue;
    }
    const table = excelTableToTable(sourceFile, definition, grid);
    if (table) {
      tables.push(table);
    }
  }

  return tables;
}

export async function workbookNamedRanges(
  workbook: StreamedWorkbook,
  sourceFile: string,
  options: ReadWorkbookNamedRangesOptions = {},
): Promise<WorkbookNamedRange[]> {
  const selectedSheets = lowercaseSet(options.sheets);
  const selectedNames = lowercaseSet(options.names);
  const grids = new SheetGrids(workbook);
  const ranges: WorkbookNamedRange[] = [];

  for (const definedName of workbook.names) {
    if (
      !definedName.name ||
      definedName.name.startsWith(BUILTIN_DEFINED_NAME_PREFIX)
    ) {
      continue;
    }
    const parsed = parseNamedRangeRef(definedName.reference);
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
      !selectedNames.has(definedName.name.toLocaleLowerCase())
    ) {
      continue;
    }

    const grid = await grids.named(parsed.sheet);
    if (!grid) {
      continue;
    }
    const table = namedRangeToTable(
      sourceFile,
      definedName.name,
      parsed.sheet,
      parsed.range,
      grid,
    );
    if (table) {
      ranges.push(table);
    }
  }

  return ranges;
}

export async function workbookWorksheetRecords(
  workbook: StreamedWorkbook,
  options: ReadWorksheetRecordsOptions,
): Promise<WorksheetRecords> {
  const requestedWorksheet = options.worksheet?.trim();
  const sheetNames = workbook.sheets.map((sheet) => sheet.name);
  const sheet = requestedWorksheet
    ? workbook.sheets.find(
        (candidate) =>
          candidate.name.toLocaleLowerCase() ===
          requestedWorksheet.toLocaleLowerCase(),
      )
    : workbook.sheets[0];
  if (!sheet) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_WORKSHEET_NOT_FOUND,
      requestedWorksheet
        ? `Worksheet "${options.worksheet}" was not found in the workbook.`
        : "The workbook does not contain a worksheet.",
      {
        details: {
          availableWorksheets: sheetNames,
          worksheet: options.worksheet,
        },
      },
    );
  }

  const worksheetName = sheet.name;
  const grid = await readSheetGrid(workbook, sheet, { text: true });
  const range = grid.range;
  if (!range) {
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

  const declared = declaredHeaderRowIndex(options.headerRow);
  // A declared row outside the used range, on either side, is refused before
  // the sheet is profiled, for the reason `declaredHeaderRowIndex` gives.
  const profile =
    declared !== undefined &&
    (declared < range.startRow || declared > range.endRow)
      ? undefined
      : profileRange(grid.value, range, grid.merges);
  const headerRowIndex =
    profile === undefined
      ? undefined
      : (declared ?? detectHeaderRow(profile.counts));
  if (
    profile === undefined ||
    headerRowIndex === undefined ||
    headerRowIndex > range.endRow
  ) {
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

  // Only the columns that hold something - by the same test the table reader
  // applies, so the two readers keep the same columns - are held to the rule
  // that a header must name them. A spacer is not an unnamed column.
  const keptOffsets = keptColumnOffsets(profile, headerRowIndex);
  const columns: string[] = [];
  for (const offset of keptOffsets) {
    const header = grid.text(headerRowIndex, range.startColumn + offset).trim();
    if (!header) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_EMPTY_HEADER,
        `Worksheet "${worksheetName}" contains an empty column header.`,
        {
          details: {
            column: range.startColumn + offset + 1,
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
    rowIndex <= range.endRow;
    rowIndex += 1
  ) {
    // Spacer columns hold nothing under the header by definition, so a row
    // blank across the kept columns is blank across the whole used width.
    const isEmpty = keptOffsets.every((offset) =>
      isBlankValue(grid.value(rowIndex, range.startColumn + offset)),
    );
    if (isEmpty) {
      skippedEmptyRows += 1;
      continue;
    }
    const row: Record<string, string> = {};
    columns.forEach((column, position) => {
      row[column] = grid.text(
        rowIndex,
        range.startColumn + keptOffsets[position]!,
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

  assertSheetName(sheetName);

  const widths = table.columns.map((column) =>
    tableColumnWidth(
      table.rows.reduce(
        (length, row) => Math.max(length, cellWidthLength(row[column] ?? null)),
        column.length,
      ),
    ),
  );
  const chunks: Uint8Array[] = [];
  let size = 0;
  const writer = new TableWorkbookWriter({
    sheetName,
    columns: table.columns,
    widths,
    rowCount: table.rows.length,
    onChunk: (chunk) => {
      chunks.push(chunk);
      size += chunk.length;
    },
  });
  for (const row of table.rows) {
    writer.writeRow(table.columns.map((column) => row[column] ?? null));
  }
  writer.finish();

  const bytes = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.length;
  }
  return bytes;
}

// Excel's own rules for a worksheet name, which a workbook that breaks them
// fails to open with. SheetJS enforced them with an unexplained error before
// the writer moved off it.
const SHEET_NAME_FORBIDDEN = /[\\/?*[\]:]/u;
// Excel accepts no control characters in a name, and a tab or line feed would
// not survive the XML attribute that carries it: parsers turn it into a space.
// eslint-disable-next-line no-control-regex -- the pattern exists to find control characters.
const SHEET_NAME_CONTROL = /[\u0000-\u001F\u007F]/u;

export function assertSheetName(sheetName: string): void {
  const problem =
    sheetName.length === 0
      ? "is empty"
      : sheetName.length > 31
        ? "is longer than 31 characters"
        : SHEET_NAME_FORBIDDEN.test(sheetName)
          ? "contains one of \\ / ? * [ ] :"
          : SHEET_NAME_CONTROL.test(sheetName)
            ? "contains a control character such as a tab or line break"
            : sheetName.startsWith("'") || sheetName.endsWith("'")
              ? "starts or ends with an apostrophe"
              : sheetName.toLowerCase() === "history"
                ? 'is "History", which Excel reserves'
                : undefined;
  if (problem !== undefined) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_SHEET_NAME,
      `The worksheet name "${sheetName}" ${problem}; Excel cannot open a workbook with that name.`,
      { details: { sheetName } },
    );
  }
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
  const workbook = await openWorkbookBytes(workbookBytes, {
    file: context.file,
    source: context.label,
    details: context.details,
  });
  const sheets = options.sheet ? [options.sheet] : undefined;

  let definitions: ExcelTableDefinition[] = [];
  let availableExcelTables: WorkbookExcelTable[] = [];
  let availableNamedRanges: WorkbookNamedRange[] = [];
  let tables: Table[];

  if (options.table) {
    definitions = [...workbook.tables];
    availableExcelTables = await workbookExcelTables(workbook, context.file, {
      includeHiddenSheets: options.includeHiddenSheets,
      sheets,
    });
    tables = availableExcelTables.filter(
      (table) =>
        table.excelTableName.toLocaleLowerCase() ===
        options.table?.toLocaleLowerCase(),
    );
  } else if (options.range) {
    availableNamedRanges = await workbookNamedRanges(workbook, context.file, {
      includeHiddenSheets: options.includeHiddenSheets,
      sheets,
    });
    tables = availableNamedRanges.filter(
      (namedRange) =>
        namedRange.rangeName.toLocaleLowerCase() ===
        options.range?.toLocaleLowerCase(),
    );
  } else {
    tables = await workbookTables(workbook, context.file, {
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
