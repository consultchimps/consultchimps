/**
 * Platform-neutral internals shared by the path-based and byte-based workbook
 * operations. This module must stay free of node:fs and node:path imports so
 * the byte entry point can run in browsers.
 */
import { ConsultChimpsError } from "@consultchimps/core";
import {
  type CellValue,
  groupTableByColumn,
  type Table,
  type TableRow,
  uniqueHeaders,
  unionTables,
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
  "inputFiles" | "inputTables" | "outputColumns" | "outputRows";
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
   * Match columns whose headers differ only in case, spacing, or punctuation
   * (for example "Failed Checks" and "Failed_Checks") instead of requiring
   * the exact same header in every worksheet.
   */
  normalizeHeaders?: boolean | undefined;
}

/**
 * The consolidation core both surfaces call: stack every worksheet table read
 * from the inputs into one union table. Keeping the refusal and the union in
 * one place is what makes the file API and the byte API produce the same
 * columns, the same row order, and the same bytes for the same workbooks.
 */
export function consolidateTables(
  tables: Table[],
  options: ConsolidateTablesOptions = {},
): Table {
  if (tables.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_TABLES,
      "No visible, non-empty worksheets were found in the input workbooks.",
    );
  }

  return unionTables(tables, {
    addSourceColumns: options.addSourceColumns,
    normalizeHeaders: options.normalizeHeaders,
  });
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
      cellDates: true,
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

function cellToPrimitive(cell: XLSX.CellObject | undefined): CellValue {
  if (!cell || cell.v === null || cell.v === undefined) {
    return null;
  }

  if (cell.v instanceof Date) {
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

function cellToDisplayText(cell: XLSX.CellObject | undefined): string {
  if (!cell || cell.v === null || cell.v === undefined) {
    return "";
  }

  if (typeof cell.w === "string") {
    return cell.w;
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
      if (cellToPrimitive(getCell(worksheet, rowIndex, columnIndex)) !== null) {
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
  configuredHeaderRow?: number,
): Table | undefined {
  const reference = worksheet["!ref"];
  if (!reference) {
    return undefined;
  }

  const range = XLSX.utils.decode_range(reference);
  const headerRowIndex = findHeaderRow(worksheet, range, configuredHeaderRow);
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
      cellToPrimitive(getCell(worksheet, rowIndex, range.s.c + index)),
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
    source: {
      file: sourceFile,
      firstDataRow: headerRowIndex + 2,
      sheet: sheetName,
    },
  };
}

function excelTableToTable(
  sourceFile: string,
  definition: ExcelTableDefinition,
  worksheet: XLSX.WorkSheet,
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
      cellToPrimitive(getCell(worksheet, rowIndex, range.s.c + index)),
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
    const value = cellToPrimitive(getCell(worksheet, range.s.r, columnIndex));
    rawHeaders.push(value === null ? null : String(value));
  }
  const columns = uniqueHeaders(rawHeaders);

  const rows: TableRow[] = [];
  const sourceRows: number[] = [];
  for (let rowIndex = range.s.r + 1; rowIndex <= range.e.r; rowIndex += 1) {
    const values = columns.map((_, index) =>
      cellToPrimitive(getCell(worksheet, rowIndex, range.s.c + index)),
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

export function workbookTables(
  workbook: XLSX.WorkBook,
  sourceFile: string,
  options: ReadWorkbookOptions = {},
): Table[] {
  const selectedSheets = lowercaseSet(options.sheets);
  const tables: Table[] = [];

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
    const table = worksheetToTable(
      sourceFile,
      sheetName,
      worksheet,
      options.headerRow,
    );
    if (table) {
      tables.push(table);
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
    const table = excelTableToTable(sourceFile, definition, worksheet);
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
  const headerRowIndex = findHeaderRow(worksheet, range, options.headerRow);
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
      const value = cellToPrimitive(cell);
      return value === null || value === "";
    });
    if (isEmpty) {
      skippedEmptyRows += 1;
      continue;
    }

    const row: Record<string, string> = {};
    columns.forEach((column, columnOffset) => {
      row[column] = cellToDisplayText(cells[columnOffset]);
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
