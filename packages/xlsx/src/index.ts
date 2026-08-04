import {
  mkdtemp,
  readFile,
  rename,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import path from "node:path";

import {
  ConsultChimpsError,
  throwIfAborted,
  type OperationControlOptions,
  type OperationPlan,
  type OperationResult,
  type PlannedOutput,
} from "@consultchimps/core";
import {
  ensureDirectory,
  ensureOutputAvailable,
  ensureParentDirectory,
  pathExists,
  refuseInputOverwrite,
} from "@consultchimps/files";
import {
  type CellValue,
  groupTableByColumn,
  type Table,
  type TableRow,
  unionTables,
  uniqueHeaders,
} from "@consultchimps/tabular";
import * as XLSX from "xlsx";

import {
  type ExcelTableDefinition,
  readExcelTableDefinitions,
} from "./excel-tables.js";
import { XLSX_ERRORS } from "./errors.js";
import { preserveWorkbookWithFilteredExcelTable } from "./preserve-table-split.js";

export { XLSX_ERRORS, type XlsxErrorCode } from "./errors.js";

export type ConsolidateWorkbooksMetric =
  "inputFiles" | "inputTables" | "outputColumns" | "outputRows";
export type ConsolidateWorkbooksPlanMetric = "inputFiles" | "outputFiles";
export type MergeWorkbooksMetric =
  "hiddenSheets" | "inputFiles" | "outputSheets";
export type SplitWorkbookByColumnMetric =
  | "groups"
  | "inputFiles"
  | "inputRows"
  | "outputFiles"
  | "outputRows"
  | "skippedRows";
export type SplitWorkbookByColumnPlanMetric = Exclude<
  SplitWorkbookByColumnMetric,
  "outputRows"
>;

export interface ReadWorkbookOptions {
  headerRow?: number | undefined;
  includeHiddenSheets?: boolean | undefined;
  sheets?: string[] | undefined;
}

export interface ReadWorkbookExcelTablesOptions {
  includeHiddenSheets?: boolean | undefined;
  sheets?: string[] | undefined;
  tables?: string[] | undefined;
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

export interface ConsolidateWorkbooksOptions
  extends ReadWorkbookOptions, OperationControlOptions {
  inputs: string[];
  output: string;
  addSourceColumns?: boolean | undefined;
  outputSheetName?: string | undefined;
  overwrite?: boolean | undefined;
}

export interface MergeWorkbooksOptions {
  includeSheetIndex?: boolean | undefined;
  overwrite?: boolean | undefined;
}

export interface SplitWorkbookByColumnOptions extends OperationControlOptions {
  input: string;
  outputDirectory: string;
  column: string;
  filenamePrefix?: string | undefined;
  headerRow?: number | undefined;
  includeBlank?: boolean | undefined;
  includeHiddenSheets?: boolean | undefined;
  overwrite?: boolean | undefined;
  /**
   * Keep the complete source workbook and replace only the selected Excel
   * Table's rows. Defaults to true when a table is selected; not available
   * for named ranges or plain worksheet splits.
   */
  preserveWorkbook?: boolean | undefined;
  range?: string | undefined;
  sheet?: string | undefined;
  table?: string | undefined;
}

export interface WorkbookExcelTable extends Table {
  excelTableName: string;
  excelTableRange: string;
}

export interface WorkbookNamedRange extends Table {
  rangeName: string;
  rangeRef: string;
}

export interface ReadWorkbookNamedRangesOptions {
  includeHiddenSheets?: boolean | undefined;
  names?: string[] | undefined;
  sheets?: string[] | undefined;
}

export interface WriteTableOptions {
  overwrite?: boolean | undefined;
  sheetName?: string | undefined;
}

const UNSAFE_FILENAME_CHARACTERS = /[<>:"/\\|?*]+/gu;
const WINDOWS_RESERVED_FILENAME =
  /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\..*)?$/iu;

function withoutControlCharacters(value: string): string {
  return [...value]
    .filter((character) => {
      const codePoint = character.codePointAt(0);
      return codePoint !== undefined && codePoint >= 32 && codePoint !== 127;
    })
    .join("");
}

function safeFilenameSegment(value: string, fallback: string): string {
  const normalized = withoutControlCharacters(value.normalize("NFKC"))
    .replace(UNSAFE_FILENAME_CHARACTERS, "-")
    .replace(/\s+/gu, " ")
    .replace(/-+/gu, "-")
    .trim()
    .replace(/[. ]+$/gu, "");
  const limited = [...normalized]
    .slice(0, 80)
    .join("")
    .replace(/[. ]+$/gu, "");
  const safe = limited || fallback;

  return WINDOWS_RESERVED_FILENAME.test(safe) ? `_${safe}` : safe;
}

function groupValueFilenameSegment(value: CellValue): string {
  if (value === null) {
    return "blank";
  }

  return safeFilenameSegment(String(value), "value");
}

function splitOutputPaths(
  outputDirectory: string,
  filenamePrefix: string,
  values: CellValue[],
): string[] {
  const usedFilenames = new Set<string>();

  return values.map((value) => {
    const segment = groupValueFilenameSegment(value);
    const base = `${filenamePrefix}-${segment}`;
    let filename = `${base}.xlsx`;
    let suffix = 2;

    while (usedFilenames.has(filename.toLocaleLowerCase())) {
      filename = `${base}-${suffix}.xlsx`;
      suffix += 1;
    }

    usedFilenames.add(filename.toLocaleLowerCase());
    return path.join(outputDirectory, filename);
  });
}

function isMissingPathError(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    error.code === "ENOENT"
  );
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

function worksheetToTable(
  filePath: string,
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
      file: path.basename(filePath),
      firstDataRow: headerRowIndex + 2,
      sheet: sheetName,
    },
  };
}

function excelTableToTable(
  filePath: string,
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
      file: path.basename(filePath),
      firstDataRow: firstDataRowIndex + 1,
      sheet: definition.sheet,
    },
  };
}

function isVisibleSheet(workbook: XLSX.WorkBook, sheetName: string): boolean {
  const metadata = workbook.Workbook?.Sheets?.find(
    (sheet) => sheet.name === sheetName,
  );
  return (metadata?.Hidden ?? 0) === 0;
}

export async function readWorkbookTables(
  filePath: string,
  options: ReadWorkbookOptions = {},
): Promise<Table[]> {
  const absolutePath = path.resolve(filePath);
  let workbook: XLSX.WorkBook;

  try {
    workbook = XLSX.read(await readFile(absolutePath), {
      cellDates: true,
      dense: false,
      type: "buffer",
    });
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${absolutePath}`,
      {
        cause: error,
        details: { filePath: absolutePath },
      },
    );
  }

  const selectedSheets = options.sheets
    ? new Set(options.sheets.map((sheet) => sheet.toLocaleLowerCase()))
    : undefined;
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
      absolutePath,
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

export async function readWorksheetRecords(
  filePath: string,
  options: ReadWorksheetRecordsOptions,
): Promise<WorksheetRecords> {
  const absolutePath = path.resolve(filePath);
  let workbook: XLSX.WorkBook;

  try {
    workbook = XLSX.read(await readFile(absolutePath), {
      cellDates: true,
      cellText: true,
      dense: false,
      type: "buffer",
    });
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${absolutePath}`,
      {
        cause: error,
        details: { filePath: absolutePath },
      },
    );
  }

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

export async function readWorkbookExcelTables(
  filePath: string,
  options: ReadWorkbookExcelTablesOptions = {},
): Promise<WorkbookExcelTable[]> {
  const absolutePath = path.resolve(filePath);
  let workbook: XLSX.WorkBook;
  let definitions: ExcelTableDefinition[];

  try {
    const workbookBytes = await readFile(absolutePath);
    workbook = XLSX.read(workbookBytes, {
      cellDates: true,
      dense: false,
      type: "buffer",
    });
    definitions = await readExcelTableDefinitions(workbookBytes);
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${absolutePath}`,
      {
        cause: error,
        details: { filePath: absolutePath },
      },
    );
  }

  const selectedSheets = options.sheets
    ? new Set(options.sheets.map((sheet) => sheet.toLocaleLowerCase()))
    : undefined;
  const selectedTables = options.tables
    ? new Set(options.tables.map((table) => table.toLocaleLowerCase()))
    : undefined;
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
    const table = excelTableToTable(absolutePath, definition, worksheet);
    if (table) {
      tables.push(table);
    }
  }

  return tables;
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
  filePath: string,
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
      file: path.basename(filePath),
      firstDataRow: range.s.r + 2,
      sheet: sheetName,
    },
  };
}

export async function readWorkbookNamedRanges(
  filePath: string,
  options: ReadWorkbookNamedRangesOptions = {},
): Promise<WorkbookNamedRange[]> {
  const absolutePath = path.resolve(filePath);
  let workbook: XLSX.WorkBook;

  try {
    workbook = XLSX.read(await readFile(absolutePath), {
      cellDates: true,
      cellText: true,
      dense: false,
      type: "buffer",
    });
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${absolutePath}`,
      {
        cause: error,
        details: { filePath: absolutePath },
      },
    );
  }

  const selectedSheets = options.sheets
    ? new Set(options.sheets.map((sheet) => sheet.toLocaleLowerCase()))
    : undefined;
  const selectedNames = options.names
    ? new Set(options.names.map((name) => name.toLocaleLowerCase()))
    : undefined;
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
      absolutePath,
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

export async function writeTable(
  outputPath: string,
  table: Table,
  options: WriteTableOptions = {},
): Promise<string> {
  if (table.columns.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_COLUMNS,
      "Cannot write a table with no columns.",
    );
  }

  const absoluteOutput = await ensureParentDirectory(outputPath);
  await ensureOutputAvailable(absoluteOutput, {
    overwrite: options.overwrite,
  });

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
  workbook.Props = {
    Author: "ConsultChimps",
    CreatedDate: new Date(0),
    LastAuthor: "ConsultChimps",
    ModifiedDate: new Date(0),
  };
  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    options.sheetName ?? "Consolidated",
  );
  await writeFile(
    absoluteOutput,
    XLSX.write(workbook, {
      bookType: "xlsx",
      compression: true,
      type: "buffer",
    }),
  );
  return absoluteOutput;
}

const CONSOLIDATE_OPERATION = "sheets.consolidate";
const SPLIT_OPERATION = "sheets.split-by-column";
const WORKBOOK_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

interface ResolvedConsolidate {
  absoluteInputs: string[];
  absoluteOutput: string;
}

function resolveConsolidateWorkbooks(
  options: ConsolidateWorkbooksOptions,
): ResolvedConsolidate {
  if (options.inputs.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_INPUTS,
      "At least one workbook is required.",
    );
  }

  const absoluteInputs = options.inputs.map((inputPath) =>
    path.resolve(inputPath),
  );
  const absoluteOutput = path.resolve(options.output);
  refuseInputOverwrite(absoluteOutput, absoluteInputs);
  return { absoluteInputs, absoluteOutput };
}

export async function planConsolidateWorkbooks(
  options: ConsolidateWorkbooksOptions,
): Promise<OperationPlan<ConsolidateWorkbooksPlanMetric>> {
  const { absoluteInputs, absoluteOutput } =
    resolveConsolidateWorkbooks(options);

  for (const absoluteInput of absoluteInputs) {
    if (!(await pathExists(absoluteInput))) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_INPUT_NOT_FOUND,
        `Workbook not found: ${absoluteInput}`,
        { details: { inputPath: absoluteInput } },
      );
    }
  }

  const exists = await pathExists(absoluteOutput);
  const warnings =
    exists && options.overwrite !== true
      ? [
          "The planned output workbook already exists; executing without overwrite will fail.",
        ]
      : [];

  return {
    operation: CONSOLIDATE_OPERATION,
    inputs: absoluteInputs,
    outputs: [
      {
        kind: "file",
        mediaType: WORKBOOK_MEDIA_TYPE,
        path: absoluteOutput,
        exists,
      },
    ],
    warnings,
    metrics: {
      inputFiles: absoluteInputs.length,
      outputFiles: 1,
    },
  };
}

export async function consolidateWorkbooks(
  options: ConsolidateWorkbooksOptions,
): Promise<OperationResult<ConsolidateWorkbooksMetric>> {
  throwIfAborted(options.signal, CONSOLIDATE_OPERATION);
  const { absoluteInputs, absoluteOutput } =
    resolveConsolidateWorkbooks(options);

  const tables: Table[] = [];
  for (const [index, absoluteInput] of absoluteInputs.entries()) {
    throwIfAborted(options.signal, CONSOLIDATE_OPERATION);
    tables.push(...(await readWorkbookTables(absoluteInput, options)));
    options.onProgress?.({
      operation: CONSOLIDATE_OPERATION,
      stage: "reading-workbooks",
      completed: index + 1,
      total: absoluteInputs.length,
      detail: path.basename(absoluteInput),
    });
  }

  if (tables.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_TABLES,
      "No visible, non-empty worksheets were found in the input workbooks.",
    );
  }

  const table = unionTables(tables, {
    addSourceColumns: options.addSourceColumns,
  });
  throwIfAborted(options.signal, CONSOLIDATE_OPERATION);
  const output = await writeTable(absoluteOutput, table, {
    overwrite: options.overwrite,
    sheetName: options.outputSheetName,
  });
  options.onProgress?.({
    operation: CONSOLIDATE_OPERATION,
    stage: "writing-output",
    completed: 1,
    total: 1,
    detail: path.basename(output),
  });

  return {
    operation: CONSOLIDATE_OPERATION,
    artifacts: [
      {
        kind: "file",
        mediaType: WORKBOOK_MEDIA_TYPE,
        path: output,
      },
    ],
    warnings: [],
    metrics: {
      inputFiles: absoluteInputs.length,
      inputTables: tables.length,
      outputColumns: table.columns.length,
      outputRows: table.rows.length,
    },
  };
}

export async function mergeWorkbooks(
  inputPaths: string[],
  outputPath: string,
  options: MergeWorkbooksOptions = {},
): Promise<OperationResult<MergeWorkbooksMetric>> {
  if (inputPaths.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_INPUTS,
      "At least one workbook is required.",
    );
  }
  const absoluteInputs = inputPaths.map((inputPath) => path.resolve(inputPath));
  const absoluteOutput = path.resolve(outputPath);
  refuseInputOverwrite(absoluteOutput, absoluteInputs);
  await ensureOutputAvailable(absoluteOutput, { overwrite: options.overwrite });
  const outputWorkbook = XLSX.utils.book_new();
  const usedNames = new Set<string>(
    options.includeSheetIndex === false ? [] : ["sheet index"],
  );
  const indexRows: string[][] = [
    [
      "Source file",
      "Original worksheet",
      "Final worksheet",
      "Source visibility",
    ],
  ];
  let sheetCount = 0;
  let hiddenCount = 0;

  for (const inputPath of absoluteInputs) {
    let inputWorkbook: XLSX.WorkBook;
    try {
      inputWorkbook = XLSX.read(await readFile(inputPath), {
        cellDates: true,
        cellStyles: true,
        dense: false,
        type: "buffer",
      });
    } catch (error) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_READ_FAILED,
        `Could not read workbook: ${inputPath}`,
        { cause: error, details: { filePath: inputPath } },
      );
    }
    for (const originalName of inputWorkbook.SheetNames) {
      const worksheet = inputWorkbook.Sheets[originalName];
      if (!worksheet) {
        continue;
      }
      const baseName = originalName.slice(0, 31) || "Sheet";
      let finalName = baseName;
      let suffix = 2;
      while (usedNames.has(finalName.toLocaleLowerCase())) {
        const suffixText = ` (${suffix})`;
        finalName = `${baseName.slice(0, 31 - suffixText.length)}${suffixText}`;
        suffix += 1;
      }
      usedNames.add(finalName.toLocaleLowerCase());
      const visibility =
        inputWorkbook.Workbook?.Sheets?.find(
          (sheet) => sheet.name === originalName,
        )?.Hidden ?? 0;
      if (visibility !== 0) {
        hiddenCount += 1;
      }
      sheetCount += 1;
      indexRows.push([
        path.basename(inputPath),
        originalName,
        finalName,
        visibility === 2
          ? "Very hidden"
          : visibility === 1
            ? "Hidden"
            : "Visible",
      ]);
      XLSX.utils.book_append_sheet(outputWorkbook, worksheet, finalName);
      outputWorkbook.Workbook ??= {};
      outputWorkbook.Workbook.Sheets ??= [];
      outputWorkbook.Workbook.Sheets.push({
        name: finalName,
        Hidden: visibility,
      });
    }
  }
  if (sheetCount === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_SHEETS,
      "No worksheets were found in the input workbooks.",
    );
  }
  if (options.includeSheetIndex !== false) {
    XLSX.utils.book_append_sheet(
      outputWorkbook,
      XLSX.utils.aoa_to_sheet(indexRows),
      "Sheet Index",
    );
  }
  const outputBytes = XLSX.write(outputWorkbook, {
    bookType: "xlsx",
    cellStyles: true,
    compression: true,
    type: "buffer",
  });
  await ensureParentDirectory(absoluteOutput);
  await writeFile(absoluteOutput, outputBytes);
  const warnings = [];
  if (hiddenCount > 0) {
    warnings.push(
      options.includeSheetIndex === false
        ? `${hiddenCount} source worksheet${hiddenCount === 1 ? " was" : "s were"} hidden in the merged workbook.`
        : `${hiddenCount} source worksheet${hiddenCount === 1 ? " was" : "s were"} hidden; see the visible "Sheet Index" worksheet.`,
    );
  }
  return {
    operation: "sheets.merge",
    artifacts: [
      {
        kind: "file",
        mediaType: WORKBOOK_MEDIA_TYPE,
        path: absoluteOutput,
      },
    ],
    warnings,
    metrics: {
      inputFiles: inputPaths.length,
      outputSheets: sheetCount,
      hiddenSheets: hiddenCount,
    },
  };
}

interface ResolvedSplit {
  absoluteInput: string;
  absoluteOutputDirectory: string;
  existingOutputs: Set<string>;
  grouped: ReturnType<typeof groupTableByColumn>;
  outputPaths: string[];
  preservedTableDefinition: ExcelTableDefinition | undefined;
  preservedWorkbookBytes: Buffer | undefined;
  table: Table;
}

function skippedRowsWarning(
  grouped: ReturnType<typeof groupTableByColumn>,
): string {
  return `Skipped ${grouped.skippedRows} row${
    grouped.skippedRows === 1 ? "" : "s"
  } with blank values in "${grouped.column}".`;
}

async function resolveSplitWorkbookByColumn(
  options: SplitWorkbookByColumnOptions,
): Promise<ResolvedSplit> {
  const absoluteInput = path.resolve(options.input);
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
  // Table splits preserve the complete workbook unless explicitly disabled.
  const preserveWorkbook =
    options.preserveWorkbook ?? options.table !== undefined;
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

  const availableExcelTables = options.table
    ? await readWorkbookExcelTables(absoluteInput, {
        includeHiddenSheets: options.includeHiddenSheets,
        sheets: options.sheet ? [options.sheet] : undefined,
      })
    : [];
  const availableNamedRanges = options.range
    ? await readWorkbookNamedRanges(absoluteInput, {
        includeHiddenSheets: options.includeHiddenSheets,
        sheets: options.sheet ? [options.sheet] : undefined,
      })
    : [];
  const tables: Table[] = options.table
    ? availableExcelTables.filter(
        (table) =>
          table.excelTableName.toLocaleLowerCase() ===
          options.table?.toLocaleLowerCase(),
      )
    : options.range
      ? availableNamedRanges.filter(
          (namedRange) =>
            namedRange.rangeName.toLocaleLowerCase() ===
            options.range?.toLocaleLowerCase(),
        )
      : await readWorkbookTables(absoluteInput, {
          headerRow: options.headerRow,
          includeHiddenSheets: options.includeHiddenSheets,
          sheets: options.sheet ? [options.sheet] : undefined,
        });

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
          inputPath: absoluteInput,
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
          inputPath: absoluteInput,
        },
      },
    );
  }

  const table = tables[0];
  if (!table) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_NO_TABLE,
      "No worksheet table was available to split.",
      { details: { inputPath: absoluteInput } },
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
          inputPath: absoluteInput,
        },
      },
    );
  }

  const preservedWorkbookBytes = preserveWorkbook
    ? await readFile(absoluteInput)
    : undefined;
  const preservedTableDefinition = preservedWorkbookBytes
    ? (await readExcelTableDefinitions(preservedWorkbookBytes)).find(
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
          inputPath: absoluteInput,
          sheet: table.source?.sheet,
          table: options.table,
        },
      },
    );
  }

  const absoluteOutputDirectory = path.resolve(options.outputDirectory);
  const inputBaseName = path.parse(absoluteInput).name;
  const filenamePrefix = safeFilenameSegment(
    options.filenamePrefix ?? inputBaseName,
    "split",
  );
  const outputPaths = splitOutputPaths(
    absoluteOutputDirectory,
    filenamePrefix,
    grouped.groups.map((group) => group.value),
  );

  outputPaths.forEach((outputPath) =>
    refuseInputOverwrite(outputPath, [absoluteInput]),
  );
  const existingOutputs = new Set<string>();
  await Promise.all(
    outputPaths.map(async (outputPath) => {
      try {
        const outputStat = await stat(outputPath);
        if (!outputStat.isFile()) {
          throw new ConsultChimpsError(
            XLSX_ERRORS.XLSX_SPLIT_OUTPUT_NOT_FILE,
            `Output path exists but is not a file: ${outputPath}`,
            { details: { outputPath } },
          );
        }
        existingOutputs.add(outputPath);
      } catch (error) {
        if (!isMissingPathError(error)) {
          throw error;
        }
      }
    }),
  );

  return {
    absoluteInput,
    absoluteOutputDirectory,
    existingOutputs,
    grouped,
    outputPaths,
    preservedTableDefinition,
    preservedWorkbookBytes,
    table,
  };
}

export async function planSplitWorkbookByColumn(
  options: SplitWorkbookByColumnOptions,
): Promise<OperationPlan<SplitWorkbookByColumnPlanMetric>> {
  const resolved = await resolveSplitWorkbookByColumn(options);
  const outputs: PlannedOutput[] = resolved.outputPaths.map((outputPath) => ({
    kind: "file",
    mediaType: WORKBOOK_MEDIA_TYPE,
    path: outputPath,
    exists: resolved.existingOutputs.has(outputPath),
  }));

  const warnings: string[] = [];
  if (resolved.grouped.skippedRows > 0) {
    warnings.push(skippedRowsWarning(resolved.grouped));
  }
  const collisions = resolved.existingOutputs.size;
  if (collisions > 0 && options.overwrite !== true) {
    warnings.push(
      `${collisions} planned output file${
        collisions === 1 ? " already exists" : "s already exist"
      }; executing without overwrite will fail.`,
    );
  }

  return {
    operation: SPLIT_OPERATION,
    inputs: [resolved.absoluteInput],
    outputs,
    warnings,
    metrics: {
      groups: resolved.grouped.groups.length,
      inputFiles: 1,
      inputRows: resolved.table.rows.length,
      outputFiles: resolved.outputPaths.length,
      skippedRows: resolved.grouped.skippedRows,
    },
  };
}

export async function splitWorkbookByColumn(
  options: SplitWorkbookByColumnOptions,
): Promise<OperationResult<SplitWorkbookByColumnMetric>> {
  throwIfAborted(options.signal, SPLIT_OPERATION);
  const {
    absoluteOutputDirectory,
    existingOutputs,
    grouped,
    outputPaths,
    preservedTableDefinition,
    preservedWorkbookBytes,
    table,
  } = await resolveSplitWorkbookByColumn(options);

  await Promise.all(
    outputPaths.map((outputPath) =>
      ensureOutputAvailable(outputPath, { overwrite: options.overwrite }),
    ),
  );
  await ensureDirectory(absoluteOutputDirectory);

  const transactionDirectory = await mkdtemp(
    path.join(absoluteOutputDirectory, ".consultchimps-split-"),
  );
  const stagedOutputs: string[] = [];
  const committedOutputs: string[] = [];
  const backups = new Map<string, string>();

  try {
    for (const [index, group] of grouped.groups.entries()) {
      throwIfAborted(options.signal, SPLIT_OPERATION);
      const stagedOutput = path.join(
        transactionDirectory,
        `output-${String(index + 1).padStart(6, "0")}.xlsx`,
      );
      if (preservedWorkbookBytes && preservedTableDefinition) {
        await writeFile(
          stagedOutput,
          await preserveWorkbookWithFilteredExcelTable(preservedWorkbookBytes, {
            definition: preservedTableDefinition,
            sourceRows: group.table.sourceRows ?? [],
          }),
        );
        stagedOutputs.push(stagedOutput);
      } else {
        stagedOutputs.push(
          await writeTable(stagedOutput, group.table, {
            sheetName: table.source?.sheet ?? "Split",
          }),
        );
      }
      options.onProgress?.({
        operation: SPLIT_OPERATION,
        stage: "staging-workbooks",
        completed: index + 1,
        total: grouped.groups.length,
        detail: path.basename(outputPaths[index] ?? stagedOutput),
      });
    }

    for (const [index, outputPath] of outputPaths.entries()) {
      const stagedOutput = stagedOutputs[index];
      if (!stagedOutput) {
        continue;
      }

      if (existingOutputs.has(outputPath)) {
        const backupPath = path.join(
          transactionDirectory,
          `backup-${String(index + 1).padStart(6, "0")}.xlsx`,
        );
        await rename(outputPath, backupPath);
        backups.set(outputPath, backupPath);
      }

      await rename(stagedOutput, outputPath);
      committedOutputs.push(outputPath);
      options.onProgress?.({
        operation: SPLIT_OPERATION,
        stage: "committing-outputs",
        completed: index + 1,
        total: outputPaths.length,
        detail: path.basename(outputPath),
      });
    }
  } catch (error) {
    const rollbackErrors: unknown[] = [];

    for (const outputPath of [...committedOutputs].reverse()) {
      try {
        await rm(outputPath, { force: true });
        const backupPath = backups.get(outputPath);
        if (backupPath) {
          await rename(backupPath, outputPath);
          backups.delete(outputPath);
        }
      } catch (rollbackError) {
        rollbackErrors.push(rollbackError);
      }
    }

    for (const [outputPath, backupPath] of backups) {
      try {
        await rename(backupPath, outputPath);
      } catch (rollbackError) {
        rollbackErrors.push(rollbackError);
      }
    }

    if (rollbackErrors.length > 0) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_SPLIT_ROLLBACK_FAILED,
        "The split failed and one or more output files could not be restored.",
        {
          cause: error,
          details: {
            outputPaths,
            rollbackErrors: rollbackErrors.map((rollbackError) =>
              rollbackError instanceof Error
                ? rollbackError.message
                : String(rollbackError),
            ),
          },
        },
      );
    }

    throw error;
  } finally {
    await rm(transactionDirectory, { force: true, recursive: true }).catch(
      () => undefined,
    );
  }

  const warnings = grouped.skippedRows > 0 ? [skippedRowsWarning(grouped)] : [];

  return {
    operation: SPLIT_OPERATION,
    artifacts: outputPaths.map((output) => ({
      kind: "file",
      mediaType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      path: output,
    })),
    warnings,
    metrics: {
      groups: grouped.groups.length,
      inputFiles: 1,
      inputRows: table.rows.length,
      outputFiles: outputPaths.length,
      outputRows: grouped.groups.reduce(
        (total, group) => total + group.table.rows.length,
        0,
      ),
      skippedRows: grouped.skippedRows,
    },
  };
}
