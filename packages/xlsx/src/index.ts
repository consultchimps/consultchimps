import {
  mkdtemp,
  readFile,
  rename,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import path from "node:path";

import { ConsultChimpsError, type OperationResult } from "@consultchimps/core";
import {
  ensureDirectory,
  ensureOutputAvailable,
  ensureParentDirectory,
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
import { preserveWorkbookWithFilteredExcelTable } from "./preserve-table-split.js";

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
  worksheet: string;
}

export interface WorksheetRecords {
  columns: string[];
  rows: Array<Record<string, string>>;
  skippedEmptyRows: number;
  sourceRows: number[];
  worksheet: string;
}

export interface ConsolidateWorkbooksOptions extends ReadWorkbookOptions {
  addSourceColumns?: boolean | undefined;
  outputSheetName?: string | undefined;
  overwrite?: boolean | undefined;
}

export interface SplitWorkbookByColumnOptions {
  column: string;
  filenamePrefix?: string | undefined;
  headerRow?: number | undefined;
  includeBlank?: boolean | undefined;
  includeHiddenSheets?: boolean | undefined;
  overwrite?: boolean | undefined;
  preserveWorkbook?: boolean | undefined;
  sheet?: string | undefined;
  table?: string | undefined;
}

export interface WorkbookExcelTable extends Table {
  excelTableName: string;
  excelTableRange: string;
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
        "XLSX_INVALID_HEADER_ROW",
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
      "XLSX_INVALID_EXCEL_TABLE",
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
      "XLSX_INVALID_EXCEL_TABLE",
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
      "XLSX_READ_FAILED",
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
      "XLSX_READ_FAILED",
      `Could not read workbook: ${absolutePath}`,
      {
        cause: error,
        details: { filePath: absolutePath },
      },
    );
  }

  const requestedWorksheet = options.worksheet.trim();
  const worksheetName = workbook.SheetNames.find(
    (candidate) =>
      candidate.toLocaleLowerCase() === requestedWorksheet.toLocaleLowerCase(),
  );
  if (!worksheetName) {
    throw new ConsultChimpsError(
      "XLSX_WORKSHEET_NOT_FOUND",
      `Worksheet "${options.worksheet}" was not found in the workbook.`,
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
      "XLSX_INVALID_HEADER_ROW",
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
      "XLSX_INVALID_HEADER_ROW",
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
        "XLSX_EMPTY_HEADER",
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
      "XLSX_DUPLICATE_HEADER",
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
      "XLSX_READ_FAILED",
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

export async function writeTable(
  outputPath: string,
  table: Table,
  options: WriteTableOptions = {},
): Promise<string> {
  if (table.columns.length === 0) {
    throw new ConsultChimpsError(
      "XLSX_NO_COLUMNS",
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

export async function consolidateWorkbooks(
  inputPaths: string[],
  outputPath: string,
  options: ConsolidateWorkbooksOptions = {},
): Promise<OperationResult> {
  if (inputPaths.length === 0) {
    throw new ConsultChimpsError(
      "XLSX_NO_INPUTS",
      "At least one workbook is required.",
    );
  }

  refuseInputOverwrite(outputPath, inputPaths);
  const tables = (
    await Promise.all(
      inputPaths.map((inputPath) => readWorkbookTables(inputPath, options)),
    )
  ).flat();

  if (tables.length === 0) {
    throw new ConsultChimpsError(
      "XLSX_NO_TABLES",
      "No visible, non-empty worksheets were found in the input workbooks.",
    );
  }

  const table = unionTables(tables, {
    addSourceColumns: options.addSourceColumns,
  });
  const output = await writeTable(outputPath, table, {
    overwrite: options.overwrite,
    sheetName: options.outputSheetName,
  });

  return {
    operation: "sheets.consolidate",
    artifacts: [
      {
        kind: "file",
        mediaType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        path: output,
      },
    ],
    warnings: [],
    metrics: {
      inputFiles: inputPaths.length,
      inputTables: tables.length,
      outputColumns: table.columns.length,
      outputRows: table.rows.length,
    },
  };
}

export async function splitWorkbookByColumn(
  inputPath: string,
  outputDirectory: string,
  options: SplitWorkbookByColumnOptions,
): Promise<OperationResult> {
  const absoluteInput = path.resolve(inputPath);
  if (options.table && options.headerRow !== undefined) {
    throw new ConsultChimpsError(
      "XLSX_SPLIT_TABLE_HEADER_ROW",
      "The headerRow option cannot be used with an Excel Table; the table defines its own headers.",
      {
        details: {
          headerRow: options.headerRow,
          table: options.table,
        },
      },
    );
  }
  if (options.preserveWorkbook && !options.table) {
    throw new ConsultChimpsError(
      "XLSX_SPLIT_PRESERVE_REQUIRES_TABLE",
      "The preserveWorkbook option requires a named Excel Table.",
      {
        details: {
          preserveWorkbook: options.preserveWorkbook,
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
  const tables = options.table
    ? availableExcelTables.filter(
        (table) =>
          table.excelTableName.toLocaleLowerCase() ===
          options.table?.toLocaleLowerCase(),
      )
    : await readWorkbookTables(absoluteInput, {
        headerRow: options.headerRow,
        includeHiddenSheets: options.includeHiddenSheets,
        sheets: options.sheet ? [options.sheet] : undefined,
      });

  if (tables.length === 0) {
    const selectedSource = options.table
      ? `Excel Table "${options.table}"`
      : options.sheet
        ? `Worksheet "${options.sheet}"`
        : undefined;
    throw new ConsultChimpsError(
      "XLSX_SPLIT_NO_TABLE",
      selectedSource
        ? `${selectedSource} was not found or has no data rows.`
        : "No visible, non-empty worksheet was found in the input workbook.",
      {
        details: {
          availableTables: availableExcelTables.map((table) => ({
            name: table.excelTableName,
            sheet: table.source?.sheet,
          })),
          inputPath: absoluteInput,
          sheet: options.sheet,
          table: options.table,
        },
      },
    );
  }

  if (tables.length > 1) {
    throw new ConsultChimpsError(
      "XLSX_SPLIT_MULTIPLE_TABLES",
      options.table
        ? `Excel Table "${options.table}" was found on multiple worksheets; choose one with the sheet option.`
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
      "XLSX_SPLIT_NO_TABLE",
      "No worksheet table was available to split.",
      { details: { inputPath: absoluteInput } },
    );
  }

  const grouped = groupTableByColumn(table, options.column, {
    includeBlank: options.includeBlank,
  });
  if (grouped.groups.length === 0) {
    throw new ConsultChimpsError(
      "XLSX_SPLIT_NO_GROUPS",
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

  const preservedWorkbookBytes = options.preserveWorkbook
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
  if (options.preserveWorkbook && !preservedTableDefinition) {
    throw new ConsultChimpsError(
      "XLSX_SPLIT_PRESERVE_TABLE_NOT_FOUND",
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

  const absoluteOutputDirectory = path.resolve(outputDirectory);
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
            "XLSX_SPLIT_OUTPUT_NOT_FILE",
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

      await ensureOutputAvailable(outputPath, {
        overwrite: options.overwrite,
      });
    }),
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
        "XLSX_SPLIT_ROLLBACK_FAILED",
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

  const warnings =
    grouped.skippedRows > 0
      ? [
          `Skipped ${grouped.skippedRows} row${
            grouped.skippedRows === 1 ? "" : "s"
          } with blank values in "${grouped.column}".`,
        ]
      : [];

  return {
    operation: "sheets.split-by-column",
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
