import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { ConsultChimpsError, type OperationResult } from "@consultchimps/core";
import {
  ensureOutputAvailable,
  ensureParentDirectory,
  refuseInputOverwrite,
} from "@consultchimps/files";
import {
  type CellValue,
  type Table,
  type TableRow,
  unionTables,
  uniqueHeaders,
} from "@consultchimps/tabular";
import * as XLSX from "xlsx";

export interface ReadWorkbookOptions {
  headerRow?: number | undefined;
  includeHiddenSheets?: boolean | undefined;
  sheets?: string[] | undefined;
}

export interface ConsolidateWorkbooksOptions extends ReadWorkbookOptions {
  addSourceColumns?: boolean | undefined;
  outputSheetName?: string | undefined;
  overwrite?: boolean | undefined;
}

export interface WriteTableOptions {
  overwrite?: boolean | undefined;
  sheetName?: string | undefined;
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
  if (headerRowIndex === undefined || headerRowIndex > range.e.r) {
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
