import { readFile, writeFile } from "node:fs/promises";
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

export interface SplitWorkbookByColumnOptions {
  column: string;
  filenamePrefix?: string | undefined;
  headerRow?: number | undefined;
  includeBlank?: boolean | undefined;
  includeHiddenSheets?: boolean | undefined;
  overwrite?: boolean | undefined;
  sheet?: string | undefined;
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

export async function splitWorkbookByColumn(
  inputPath: string,
  outputDirectory: string,
  options: SplitWorkbookByColumnOptions,
): Promise<OperationResult> {
  const absoluteInput = path.resolve(inputPath);
  const tables = await readWorkbookTables(absoluteInput, {
    headerRow: options.headerRow,
    includeHiddenSheets: options.includeHiddenSheets,
    sheets: options.sheet ? [options.sheet] : undefined,
  });

  if (tables.length === 0) {
    throw new ConsultChimpsError(
      "XLSX_SPLIT_NO_TABLE",
      options.sheet
        ? `Worksheet "${options.sheet}" was not found or has no data rows.`
        : "No visible, non-empty worksheet was found in the input workbook.",
      {
        details: {
          inputPath: absoluteInput,
          sheet: options.sheet,
        },
      },
    );
  }

  if (tables.length > 1) {
    throw new ConsultChimpsError(
      "XLSX_SPLIT_MULTIPLE_TABLES",
      "The workbook contains multiple non-empty worksheets; choose one with the sheet option.",
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
  await Promise.all(
    outputPaths.map((outputPath) =>
      ensureOutputAvailable(outputPath, {
        overwrite: options.overwrite,
      }),
    ),
  );
  await ensureDirectory(absoluteOutputDirectory);

  const outputs: string[] = [];
  for (const [index, group] of grouped.groups.entries()) {
    const outputPath = outputPaths[index];
    if (!outputPath) {
      continue;
    }

    outputs.push(
      await writeTable(outputPath, group.table, {
        overwrite: options.overwrite,
        sheetName: table.source?.sheet ?? "Split",
      }),
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
    artifacts: outputs.map((output) => ({
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
      outputFiles: outputs.length,
      outputRows: grouped.groups.reduce(
        (total, group) => total + group.table.rows.length,
        0,
      ),
      skippedRows: grouped.skippedRows,
    },
  };
}
