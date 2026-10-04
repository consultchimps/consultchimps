/**
 * Read helpers for assertions, through the package's own streaming reader.
 * Values are what a table holds: dates as ISO text, error cells as their text.
 */
import type { CellValue } from "@consultchimps/tabular";

import {
  openWorkbookBytes,
  readSheetGrid,
} from "../../src/operations/sheet-grid.js";

async function open(bytes: Uint8Array) {
  return openWorkbookBytes(bytes, {
    source: "book.xlsx",
    file: "book.xlsx",
    details: { source: "book.xlsx" },
  });
}

/** Every worksheet's name, hidden ones included, in workbook order. */
export async function sheetNamesOf(bytes: Uint8Array): Promise<string[]> {
  return (await open(bytes)).sheets.map((sheet) => sheet.name);
}

export interface SheetRowsOptions {
  /** Return the text each cell shows instead of its value. */
  readonly text?: boolean;
}

/**
 * One worksheet's used range as rows, from its first row and column, with
 * null (or "" for text) in every empty cell.
 */
export async function sheetRows(
  bytes: Uint8Array,
  sheetName: string,
  options: SheetRowsOptions = {},
): Promise<CellValue[][]> {
  const workbook = await open(bytes);
  const sheet = workbook.sheets.find(
    (candidate) => candidate.name === sheetName,
  );
  if (sheet === undefined) throw new Error(`No worksheet "${sheetName}".`);
  const text = options.text === true;
  const grid = await readSheetGrid(workbook, sheet, { text });
  const range = grid.range;
  if (range === undefined) return [];
  const rows: CellValue[][] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const cells: CellValue[] = [];
    for (let column = range.startColumn; column <= range.endColumn; column++) {
      cells.push(text ? grid.text(row, column) : grid.value(row, column));
    }
    rows.push(cells);
  }
  return rows;
}

/**
 * One worksheet as records keyed by its first row, skipping blank rows and
 * leaving out empty cells, unless `keepEmpty` keeps them as null.
 */
export async function sheetRecords(
  bytes: Uint8Array,
  sheetName: string,
  options: SheetRowsOptions & { readonly keepEmpty?: boolean } = {},
): Promise<Record<string, CellValue>[]> {
  const [header = [], ...body] = await sheetRows(bytes, sheetName, options);
  const empty = options.text === true ? "" : null;
  const records: Record<string, CellValue>[] = [];
  for (const row of body) {
    if (row.every((value) => value === empty)) continue;
    const record: Record<string, CellValue> = {};
    header.forEach((key, index) => {
      const value = row[index] ?? null;
      if (value === empty && options.keepEmpty !== true) return;
      record[String(key)] = value === empty ? null : value;
    });
    records.push(record);
  }
  return records;
}
