/**
 * Read helpers for assertions, through the package's own streaming reader.
 * Values are what a table holds: dates as ISO text, error cells as their text.
 */
import type { CellValue } from "@consultchimps/tabular";
import JSZip from "jszip";

import { WorkbookModel } from "../../src/model/index.js";
import { decodeCell } from "../../src/model/references.js";

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

/** Every worksheet's visibility, by name. */
export async function sheetVisibilities(
  bytes: Uint8Array,
): Promise<Record<string, "visible" | "hidden" | "veryHidden">> {
  const model = await WorkbookModel.load(bytes);
  return Object.fromEntries(
    model.sheets.map((sheet) => [sheet.name, sheet.visibility]),
  );
}

/** The raw XML of the worksheet part a sheet name points at. */
export async function sheetXml(
  bytes: Uint8Array,
  sheetName: string,
): Promise<string> {
  const model = await WorkbookModel.load(bytes);
  const sheet = model.sheets.find((candidate) => candidate.name === sheetName);
  if (sheet === undefined) throw new Error(`No worksheet "${sheetName}".`);
  const entry = (await JSZip.loadAsync(bytes)).file(sheet.partPath);
  if (entry === null) throw new Error(`No part "${sheet.partPath}".`);
  return entry.async("text");
}

/** The value one cell holds, such as "B4", or null when it is empty. */
export async function sheetCell(
  bytes: Uint8Array,
  sheetName: string,
  reference: string,
): Promise<CellValue> {
  const workbook = await open(bytes);
  const sheet = workbook.sheets.find(
    (candidate) => candidate.name === sheetName,
  );
  if (sheet === undefined) throw new Error(`No worksheet "${sheetName}".`);
  const cell = decodeCell(reference);
  if (cell === undefined) throw new Error(`Not a cell: ${reference}`);
  const grid = await readSheetGrid(workbook, sheet);
  return grid.value(cell.row - 1, cell.column);
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
