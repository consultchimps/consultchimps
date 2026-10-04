/**
 * L3: one worksheet's cells held in memory, read through the streaming reader
 * (ADR 0006), for the readers that return whole tables: worksheet tables,
 * Excel Tables and named ranges. Their results hold every row anyway, so the
 * grid costs no more than what they return.
 */
import type { CellValue } from "@consultchimps/tabular";

import { CellError } from "../package/cell-error.js";
import { bytesSource } from "../package/index.js";
import {
  StreamedWorkbook,
  type CellRectangle,
  type StreamedCell,
  type StreamedSheet,
  type StreamedValue,
  type StreamedWorkbookContext,
} from "./consolidate/reader.js";

/** One worksheet's cells, in the engine's zero-based numbering. */
export interface SheetGrid {
  /** The used range, or undefined when the worksheet holds nothing. */
  readonly range: CellRectangle | undefined;
  readonly merges: readonly CellRectangle[];
  /**
   * The value a cell holds, or null for an empty cell. An error cell holds its
   * text, `#DIV/0!`, which is what the worksheet shows.
   */
  value(row: number, column: number): CellValue;
}

/** A table value for a streamed one: an error cell becomes its text. */
export function tableValue(value: StreamedValue | undefined): CellValue {
  if (value === undefined) return null;
  return value instanceof CellError ? value.text : value;
}

/** Open workbook bytes for streaming, refusing an unreadable package. */
export function openWorkbookBytes(
  bytes: Uint8Array,
  context: StreamedWorkbookContext,
): Promise<StreamedWorkbook> {
  return StreamedWorkbook.open(bytesSource(context.source, bytes), context);
}

/**
 * Read one worksheet whole. Cells outside the used range are kept, because an
 * Excel Table or a named range declares its own rectangle.
 */
export async function readSheetGrid(
  workbook: StreamedWorkbook,
  sheet: StreamedSheet,
): Promise<SheetGrid> {
  let rows = new Map<number, Map<number, StreamedValue>>();
  const read = await workbook.readWorksheet(
    sheet,
    {
      begin() {
        rows = new Map();
      },
      row(row: number, cells: readonly StreamedCell[]) {
        rows.set(
          row,
          new Map(cells.map((cell) => [cell.column, cell.value] as const)),
        );
      },
    },
    { clip: false },
  );
  return {
    range: read.range,
    merges: read.merges,
    value: (row, column) => tableValue(rows.get(row)?.get(column)),
  };
}

/**
 * The grids of the worksheets a read asks for, each read once however many
 * Excel Tables or names sit on it.
 */
export class SheetGrids {
  readonly #workbook: StreamedWorkbook;
  readonly #grids = new Map<string, SheetGrid>();

  constructor(workbook: StreamedWorkbook) {
    this.#workbook = workbook;
  }

  /** The grid of the worksheet named exactly `name`, or undefined when none is. */
  async named(name: string): Promise<SheetGrid | undefined> {
    const cached = this.#grids.get(name);
    if (cached !== undefined) return cached;
    const sheet = this.#workbook.sheets.find(
      (candidate) => candidate.name === name,
    );
    if (sheet === undefined) return undefined;
    const grid = await readSheetGrid(this.#workbook, sheet);
    this.#grids.set(name, grid);
    return grid;
  }
}
