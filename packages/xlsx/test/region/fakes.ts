/**
 * Hand-written in-memory implementations of the L1 model interfaces.
 *
 * The region layer only needs rows, cells, a used range, and a `deleteRows`
 * call it can observe, so the tests build worksheets from plain grids instead
 * of from OOXML. That keeps the L2 tests independent of the L0/L1 stream and
 * lets them assert exactly which rows a binding asked to delete and whether it
 * asked for renumbering.
 */

import { RowRelocation } from "../../src/model/references.js";
import type {
  CellFormula,
  CellModel,
  CellRange,
  CellRef,
  CellValue,
  ColumnIndex,
  DefinedNameEntry,
  DeleteRowsOptions,
  DeleteRowsReport,
  RelocateRowsOptions,
  RowModel,
  RowNumber,
  SheetInfo,
  WorkbookModel,
  WorkbookTableInfo,
  WorksheetModel,
} from "../../src/model/types.js";
import { parseCellRange } from "../../src/region/values.js";

export interface FakeCellSpec {
  readonly formula?: CellFormula | undefined;
  readonly styleIndex?: number | undefined;
  readonly text?: string | undefined;
  /** OOXML `t`; omitted means a number, as in the file format. */
  readonly type?: string | undefined;
}

/**
 * A grid entry. `null` and `undefined` mean the cell is absent; an empty
 * string means a present but blank cell.
 */
export type FakeCellInput =
  FakeCellSpec | boolean | null | number | string | undefined;

export interface FakeSheetSpec {
  readonly firstColumn?: number | undefined;
  readonly firstRow?: number | undefined;
  readonly grid: readonly (readonly FakeCellInput[])[];
  readonly name: string;
  readonly partPath?: string | undefined;
  readonly visibility?: "hidden" | "veryHidden" | "visible" | undefined;
}

/** One observed `deleteRows` call, for assertions. */
export interface DeleteRowsCall {
  readonly renumber: boolean;
  readonly rows: readonly RowNumber[];
}

/** One observed relocation, as `[sourceRow, destination]` pairs. */
export interface RelocateRowsCall {
  readonly moves: ReadonlyArray<readonly [RowNumber, RowNumber | null]>;
  readonly resizeTables: boolean;
}

interface FakeRow {
  attributes: Record<string, string>;
  cells: Map<ColumnIndex, CellModel>;
}

function isCellSpec(input: FakeCellInput): input is FakeCellSpec {
  return typeof input === "object" && input !== null;
}

function toCellSpec(input: FakeCellInput): FakeCellSpec | undefined {
  if (input === null || input === undefined) {
    return undefined;
  }
  if (isCellSpec(input)) {
    return input;
  }
  if (typeof input === "number") {
    return { text: String(input) };
  }
  if (typeof input === "boolean") {
    return { text: input ? "1" : "0", type: "b" };
  }
  return { text: input, type: "s" };
}

export class FakeWorksheetModel implements WorksheetModel {
  readonly deleteCalls: DeleteRowsCall[] = [];
  readonly relocateCalls: RelocateRowsCall[] = [];
  readonly info: SheetInfo;
  private rowsByNumber: Map<RowNumber, FakeRow>;

  constructor(spec: FakeSheetSpec) {
    this.info = {
      name: spec.name,
      partPath: spec.partPath ?? `xl/worksheets/${spec.name}.xml`,
      visibility: spec.visibility ?? "visible",
    };
    this.rowsByNumber = new Map();
    const firstRow = spec.firstRow ?? 1;
    const firstColumn = spec.firstColumn ?? 0;
    spec.grid.forEach((rowInputs, rowOffset) => {
      const row = firstRow + rowOffset;
      const cells = new Map<ColumnIndex, CellModel>();
      rowInputs.forEach((input, columnOffset) => {
        const cell = toCellSpec(input);
        if (!cell) {
          return;
        }
        const column = firstColumn + columnOffset;
        cells.set(column, {
          formula: cell.formula,
          ref: { column, row },
          styleIndex: cell.styleIndex,
          type: cell.type,
          value: cell.text ?? "",
        });
      });
      if (cells.size > 0) {
        this.rowsByNumber.set(row, { attributes: {}, cells });
      }
    });
  }

  get usedRange(): CellRange | undefined {
    let minRow: number | undefined;
    let maxRow: number | undefined;
    let minColumn: number | undefined;
    let maxColumn: number | undefined;
    for (const [row, entry] of this.rowsByNumber) {
      for (const column of entry.cells.keys()) {
        minRow = minRow === undefined ? row : Math.min(minRow, row);
        maxRow = maxRow === undefined ? row : Math.max(maxRow, row);
        minColumn =
          minColumn === undefined ? column : Math.min(minColumn, column);
        maxColumn =
          maxColumn === undefined ? column : Math.max(maxColumn, column);
      }
    }
    if (
      minRow === undefined ||
      maxRow === undefined ||
      minColumn === undefined ||
      maxColumn === undefined
    ) {
      return undefined;
    }
    return {
      end: { column: maxColumn, row: maxRow },
      start: { column: minColumn, row: minRow },
    };
  }

  rows(): readonly RowModel[] {
    return [...this.rowsByNumber.keys()]
      .sort((left, right) => left - right)
      .map((number) => this.rowModel(number));
  }

  row(number: RowNumber): RowModel | undefined {
    return this.rowsByNumber.has(number) ? this.rowModel(number) : undefined;
  }

  cellText(ref: CellRef): string | undefined {
    return this.rowsByNumber.get(ref.row)?.cells.get(ref.column)?.value;
  }

  /**
   * The typing the real model performs, minus number formats: the grid has no
   * styles, so a numeric cell is always a number rather than a date.
   */
  cellValue(ref: CellRef): CellValue {
    const cell = this.rowsByNumber.get(ref.row)?.cells.get(ref.column);
    const text = cell?.value;
    if (!cell || text === undefined) {
      return undefined;
    }
    switch (cell.type) {
      case "b":
        return text.trim() === "1" || text.trim().toLowerCase() === "true";
      case "d": {
        const parsed = new Date(text);
        return Number.isNaN(parsed.getTime()) ? text : parsed;
      }
      case undefined: {
        const trimmed = text.trim();
        if (trimmed === "") {
          return undefined;
        }
        const numeric = Number(trimmed);
        return Number.isFinite(numeric) ? numeric : trimmed;
      }
      default:
        return text;
    }
  }

  get lastRow(): RowNumber {
    return [...this.rowsByNumber.keys()].reduce(
      (last, row) => Math.max(last, row),
      0,
    );
  }

  deleteRows(
    rows: ReadonlySet<RowNumber>,
    options: DeleteRowsOptions,
  ): DeleteRowsReport {
    this.deleteCalls.push({
      renumber: options.renumber,
      rows: [...rows].sort((left, right) => left - right),
    });
    return this.applyRowRelocation(
      RowRelocation.compacting(
        rows,
        Math.max(this.lastRow, ...rows, 0),
        options.renumber,
      ),
    );
  }

  applyRowRelocation(
    relocation: RowRelocation,
    options: RelocateRowsOptions | undefined = undefined,
  ): DeleteRowsReport {
    const moves: Array<readonly [RowNumber, RowNumber | null]> = [];
    const rewritten = new Map<RowNumber, FakeRow>();
    let deletedRows = 0;

    for (const [number, entry] of [...this.rowsByNumber.entries()].sort(
      ([left], [right]) => left - right,
    )) {
      const destination = relocation.target(number);
      moves.push([number, destination]);
      if (destination === null) {
        deletedRows += 1;
        continue;
      }
      rewritten.set(destination, {
        attributes: entry.attributes,
        cells: new Map(
          [...entry.cells].map(([column, cell]) => [
            column,
            { ...cell, ref: { column, row: destination } },
          ]),
        ),
      });
    }

    this.relocateCalls.push({
      moves,
      resizeTables: options?.resizeTables !== false,
    });
    this.rowsByNumber = rewritten;
    return {
      adjusted: {
        calcChainEntries: 0,
        conditionalFormatting: 0,
        dataValidations: 0,
        formulaReferences: 0,
        hyperlinks: 0,
        mergedRanges: 0,
        tableRefs: 0,
      },
      deletedRows,
      retainedRows: rewritten.size,
    };
  }

  private rowModel(number: RowNumber): RowModel {
    const entry = this.rowsByNumber.get(number);
    return {
      attributes: entry?.attributes ?? {},
      cells: [...(entry?.cells.values() ?? [])].sort(
        (left, right) => left.ref.column - right.ref.column,
      ),
      number,
    };
  }
}

export interface FakeWorkbookSpec {
  readonly definedNames?: readonly DefinedNameEntry[] | undefined;
  readonly sheets: readonly FakeSheetSpec[];
  readonly tables?: readonly WorkbookTableInfo[] | undefined;
}

export class FakeWorkbookModel implements WorkbookModel {
  readonly sheets: readonly SheetInfo[];
  readonly worksheets: readonly FakeWorksheetModel[];
  private readonly names: readonly DefinedNameEntry[];
  private readonly tableInfos: readonly WorkbookTableInfo[];

  constructor(spec: FakeWorkbookSpec) {
    this.worksheets = spec.sheets.map((sheet) => new FakeWorksheetModel(sheet));
    this.sheets = this.worksheets.map((worksheet) => worksheet.info);
    this.names = spec.definedNames ?? [];
    this.tableInfos = spec.tables ?? [];
  }

  worksheet(name: string): WorksheetModel | undefined {
    return this.worksheets.find((sheet) => sheet.info.name === name);
  }

  /** Same lookup, typed so tests can read `deleteCalls`. */
  fakeWorksheet(name: string): FakeWorksheetModel {
    const worksheet = this.worksheets.find((sheet) => sheet.info.name === name);
    if (!worksheet) {
      throw new Error(`Test worksheet "${name}" was not defined.`);
    }
    return worksheet;
  }

  definedNames(): Promise<readonly DefinedNameEntry[]> {
    return Promise.resolve(this.names);
  }

  tables(): Promise<readonly WorkbookTableInfo[]> {
    return Promise.resolve(this.tableInfos);
  }

  save(): Promise<Uint8Array> {
    return Promise.resolve(new Uint8Array());
  }
}

export interface FakeTableSpec {
  readonly columnNames: readonly string[];
  readonly headerRow?: RowNumber | undefined;
  readonly name: string;
  /** A1 range covering header, body, and totals row, such as "A1:C6". */
  readonly ref: string;
  readonly sheetName: string;
  readonly totalsRow?: boolean | undefined;
}

/** Build a `WorkbookTableInfo` from an A1 range, the way a reader would. */
export function fakeTable(spec: FakeTableSpec): WorkbookTableInfo {
  const range = parseCellRange(spec.ref);
  if (!range) {
    throw new Error(`Test table range "${spec.ref}" is not a valid range.`);
  }
  return {
    columnNames: spec.columnNames,
    headerRow: spec.headerRow ?? range.start.row,
    name: spec.name,
    partPath: `xl/tables/${spec.name}.xml`,
    range,
    sheetName: spec.sheetName,
    totalsRow: spec.totalsRow ?? false,
  };
}

/** A formula cell for grids, defaulting to an ordinary (non-shared) formula. */
export function formulaCell(
  text: string,
  kind: CellFormula["kind"] = "normal",
  cached = "0",
): FakeCellSpec {
  return { formula: { kind, text }, text: cached };
}
