/**
 * L3: the workbook-keeping split read and written as streams (ADR 0006).
 *
 * The split keeps every part of the workbook and filters each worksheet that
 * carries the column. It used to load the whole package into the document
 * model once to analyse it, and again for every group. Now:
 *
 * - Every part but the worksheets is held, as small as it always is. Each
 *   worksheet is held as a stub: the text before its cells and the text after
 *   them, around an empty `sheetData`. The model, the values conversion and
 *   the pivot removal run on this light package unchanged, so every edit they
 *   make to tables, the calculation chain, comments, relationships and the
 *   worksheets' own merged ranges, conditional formats and the rest is the
 *   edit they always made.
 * - The rows are read from the source as a stream whenever they are needed:
 *   twice to analyse a worksheet, and once per output to write it. Each row
 *   goes through the same per-row steps the whole part went through: stale
 *   cached results blanked, formulas converted to values, and the row
 *   relocation the model applied to the stub, through the model's own row code.
 * - The output is written with `JsZipWriter`, which gives JSZip's bytes.
 *
 * What is held is the light package, one compressed output part, and per
 * filtered worksheet one group number per row.
 */
import { readWorkbookSheetsFrom } from "../excel-tables.js";
import type { RowRelocation } from "../model/references.js";
import { encodeCell } from "../model/references.js";
import type {
  CellRange,
  RowNumber,
  WorkbookTableInfo,
} from "../model/types.js";
import { WorkbookModel } from "../model/workbook-model.js";
import {
  cellTextOf,
  cellValueOf,
  declaredUsedRange,
  WorksheetRow,
  WorksheetModel,
  type WorksheetHost,
} from "../model/worksheet-model.js";
import {
  bytesPartFeed,
  readWorksheetPart,
  zipPartFeed,
  type PartFeed,
  type StreamedRow,
} from "../model/worksheet-stream.js";
import { findElement } from "../model/xml.js";
import { JsZipWriter, type ZipSink } from "../package/jszip-writer.js";
import {
  FIXED_PACKAGE_DATE,
  WorkbookPackage,
} from "../package/workbook-package.js";
import type { ZipReader } from "../package/zip-reader.js";
import {
  cellKey,
  detectHeaderRow,
  headerCellNameOf,
  isBlankValue,
  settleHeaderCandidate,
  type RowValueCount,
} from "../region/header-detection.js";
import {
  formulaGuardOfCell,
  RangeBinding,
  type FormulaGuardVerdict,
} from "../region/range-binding.js";
import { findMatchingTable } from "../region/resolve.js";
import { isTableEditReport, TableBinding } from "../region/table-binding.js";
import type { ColumnInfo, DataRegion } from "../region/types.js";
import {
  normalizeHeader,
  normalizeSplitValue,
  type NormalizedValue,
} from "../region/values.js";
import {
  blankStaleCells,
  type BlankedCachedFormula,
} from "../tier1/stale-values.js";
import { stripPivotPartsIn } from "../tier1/pivot.js";
import {
  convertTablesAndCalcChain,
  isConvertedWorksheetPart,
  removeWorksheetFormulas,
  worksheetNamesByPart,
} from "../values-only.js";

/** A worksheet held as the text around its cells. */
interface Stub {
  /** The stub's bytes as first held: the part is unedited while they stay. */
  readonly bytes: Uint8Array;
  /** Reads the original part again, `between` run after each piece. */
  readonly feed: (between?: () => Promise<void>) => PartFeed;
}

/** The workbook with its worksheets held as stubs, and the source to read rows from. */
export interface StreamedSplitPackage {
  readonly base: WorkbookPackage;
  readonly stubs: ReadonlyMap<string, Stub>;
  /**
   * Parts no step reads, held as an empty placeholder and copied from the
   * source when written, while the placeholder stands.
   */
  readonly opaque: ReadonlyMap<
    string,
    { bytes: Uint8Array; feed: (between?: () => Promise<void>) => PartFeed }
  >;
}

/** What the analysis settles about one worksheet that carries the column. */
export interface StreamedSheetRegion {
  readonly name: string;
  readonly worksheetPart: string;
  /** The Excel Table the region is, when it is one. */
  readonly table: WorkbookTableInfo | undefined;
  readonly body: CellRange;
  readonly columns: readonly ColumnInfo[] | undefined;
  readonly headerRow: RowNumber;
  readonly declared: boolean;
  readonly splitColumn: number;
  /** The highest row number the worksheet stores, as the model counts it. */
  readonly lastRow: number;
  /** The formula guard over the worksheet as written. */
  readonly guard: FormulaGuardVerdict;
  /** The formula guard once its formulas are converted to values. */
  readonly guardAfterValues: FormulaGuardVerdict;
  /** Per body row from `body.start.row`, the group index, or -1 for blank. */
  readonly groupOfRow: Int32Array;
}

const encoder = new TextEncoder();
const decoder = new TextDecoder();

/** Split a stub's text at its empty `sheetData`, as the model does. */
function stubPieces(xml: string): { prefix: string; suffix: string } {
  const sheetData = findElement(xml, "sheetData");
  if (!sheetData) {
    throw new Error("Worksheet package part does not contain sheetData.");
  }
  return sheetData.selfClosing
    ? { prefix: xml.slice(0, sheetData.end), suffix: xml.slice(sheetData.end) }
    : {
        prefix: xml.slice(0, sheetData.innerStart),
        suffix: xml.slice(sheetData.innerEnd),
      };
}

/**
 * Open a workbook for the streamed split: every part but the worksheets read,
 * each worksheet held as a stub, and `scan` called with each worksheet's rows
 * as they are read, the only time they are read for it.
 */
export async function openStreamedSplitPackage(
  zip: ZipReader,
  scanner: (listed: WorkbookPackage) => (
    part: string,
    prefix: string,
  ) =>
    | {
        row(row: StreamedRow): void;
        end(suffix: string): void;
      }
    | undefined,
): Promise<StreamedSplitPackage> {
  const entries = zip.entries();
  const held = new Map<string, Uint8Array>();
  const opaque = new Map<
    string,
    { bytes: Uint8Array; feed: (between?: () => Promise<void>) => PartFeed }
  >();
  for (const entry of entries) {
    if (isConvertedWorksheetPart(entry.name)) continue;
    if (isOpaquePart(entry.name)) {
      const bytes = new Uint8Array(0);
      opaque.set(entry.name, {
        bytes,
        feed: (between) => zipPartFeed(zip, entry.name, between),
      });
      held.set(entry.name, bytes);
      continue;
    }
    // Held whole, as the operation always held it, however large.
    held.set(entry.name, (await zip.readBytes(entry.name, Infinity))!);
  }
  // The worksheets the workbook lists, wherever they are, are read as
  // worksheets too.
  const listed = WorkbookPackage.fromParts(
    entries.flatMap((entry) => {
      const bytes = held.get(entry.name);
      return bytes ? [{ name: entry.name, bytes, date: entry.date }] : [];
    }),
  );
  const sheetParts = new Set<string>();
  for (const sheet of readWorkbookSheetsFrom(listed)) {
    sheetParts.add(sheet.worksheetPart);
  }
  const scan = scanner(listed);

  const stubs = new Map<string, Stub>();
  const parts: Array<{ name: string; bytes: Uint8Array; date: Date }> = [];
  for (const entry of entries) {
    const inMemory = held.get(entry.name);
    const isWorksheet =
      isConvertedWorksheetPart(entry.name) || sheetParts.has(entry.name);
    if (!isWorksheet) {
      parts.push({ name: entry.name, bytes: inMemory!, date: entry.date });
      continue;
    }
    const feed = (between?: () => Promise<void>): PartFeed =>
      inMemory
        ? bytesPartFeed(inMemory)
        : zipPartFeed(zip, entry.name, between);
    let prefix = "";
    let suffix = "";
    let rows: ReturnType<typeof scan>;
    try {
      await readWorksheetPart(feed(), {
        prefix: (text) => {
          prefix = text;
          rows = scan(entry.name, text);
        },
        row: (row) => rows?.row(row),
        text: () => undefined,
        suffix: (text) => {
          suffix = text;
          rows?.end(text);
        },
      });
    } catch (error) {
      // A part with no sheetData is no worksheet the model can read; held
      // whole, it is refused or ignored exactly as it was.
      if (
        error instanceof Error &&
        error.message === "Worksheet package part does not contain sheetData."
      ) {
        parts.push({
          name: entry.name,
          bytes: inMemory ?? (await zip.readBytes(entry.name))!,
          date: entry.date,
        });
        continue;
      }
      throw error;
    }
    const bytes = encoder.encode(prefix + suffix);
    stubs.set(entry.name, { bytes, feed });
    parts.push({ name: entry.name, bytes, date: entry.date });
  }
  return { base: WorkbookPackage.fromParts(parts), stubs, opaque };
}

/**
 * Parts the split's steps never read: anything not XML, relationships or VML,
 * and the pivot parts, which are removed by name. They can be the largest in
 * a package, so they stay in the source.
 */
function isOpaquePart(name: string): boolean {
  const lower = name.toLowerCase();
  return (
    !/\.(xml|rels|vml)$/u.test(lower) ||
    lower.startsWith("xl/pivotcache/") ||
    lower.startsWith("xl/pivottables/")
  );
}

/** Write a part that is not a worksheet stub: its bytes, or the source's. */
export async function writeHeldPart(
  splitPackage: StreamedSplitPackage,
  workbookPackage: WorkbookPackage,
  part: string,
  writing: { push(bytes: Uint8Array): void; close(): Promise<void> },
  between?: () => Promise<void>,
): Promise<void> {
  const bytes = workbookPackage.partBytes(part)!;
  const opaque = splitPackage.opaque.get(part);
  if (opaque && opaque.bytes === bytes) {
    await opaque.feed(between)((chunk) => writing.push(chunk));
  } else {
    writing.push(bytes);
  }
  await writing.close();
}

/** A held row the header rule may still need, with its header names. */
interface HeldRow {
  readonly row: number;
  /** Column by name, for the cells inside the used range. */
  readonly names: ReadonlyMap<number, string>;
  /** Columns holding a value. */
  readonly valued: readonly number[];
  readonly count: number;
}

/**
 * The first pass over one worksheet: what `resolveRegions` reads to find the
 * header cell and the region's columns, without keeping the rows. Rows the
 * header rule may still need are held until it settles; past that, a row only
 * moves the per-column last value row and the first cell naming the column.
 */
export class SheetSummary {
  readonly #host: WorksheetHost;
  readonly #target: string;
  readonly #declared: number | undefined;
  readonly #used: CellRange | undefined;
  readonly #values: boolean;
  readonly #prefix: string;
  #counts: RowValueCount[] = [];
  #held: HeldRow[] = [];
  #candidate = 0;
  #settled = false;
  #ordered = true;
  #previousRow = 0;
  #declaredNames: ReadonlyMap<number, string> | undefined;
  #match:
    { row: number; column: number; names: Map<number, string> } | undefined;
  lastValueRow: Map<number, number> = new Map();
  lastRow: number = 0;
  #cells = 0;
  #extent = {
    startRow: Infinity,
    endRow: 0,
    startColumn: Infinity,
    endColumn: -1,
  };
  guard: FormulaGuardVerdict | undefined;
  guardAfterValues: FormulaGuardVerdict | undefined;
  merges: readonly CellRange[] = [];

  constructor(
    host: WorksheetHost,
    column: string,
    declaredHeaderRow: number | undefined,
    prefix: string,
    values: boolean,
  ) {
    this.#host = host;
    this.#target = normalizeHeader(column);
    this.#declared = declaredHeaderRow;
    this.#used = declaredUsedRange(prefix);
    this.#prefix = prefix;
    this.#values = values;
  }

  /** The used range: the declared one, or the extent of every cell. */
  get usedRange(): CellRange | undefined {
    if (this.#used) return this.#used;
    if (this.#cells === 0) return undefined;
    return {
      start: { row: this.#extent.startRow, column: this.#extent.startColumn },
      end: { row: this.#extent.endRow, column: this.#extent.endColumn },
    };
  }

  #inUsedColumns(column: number): boolean {
    return (
      !this.#used ||
      (column >= this.#used.start.column && column <= this.#used.end.column)
    );
  }

  row(streamed: StreamedRow): void {
    const row = streamed.parse();
    const text = streamed.text;
    const implied = streamed.implied;
    this.lastRow = Math.max(this.lastRow, row.number);
    if (row.number <= this.#previousRow) this.#ordered = false;
    this.#previousRow = row.number;
    const cells = row.cells;
    for (const cell of cells) {
      this.#cells += 1;
      const extent = this.#extent;
      extent.startRow = Math.min(extent.startRow, cell.row);
      extent.endRow = Math.max(extent.endRow, cell.row);
      extent.startColumn = Math.min(extent.startColumn, cell.column);
      extent.endColumn = Math.max(extent.endColumn, cell.column);
      this.guard ??= formulaGuardOfCell(cell.formula, {
        row: cell.row,
        column: cell.column,
      });
    }
    if (this.#values && this.guardAfterValues === undefined) {
      const converted = removeWorksheetFormulas(text, "").xml;
      const after =
        converted === text ? row : new WorksheetRow(converted, implied);
      for (const cell of after.cells) {
        this.guardAfterValues ??= formulaGuardOfCell(cell.formula, {
          row: cell.row,
          column: cell.column,
        });
      }
    }

    const used = this.#used;
    if (used && (row.number < used.start.row || row.number > used.end.row)) {
      if (row.number === this.#declared) this.#declaredNames = new Map();
      return;
    }
    const keepNames =
      !this.#settled ||
      !this.#ordered ||
      row.number === this.#declared ||
      this.#match === undefined ||
      row.number <= this.#match.row;
    let names: Map<number, string> | undefined;
    const valued: number[] = [];
    for (const cell of cells) {
      if (!this.#inUsedColumns(cell.column)) continue;
      const value = cellValueOf(cell, this.#host);
      if (!isBlankValue(value)) {
        valued.push(cell.column);
        const last = this.lastValueRow.get(cell.column);
        if (last === undefined || last < row.number) {
          this.lastValueRow.set(cell.column, row.number);
        }
      }
      if (keepNames) {
        names ??= new Map();
        // The model looks a column up in a row by its last cell.
        names.set(
          cell.column,
          headerCellNameOf(value, () => cellTextOf(cell, this.#host)),
        );
      }
    }
    if (names) {
      if (row.number === this.#declared) this.#declaredNames = names;
      // The first cell, top to bottom then left to right, naming the column.
      const match = [...names]
        .filter(
          ([, name]) => name !== "" && normalizeHeader(name) === this.#target,
        )
        .map(([column]) => column)
        .sort((left, right) => left - right)[0];
      if (this.#match !== undefined && row.number === this.#match.row) {
        // A repeated row number: the model looks the row up by number, and
        // the last row with it wins.
        this.#match =
          match === undefined
            ? undefined
            : { row: row.number, column: match, names };
      } else if (
        match !== undefined &&
        (this.#match === undefined || row.number < this.#match.row)
      ) {
        this.#match = { row: row.number, column: match, names };
      }
    }
    if (valued.length === 0 || this.#declared !== undefined) return;
    if (this.#settled && this.#ordered) return;
    this.#counts.push({
      row: row.number,
      values: valued.length,
      bannerValues: 0,
    });
    this.#held.push({
      row: row.number,
      names: names ?? new Map(),
      valued,
      count: this.#counts.length - 1,
    });
    if (this.#ordered) {
      const { index, settled } = settleHeaderCandidate(
        this.#counts,
        this.#candidate,
      );
      this.#candidate = index;
      this.#settled = settled;
    }
  }

  /** The text after the rows, which declares the merged ranges. */
  end(suffix: string): void {
    this.merges = WorksheetModel.parse(
      this.#prefix + suffix,
      { name: "", partPath: "", visibility: "visible" },
      this.#host,
    ).mergedRanges();
  }

  /** The header cell, as `locateHeaderCell` finds it. */
  headerCell():
    | { row: number; column: number; names: ReadonlyMap<number, string> }
    | undefined {
    const used = this.usedRange;
    if (!used) return undefined;
    const inRow = (
      row: number,
      names: ReadonlyMap<number, string> | undefined,
    ) => {
      if (!names) return undefined;
      const column = [...names]
        .filter(
          ([, name]) => name !== "" && normalizeHeader(name) === this.#target,
        )
        .map(([candidate]) => candidate)
        .sort((left, right) => left - right)[0];
      return column === undefined ? undefined : { row, column, names };
    };
    if (this.#declared !== undefined) {
      return inRow(this.#declared, this.#declaredNames);
    }
    const banners = new Set<string>();
    for (const range of this.merges) {
      if (range.end.column > range.start.column) {
        banners.add(cellKey(range.start.row, range.start.column));
      }
    }
    const counts = this.#counts.map((count) => count);
    for (const held of this.#held) {
      const bannerValues = held.valued.filter((column) =>
        banners.has(cellKey(held.row, column)),
      ).length;
      counts[held.count] = { ...counts[held.count]!, bannerValues };
    }
    counts.sort((left, right) => left.row - right.row);
    const detected = detectHeaderRow(counts);
    const detectedRow =
      detected === undefined
        ? undefined
        : this.#held.findLast((held) => held.row === detected);
    return (
      (detectedRow ? inRow(detectedRow.row, detectedRow.names) : undefined) ??
      (this.#match
        ? {
            row: this.#match.row,
            column: this.#match.column,
            names: this.#match.names,
          }
        : undefined)
    );
  }
}

/** Builds the region a summary settles, as `bindingForHeader` does. */
export function summaryRegion(
  summary: SheetSummary,
  sheetName: string,
  columnText: string,
  tables: readonly WorkbookTableInfo[],
):
  | Omit<
      StreamedSheetRegion,
      | "worksheetPart"
      | "lastRow"
      | "guard"
      | "guardAfterValues"
      | "groupOfRow"
      | "name"
      | "declared"
    >
  | undefined {
  const header = summary.headerCell();
  if (!header) return undefined;
  const target = normalizeHeader(columnText);
  const table = findMatchingTable(
    tables,
    sheetName,
    header.row,
    header.column,
    columnText,
  );
  if (table) {
    const offset = table.columnNames.findIndex(
      (name) => normalizeHeader(name) === target,
    );
    if (offset < 0) return undefined;
    return {
      table,
      body: {
        end: {
          column: table.range.end.column,
          row: table.range.end.row - (table.totalsRow ? 1 : 0),
        },
        start: { column: table.range.start.column, row: table.headerRow + 1 },
      },
      columns: undefined,
      headerRow: table.headerRow,
      splitColumn: table.range.start.column + offset,
    };
  }
  const used = summary.usedRange!;
  const body = {
    end: { column: used.end.column, row: used.end.row },
    start: { column: used.start.column, row: header.row + 1 },
  };
  const columns: ColumnInfo[] = [];
  for (let column = body.start.column; column <= body.end.column; column += 1) {
    const last = summary.lastValueRow.get(column);
    if (last !== undefined && last >= header.row && last <= body.end.row) {
      columns.push({ index: column, name: header.names.get(column) ?? "" });
    }
  }
  const split = columns.find(
    (candidate) => normalizeHeader(candidate.name) === target,
  );
  if (!split) return undefined;
  return {
    table: undefined,
    body,
    columns,
    headerRow: header.row,
    splitColumn: split.index,
  };
}

/** The second pass: each body row's split value, and the uncached split cells. */
export async function readSplitColumn(
  feed: PartFeed,
  host: WorksheetHost,
  region: { body: CellRange; splitColumn: number },
  sheetName: string,
  strict: boolean,
  key: (value: NormalizedValue, row: number) => number,
  uncached: string[],
): Promise<Int32Array> {
  const first = region.body.start.row;
  const length = Math.max(0, region.body.end.row - first + 1);
  const groups = new Int32Array(length).fill(-1);
  const column = region.splitColumn;
  await readWorksheetPart(feed, {
    prefix: () => undefined,
    text: () => undefined,
    suffix: () => undefined,
    row: (streamed) => {
      if (streamed.number < first || streamed.number > region.body.end.row) {
        return;
      }
      const row = streamed.parse();
      const firstCell = row.cells.find((cell) => cell.column === column);
      if (firstCell?.formula !== undefined && !firstCell.hasCachedValue) {
        uncached.push(`${sheetName}!${encodeCell(column, row.number)}`);
      }
      const value = normalizeSplitValue(
        cellValueOf(row.cellAt(column), host),
        strict,
      );
      groups[row.number - first] =
        value === undefined ? -1 : key(value, row.number);
    },
  });
  return groups;
}

/** The light model of a split package, for its sheets, tables and cell values. */
export function lightModel(splitPackage: StreamedSplitPackage): WorkbookModel {
  return WorkbookModel.fromPackage(splitPackage.base.clone());
}

/** A step a worksheet's rows go through on the way out, in order. */
type RowStep =
  | {
      readonly kind: "blank";
      readonly sheet: string;
      readonly own: ReadonlySet<number> | undefined;
      readonly byName: ReadonlyMap<string, ReadonlySet<number>>;
    }
  | { readonly kind: "convert" }
  | { readonly kind: "relocate"; readonly relocation: RowRelocation };

/** Where a located cell was found in its part, so lists keep document order. */
interface PartFindings<T> {
  prefix: T[];
  rows: T[];
  suffix: T[];
}

function emptyFindings<T>(): PartFindings<T> {
  return { prefix: [], rows: [], suffix: [] };
}

export interface StreamedGroupOutput {
  calcChainEntriesRemoved: number;
  formulaCellsBlanked: number;
  formulaCellsConverted: number;
  formulaCellsWithoutCachedValues: number;
  pivotTablesRemoved: number;
  staleAggregates: string[];
  tableFallbackSheets: string[];
  uncachedFormulas: string[];
}

/**
 * Write one group's workbook to `sink`: the analysis's regions filtered to
 * `groupIndex`, as `buildGroupWorkbook` built it.
 */
export async function writeSplitGroup(
  splitPackage: StreamedSplitPackage,
  regions: readonly StreamedSheetRegion[],
  groupIndex: number,
  values: boolean,
  sink: ZipSink,
  between?: () => Promise<void>,
): Promise<StreamedGroupOutput> {
  const workbookPackage = splitPackage.base.clone();
  const steps = new Map<string, RowStep[]>();
  const addStep = (part: string, step: RowStep): void => {
    const list = steps.get(part) ?? [];
    list.push(step);
    steps.set(part, list);
  };
  const blanked = new Map<string, PartFindings<BlankedCachedFormula>>();
  const converted = new Map<
    string,
    { count: number; missing: PartFindings<string> }
  >();
  const output: StreamedGroupOutput = {
    calcChainEntriesRemoved: 0,
    formulaCellsBlanked: 0,
    formulaCellsConverted: 0,
    formulaCellsWithoutCachedValues: 0,
    pivotTablesRemoved: 0,
    staleAggregates: [],
    tableFallbackSheets: [],
    uncachedFormulas: [],
  };
  const containsFilteredTable = regions.some(
    (region) => region.table !== undefined,
  );
  const keeps = (region: StreamedSheetRegion, row: number): boolean =>
    region.groupOfRow[row - region.body.start.row] === groupIndex;

  // Stale cached results, then formulas, through the stubs; each step is
  // remembered so the rows take it too.
  const editStub = (
    part: string,
    edit: (piece: string, where: "prefix" | "suffix") => string,
  ): string => {
    const pieces = stubPieces(workbookPackage.readText(part)!);
    return edit(pieces.prefix, "prefix") + edit(pieces.suffix, "suffix");
  };

  const convert = (): void => {
    const names = worksheetNamesByPart(workbookPackage);
    for (const part of workbookPackage.partNames()) {
      if (!isConvertedWorksheetPart(part)) continue;
      const record = converted.get(part) ?? {
        count: 0,
        missing: emptyFindings<string>(),
      };
      converted.set(part, record);
      const locate = (cell: string): string =>
        `${names.get(part) ?? part}!${cell}`;
      if (splitPackage.stubs.has(part)) {
        // The whole-part conversion writes every worksheet back, changed or not.
        workbookPackage.writeText(
          part,
          editStub(part, (piece, where) => {
            const conversion = removeWorksheetFormulas(piece, part);
            record.count += conversion.formulasConverted;
            for (const missing of conversion.formulasWithoutCachedValues) {
              record.missing[where].push(locate(missing.cell));
            }
            return conversion.xml;
          }),
        );
        addStep(part, { kind: "convert" });
      } else {
        const conversion = removeWorksheetFormulas(
          workbookPackage.requireText(part),
          part,
        );
        workbookPackage.writeText(part, conversion.xml);
        record.count += conversion.formulasConverted;
        for (const missing of conversion.formulasWithoutCachedValues) {
          record.missing.rows.push(locate(missing.cell));
        }
      }
    }
    convertTablesAndCalcChain(workbookPackage);
  };

  if (values) {
    const deletedByPart = new Map<string, Set<number>>();
    for (const region of regions) {
      const removed = new Set<number>();
      for (
        let row = region.body.start.row;
        row <= region.body.end.row;
        row += 1
      ) {
        if (!keeps(region, row)) removed.add(row);
      }
      deletedByPart.set(region.worksheetPart, removed);
    }
    if ([...deletedByPart.values()].some((rows) => rows.size > 0)) {
      const identities = readWorkbookSheetsFrom(workbookPackage);
      const byName = new Map<string, ReadonlySet<number>>();
      for (const identity of identities) {
        const rows = deletedByPart.get(identity.worksheetPart);
        if (rows && rows.size > 0) {
          byName.set(identity.name.trim().toLowerCase(), rows);
        }
      }
      for (const identity of identities) {
        const part = identity.worksheetPart;
        if (workbookPackage.readText(part) === undefined) continue;
        const own = deletedByPart.get(part);
        const findings = emptyFindings<BlankedCachedFormula>();
        blanked.set(part, findings);
        if (splitPackage.stubs.has(part)) {
          const before = workbookPackage.readText(part)!;
          const after = editStub(part, (piece, where) =>
            blankStaleCells(piece, identity.name, own, byName, findings[where]),
          );
          if (after !== before) workbookPackage.writeText(part, after);
          addStep(part, { kind: "blank", sheet: identity.name, own, byName });
        } else {
          const xml = workbookPackage.readText(part)!;
          const rewritten = blankStaleCells(
            xml,
            identity.name,
            own,
            byName,
            findings.rows,
          );
          if (rewritten !== xml) workbookPackage.writeText(part, rewritten);
        }
      }
    }
    if (containsFilteredTable) convert();
  }

  const model = WorkbookModel.fromPackage(workbookPackage);
  const tables = await model.tables();
  for (const region of regions) {
    const worksheet = model.worksheet(region.name);
    if (!worksheet) continue;
    Object.defineProperty(worksheet, "lastRow", { get: () => region.lastRow });
    const apply = worksheet.applyRowRelocation.bind(worksheet);
    let captured: RowRelocation | undefined;
    worksheet.applyRowRelocation = (relocation, options) => {
      captured = relocation;
      return apply(relocation, options);
    };
    const binding = streamedBinding(worksheet, region, tables);
    const guard =
      values && containsFilteredTable ? region.guardAfterValues : region.guard;
    binding.formulaGuard = () => guard;
    const report = binding.filterRows((row) => keeps(region, row));
    if (isTableEditReport(report) && !report.tableResized) {
      output.tableFallbackSheets.push(region.name);
    }
    if (captured) {
      addStep(region.worksheetPart, {
        kind: "relocate",
        relocation: captured,
      });
    }
  }
  model.flush();
  output.calcChainEntriesRemoved = values ? 0 : model.calcChainEntriesRemoved;

  if (values && !containsFilteredTable) convert();

  const stripped = stripPivotPartsIn(workbookPackage);
  output.pivotTablesRemoved = stripped?.removedPivotTables ?? 0;

  // Write every part, the worksheets' rows streamed through their steps.
  const writer = new JsZipWriter(sink);
  for (const part of workbookPackage.partNames()) {
    const stub = splitPackage.stubs.get(part);
    const current = workbookPackage.partBytes(part)!;
    const partSteps = steps.get(part) ?? [];
    if (!stub) {
      await writeHeldPart(
        splitPackage,
        workbookPackage,
        part,
        writer.part(part, workbookPackage.partDate(part)),
        between,
      );
      continue;
    }
    if (current === stub.bytes && partSteps.length === 0) {
      const writing = writer.part(part, workbookPackage.partDate(part));
      await stub.feed(between)((chunk) => writing.push(chunk));
      await writing.close();
      continue;
    }
    const pieces = stubPieces(decoder.decode(current));
    const findings = blanked.get(part);
    const conversion = converted.get(part);
    let rowsBlanked = false;
    let previousDestination = 0;
    const through = (
      text: string,
      row: StreamedRow | undefined,
    ): string | undefined => {
      let result = text;
      let parsed: WorksheetRow | undefined;
      let fresh = true;
      for (const step of partSteps) {
        if (step.kind === "blank") {
          const next = blankStaleCells(
            result,
            step.sheet,
            step.own,
            step.byName,
            findings!.rows,
          );
          if (next !== result) {
            rowsBlanked = true;
            result = next;
            fresh = false;
          }
        } else if (step.kind === "convert") {
          const conversionResult = removeWorksheetFormulas(result, part);
          conversion!.count += conversionResult.formulasConverted;
          for (const missing of conversionResult.formulasWithoutCachedValues) {
            conversion!.missing.rows.push(missing.cell);
          }
          if (conversionResult.xml !== result) {
            result = conversionResult.xml;
            fresh = false;
          }
        } else if (row !== undefined) {
          // The model parses the text the earlier steps left, moves the row
          // and writes it back. The earlier steps never touch a row's number.
          const destination = step.relocation.target(row.number);
          if (destination === null) return undefined;
          parsed =
            fresh && parsed === undefined
              ? row.parse()
              : new WorksheetRow(result, row.implied);
          if (destination <= previousDestination) {
            throw new Error(
              "A row relocation would reorder worksheet rows; relocations must preserve row order.",
            );
          }
          previousDestination = destination;
          parsed.relocateFormulas(step.relocation);
          parsed.moveTo(destination);
          result = parsed.toXml();
          fresh = false;
        }
      }
      return result;
    };
    const writing = writer.part(part, workbookPackage.partDate(part));
    writing.text(pieces.prefix);
    await readWorksheetPart(stub.feed(between), {
      prefix: () => undefined,
      suffix: () => undefined,
      text: (text) => {
        const out = through(text, undefined);
        if (out !== undefined) writing.text(out);
      },
      row: (row) => {
        const out = through(row.text, row);
        if (out !== undefined) writing.text(out);
      },
    });
    writing.text(pieces.suffix);
    // Rows a stale result was cleared in were rewritten, so the part was too.
    await writing.close(
      rowsBlanked ? FIXED_PACKAGE_DATE : workbookPackage.partDate(part),
    );
  }
  await writer.finish();

  // Findings in the order the whole-part steps found them.
  const names = worksheetNamesByPart(workbookPackage);
  for (const identity of readWorkbookSheetsFrom(splitPackage.base)) {
    const findings = blanked.get(identity.worksheetPart);
    if (!findings) continue;
    for (const cell of [
      ...findings.prefix,
      ...findings.rows,
      ...findings.suffix,
    ]) {
      output.formulaCellsBlanked += 1;
      output.staleAggregates.push(`${cell.sheet}!${cell.cell}`);
    }
  }
  for (const [part, record] of converted) {
    output.formulaCellsConverted += record.count;
    const rows = splitPackage.stubs.has(part)
      ? record.missing.rows.map((cell) => `${names.get(part) ?? part}!${cell}`)
      : record.missing.rows;
    for (const location of [
      ...record.missing.prefix,
      ...rows,
      ...record.missing.suffix,
    ]) {
      output.formulaCellsWithoutCachedValues += 1;
      output.uncachedFormulas.push(location);
    }
  }
  return output;
}

/** A region bound to its stub, with the analysis's answers in place of the rows. */
function streamedBinding(
  worksheet: WorksheetModel,
  region: StreamedSheetRegion,
  tables: readonly WorkbookTableInfo[],
): DataRegion & { formulaGuard: () => FormulaGuardVerdict } {
  if (region.table) {
    const table =
      tables.find(
        (candidate) =>
          candidate.name === region.table!.name &&
          candidate.sheetName === region.table!.sheetName,
      ) ?? region.table;
    return new TableBinding(worksheet, table);
  }
  return new RangeBinding({
    body: region.body,
    columns: region.columns,
    headerRow: region.headerRow,
    origin: region.declared
      ? { kind: "declared-header" }
      : { kind: "detected-header" },
    worksheet,
  });
}
