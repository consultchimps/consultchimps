/**
 * L3: the region splits read and written as streams (ADR 0006): a split of one
 * Excel Table, named range or worksheet, written as a compact workbook of
 * values or, for an Excel Table, as the whole workbook with only that table's
 * rows changed.
 *
 * These splits used to read the region into a table in memory, copy it into
 * one table per group, and build every output whole. Now the region is read
 * three ways, each a stream:
 *
 * - once per candidate worksheet, to learn which candidates hold rows, as the
 *   table readers decide it;
 * - once more for the chosen region, giving each row its group, and each
 *   group its row count and column widths, which a compact output needs before
 *   its first row;
 * - once per batch of groups, writing each group's rows as they are read.
 *
 * What is held is one group number per region row and, per group, its widths.
 * A worksheet whose rows are stored out of order is read whole, as the table
 * readers read it.
 */
import {
  openTableWriter,
  XLSX_TABLE_OUTPUT,
  type TableOutputFormat,
  type TableRowWriter,
} from "../table-output.js";
import {
  ConsultChimpsError,
  type RandomAccessSource,
} from "@consultchimps/core";
import {
  columnKey,
  uniqueHeaders,
  type CellValue,
} from "@consultchimps/tabular";

import { XLSX_ERRORS } from "../errors.js";
import type { ExcelTableDefinition } from "../excel-tables.js";
import { writableCellValue } from "../model/date-cells.js";
import { decodeRange, type CellRectangle } from "../model/references.js";
import { SheetProfile } from "../operations/consolidate/consolidate.js";
import type { CsvReadOptions } from "../csv/options.js";
import type {
  SheetBook,
  StreamedCell,
  StreamedSheet,
  WorksheetRead,
} from "../operations/consolidate/reader.js";
import { openSheetBook } from "../operations/sheet-book.js";
import { tableValue } from "../operations/sheet-grid.js";
import {
  readWorksheetPart,
  type StreamedRow,
} from "../model/worksheet-stream.js";
import { findElement } from "../model/xml.js";
import { JsZipWriter } from "../package/jszip-writer.js";
import { cellWidthLength, tableColumnWidth } from "../package/table-writer.js";
import { ZipReader } from "../package/zip-reader.js";
import {
  relocateTableRow,
  rewriteTablePart,
  tableRewrite,
  tableRowElementName,
  tableRowNumber,
  tableSheetDataBounds,
  filterWholeWorksheetRows,
} from "../preserve-table-split.js";
import { isBlankValue } from "../region/header-detection.js";
import { normalizeSplitValue } from "../region/values.js";
import { stripPivotPartsIn } from "../tier1/pivot.js";
import {
  uncachedFormulaHint,
  uncachedLocationsWithin,
} from "../uncached-formulas.js";
import {
  convertTablesAndCalcChain,
  isConvertedWorksheetPart,
  removeWorksheetFormulas,
  worksheetNamesByPart,
} from "../values-only.js";
import {
  openStreamedSplitPackage,
  writeHeldPart,
  type StreamedSplitPackage,
} from "./streamed-package.js";

export interface RegionSplitSelection {
  column: string;
  headerRow?: number | undefined;
  includeBlank?: boolean | undefined;
  includeHiddenSheets?: boolean | undefined;
  range?: string | undefined;
  sheet?: string | undefined;
  table?: string | undefined;
  /** How to read the input when it is a CSV file (ADR 0007). */
  csv?: CsvReadOptions | undefined;
  /**
   * Group by the default split matching (trimmed, case folded, numeric text
   * as numbers) unless `strict`, as the whole-workbook split does. A CSV
   * input's split asks for it, since that is its default split.
   */
  tolerantMatching?: boolean | undefined;
  strict?: boolean | undefined;
}

export interface RegionSplitContext {
  label: string;
  file: string;
  details: Record<string, unknown>;
}

/** One group of a region split. */
export interface RegionGroup {
  /** The split value, as the table holds it. */
  readonly value: CellValue;
  readonly rows: number;
  /** Per column, the width a compact output gives it. */
  readonly widths: readonly number[];
}

/** Everything a region split learned by reading its source. */
export interface ResolvedRegionSplit {
  readonly workbook: SheetBook;
  readonly sheet: StreamedSheet;
  /** Whether the region's worksheet had to be read whole, out of order. */
  readonly gathered: boolean;
  /** Whether its cells are read inside the used range only. */
  readonly clip: boolean;
  /** The region's columns, zero-based, and their table names. */
  readonly columns: readonly number[];
  readonly names: readonly string[];
  /** The matched split column's name. */
  readonly column: string;
  /** Zero-based first and last row the region's records can be on. */
  readonly first: number;
  readonly last: number;
  /** Per row from `first`, the group index, or -1. */
  readonly groupOfRow: Int32Array;
  readonly groups: readonly RegionGroup[];
  /** The most characters any one cell of the region holds. */
  readonly longestCell: number;
  /** How the input is named in messages, and its error details. */
  readonly label: string;
  readonly details: Record<string, unknown>;
  readonly inputRows: number;
  readonly skippedRows: number;
  /** `Sheet!B4` of every formula with no cached value inside the region. */
  readonly uncachedFormulas: readonly string[];
  /** The sheet a compact output is named after. */
  readonly sheetName: string;
  /** The Excel Table, for a split that keeps the workbook. */
  readonly definition: ExcelTableDefinition | undefined;
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
  return { range: range.replaceAll("$", "").toUpperCase(), sheet };
}

function lowercaseSet(values: string[] | undefined): Set<string> | undefined {
  return values
    ? new Set(values.map((value) => value.toLocaleLowerCase()))
    : undefined;
}

function isVisibleSheet(workbook: SheetBook, name: string): boolean {
  return workbook.sheets.find((sheet) => sheet.name === name)?.visible ?? true;
}

/** A region the table readers would read, before its rows are read. */
interface Candidate {
  readonly kind: "worksheet" | "table" | "range";
  readonly sheet: StreamedSheet;
  readonly name: string;
  readonly rangeRef?: string;
  readonly definition?: ExcelTableDefinition;
  /** Filled by the first read. */
  hasRows: boolean;
  read?: WorksheetRead;
  headerValues?: Array<string | null>;
  profile?: ReturnType<SheetProfile["finish"]>;
}

/** Rows of a fixed rectangle: whether any record is there, and its header. */
class RectangleScan {
  hasRows = false;
  header: Map<number, CellValue> = new Map();
  constructor(
    readonly rectangle: CellRectangle,
    readonly firstDataRow: number,
    readonly lastDataRow: number,
  ) {}

  begin(): void {
    this.hasRows = false;
    this.header = new Map();
  }

  row(row: number, cells: readonly StreamedCell[]): void {
    const { startColumn, endColumn } = this.rectangle;
    if (row === this.rectangle.startRow) {
      for (const cell of cells) {
        if (cell.column >= startColumn && cell.column <= endColumn) {
          this.header.set(cell.column, tableValue(cell.value));
        }
      }
    }
    if (this.hasRows || row < this.firstDataRow || row > this.lastDataRow) {
      return;
    }
    for (const cell of cells) {
      if (cell.column < startColumn || cell.column > endColumn) continue;
      const value = tableValue(cell.value);
      if (value !== null && value !== "") {
        this.hasRows = true;
        return;
      }
    }
  }
}

/** Several consumers fed by one read of a worksheet. */
function fanOut(
  consumers: ReadonlyArray<{
    begin(): void;
    row(row: number, cells: readonly StreamedCell[]): void;
  }>,
): { begin(): void; row(row: number, cells: readonly StreamedCell[]): void } {
  return {
    begin: () => consumers.forEach((consumer) => consumer.begin()),
    row: (row, cells) =>
      consumers.forEach((consumer) => consumer.row(row, cells)),
  };
}

/** The regions the selection names, in the order the table readers list them. */
function candidatesOf(
  workbook: SheetBook,
  selection: RegionSplitSelection,
): Candidate[] {
  const selectedSheets = lowercaseSet(
    selection.sheet ? [selection.sheet] : undefined,
  );
  const sheetNamed = (name: string): StreamedSheet | undefined =>
    workbook.sheets.find((sheet) => sheet.name === name);
  const candidates: Candidate[] = [];
  if (selection.table) {
    for (const definition of workbook.tables) {
      if (
        !selection.includeHiddenSheets &&
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
      const sheet = sheetNamed(definition.sheet);
      if (!sheet) continue;
      candidates.push({
        kind: "table",
        sheet,
        name: definition.name,
        definition,
        hasRows: false,
      });
    }
    return candidates;
  }
  if (selection.range) {
    for (const definedName of workbook.names) {
      if (
        !definedName.name ||
        definedName.name.startsWith(BUILTIN_DEFINED_NAME_PREFIX)
      ) {
        continue;
      }
      const parsed = parseNamedRangeRef(definedName.reference);
      if (!parsed) continue;
      if (
        !selection.includeHiddenSheets &&
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
      const sheet = sheetNamed(parsed.sheet);
      if (!sheet) continue;
      candidates.push({
        kind: "range",
        sheet,
        name: definedName.name,
        rangeRef: parsed.range,
        hasRows: false,
      });
    }
    return candidates;
  }
  for (const sheet of workbook.sheets) {
    if (!selection.includeHiddenSheets && !sheet.visible) continue;
    if (selectedSheets && !selectedSheets.has(sheet.name.toLocaleLowerCase())) {
      continue;
    }
    candidates.push({
      kind: "worksheet",
      sheet,
      name: sheet.name,
      hasRows: false,
    });
  }
  return candidates;
}

/** The rectangle an Excel Table's records lie in, checked as the reader checks it. */
function tableRectangle(definition: ExcelTableDefinition): {
  range: CellRectangle;
  first: number;
  last: number;
} {
  const range = decodeRange(definition.range.toUpperCase());
  const rangeColumnCount = range.endColumn - range.startColumn + 1;
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
  return {
    range,
    first: range.startRow + (definition.headerRow ? 1 : 0),
    last: range.endRow - (definition.totalsRow ? 1 : 0),
  };
}

/**
 * Read every candidate's worksheet once, each worksheet once however many
 * candidates sit on it, and learn which candidates hold rows.
 */
async function readCandidates(
  workbook: SheetBook,
  candidates: Candidate[],
  headerRow: number | undefined,
): Promise<void> {
  const read = new Set<StreamedSheet>();
  for (const candidate of candidates) {
    if (!read.has(candidate.sheet)) {
      read.add(candidate.sheet);
      const onSheet = candidates.filter(
        (other) => other.sheet === candidate.sheet,
      );
      const scans = new Map<Candidate, RectangleScan | SheetProfile>();
      for (const other of onSheet) {
        if (other.kind === "worksheet") {
          scans.set(other, new SheetProfile(headerRow, false));
        } else if (other.kind === "range") {
          const range = decodeRange(other.rangeRef!);
          scans.set(
            other,
            new RectangleScan(range, range.startRow + 1, range.endRow),
          );
        } else {
          const range = decodeRange(other.definition!.range.toUpperCase());
          if (
            range.endColumn - range.startColumn + 1 ===
            other.definition!.columns.length
          ) {
            const { first, last } = tableRectangle(other.definition!);
            scans.set(other, new RectangleScan(range, first, last));
          }
        }
      }
      const worksheetRead = await workbook.readWorksheet(
        candidate.sheet,
        fanOut([...scans.values()]),
        // Worksheet tables are read inside the used range; a named range or
        // an Excel Table declares its own rectangle.
        { clip: onSheet.every((other) => other.kind === "worksheet") },
      );
      for (const other of onSheet) {
        other.read = worksheetRead;
        const scan = scans.get(other);
        if (scan instanceof SheetProfile) {
          other.profile = scan.finish(worksheetRead, headerRow);
          other.hasRows = other.profile.table !== undefined;
        } else if (scan) {
          other.hasRows = scan.hasRows;
          if (other.kind === "range") {
            const range = decodeRange(other.rangeRef!);
            const headers: Array<string | null> = [];
            for (
              let column = range.startColumn;
              column <= range.endColumn;
              column += 1
            ) {
              const value = scan.header.get(column) ?? null;
              headers.push(value === null ? null : String(value));
            }
            other.headerValues = headers;
          }
        }
      }
    }
    // The table reader checks an Excel Table's columns once its worksheet is read.
    if (candidate.kind === "table") tableRectangle(candidate.definition!);
  }
}

/** The sheets' formulas with no cached value, for the refusal that names none. */
async function uncachedOnSheets(
  workbook: SheetBook,
  sheetName: string | undefined,
  includeHiddenSheets: boolean | undefined,
): Promise<string[]> {
  const locations: string[] = [];
  for (const sheet of workbook.sheets) {
    if (
      sheetName === undefined
        ? !includeHiddenSheets && !sheet.visible
        : sheet.name.toLocaleLowerCase() !== sheetName.toLocaleLowerCase()
    ) {
      continue;
    }
    const read = await workbook.readWorksheet(
      sheet,
      { begin: () => undefined, row: () => undefined },
      { clip: false },
    );
    for (const location of uncachedLocationsWithin(
      sheet.name,
      read.uncachedFormulas,
      undefined,
    )) {
      locations.push(location);
    }
  }
  return locations;
}

/**
 * Settle a region split from its source: the one region it reads, each row's
 * group, and each group's widths, refusing exactly what the table readers and
 * the grouping refused.
 */
export async function resolveRegionSplit(
  source: RandomAccessSource,
  context: RegionSplitContext,
  selection: RegionSplitSelection,
  preserveWorkbook: boolean,
): Promise<ResolvedRegionSplit> {
  const workbook = await openSheetBook(
    source,
    {
      file: context.file,
      source: context.label,
      details: context.details,
    },
    selection.csv,
  );
  const candidates = candidatesOf(workbook, selection);
  await readCandidates(workbook, candidates, selection.headerRow);
  const requested = (selection.table ?? selection.range)?.toLocaleLowerCase();
  const found = candidates.filter(
    (candidate) =>
      candidate.hasRows &&
      (requested === undefined ||
        candidate.name.toLocaleLowerCase() === requested),
  );

  if (found.length === 0) {
    const uncached = await uncachedOnSheets(
      workbook,
      selection.sheet,
      selection.includeHiddenSheets,
    );
    const selectedSource = selection.table
      ? `Excel Table "${selection.table}"`
      : selection.range
        ? `Named range "${selection.range}"`
        : selection.sheet
          ? `Worksheet "${selection.sheet}"`
          : undefined;
    const available = (kind: Candidate["kind"]) =>
      candidates
        .filter((candidate) => candidate.kind === kind && candidate.hasRows)
        .map((candidate) => ({
          name: candidate.name,
          sheet: candidate.sheet.name,
        }));
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_NO_TABLE,
      selectedSource
        ? `${selectedSource} was not found or has no data rows.${uncachedFormulaHint(uncached)}`
        : `No visible, non-empty worksheet was found in the input workbook.${uncachedFormulaHint(uncached)}`,
      {
        details: {
          uncachedFormulas: uncached,
          availableRanges: available("range"),
          availableTables: available("table"),
          ...context.details,
          range: selection.range,
          sheet: selection.sheet,
          table: selection.table,
        },
      },
    );
  }
  if (found.length > 1) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_MULTIPLE_TABLES,
      selection.table
        ? `Excel Table "${selection.table}" was found on multiple worksheets; choose one with the sheet option.`
        : selection.range
          ? `Named range "${selection.range}" is defined more than once; choose a worksheet with the sheet option.`
          : "The workbook contains multiple non-empty worksheets; choose one with the sheet option.",
      {
        details: {
          availableSheets: found.map((candidate) => candidate.sheet.name),
          ...context.details,
        },
      },
    );
  }

  const chosen = found[0]!;
  const read = chosen.read!;
  let columns: number[];
  let names: string[] | undefined;
  let first: number;
  let last: number;
  let rectangle: CellRectangle;
  let headerRowIndex: number | undefined;
  if (chosen.kind === "worksheet") {
    const table = chosen.profile!.table!;
    const range = read.range!;
    columns = [...table.columns];
    headerRowIndex = table.headerRow;
    first = table.headerRow + 1;
    last = range.endRow;
    rectangle = { ...range, startRow: table.headerRow };
  } else if (chosen.kind === "table") {
    const {
      range,
      first: firstRow,
      last: lastRow,
    } = tableRectangle(chosen.definition!);
    columns = chosen.definition!.columns.map(
      (_, index) => range.startColumn + index,
    );
    names = uniqueHeaders(
      chosen.definition!.columns.map((column) => column || null),
    );
    first = firstRow;
    last = lastRow;
    rectangle = range;
  } else {
    const range = decodeRange(chosen.rangeRef!);
    columns = [];
    for (
      let column = range.startColumn;
      column <= range.endColumn;
      column += 1
    ) {
      columns.push(column);
    }
    names = uniqueHeaders(chosen.headerValues!);
    first = range.startRow + 1;
    last = range.endRow;
    rectangle = range;
  }
  const uncachedFormulas = uncachedLocationsWithin(
    chosen.sheet.name,
    read.uncachedFormulas,
    rectangle,
  );

  // The grouping read: each record's group, each group's widths.
  const length = Math.max(0, last - first + 1);
  const groupOfRow = new Int32Array(length).fill(-1);
  const includeBlank = selection.includeBlank ?? true;
  let column: string | undefined;
  let matched = -1;
  const groups: Array<{ value: CellValue; rows: number; widths: number[] }> =
    [];
  const groupIndex = new Map<string, number>();
  let inputRows = 0;
  let skippedRows = 0;
  const settleNames = (headerCells: ReadonlyMap<number, CellValue>): void => {
    names ??= uniqueHeaders(
      columns.map((index) => {
        const value = headerCells.get(index) ?? null;
        return isBlankValue(value) ? null : String(value);
      }),
    );
    const requestedKey = columnKey(selection.column);
    matched = names.findIndex((name) => columnKey(name) === requestedKey);
    if (matched < 0) {
      const error = new ConsultChimpsError(
        "TABLE_COLUMN_NOT_FOUND",
        `Column "${selection.column}" was not found in the table.`,
        { details: { availableColumns: names, column: selection.column } },
      );
      if (uncachedFormulas.length === 0) throw error;
      throw new ConsultChimpsError(
        error.code,
        `${error.message}${uncachedFormulaHint([...uncachedFormulas])}`,
        {
          cause: error,
          details: {
            ...error.details,
            uncachedFormulas: [...uncachedFormulas],
          },
        },
      );
    }
    column = names[matched]!;
  };
  if (names) settleNames(new Map());
  const headerCells = new Map<number, CellValue>();
  let begun = false;
  await workbook.readWorksheet(
    chosen.sheet,
    {
      begin: () => {
        if (begun) {
          groupOfRow.fill(-1);
          groups.length = 0;
          groupIndex.clear();
          inputRows = 0;
          skippedRows = 0;
          headerCells.clear();
        }
        begun = true;
      },
      row: (row, cells) => {
        if (row === headerRowIndex) {
          for (const cell of cells)
            headerCells.set(cell.column, tableValue(cell.value));
          return;
        }
        if (row < first || row > last) return;
        if (matched < 0) settleNames(headerCells);
        const byColumn = new Map<number, CellValue>();
        for (const cell of cells)
          byColumn.set(cell.column, tableValue(cell.value));
        const values = columns.map((index) => byColumn.get(index) ?? null);
        if (values.every((value) => value === null || value === "")) return;
        inputRows += 1;
        const raw = values[matched] ?? null;
        const value =
          raw === null || (typeof raw === "string" && raw.trim() === "")
            ? null
            : raw;
        if (value === null && !includeBlank) {
          skippedRows += 1;
          return;
        }
        const key =
          value === null
            ? "null"
            : selection.tolerantMatching === true
              ? (normalizeSplitValue(value, selection.strict === true)?.key ??
                "null")
              : `${typeof value}:${String(value)}`;
        let group = groupIndex.get(key);
        if (group === undefined) {
          group = groups.length;
          groupIndex.set(key, group);
          groups.push({
            value,
            rows: 0,
            widths: names!.map((name) => name.length),
          });
        }
        const entry = groups[group]!;
        entry.rows += 1;
        values.forEach((cellValue, position) => {
          entry.widths[position] = Math.max(
            entry.widths[position]!,
            cellWidthLength(cellValue),
          );
        });
        groupOfRow[row - first] = group;
      },
    },
    { clip: chosen.kind === "worksheet", gather: read.gathered },
  );
  if (matched < 0) settleNames(headerCells);
  // A CSV field can be longer than a cell holds; a workbook's cannot. The
  // writer refuses it before it opens any output.
  let longestCell = 0;
  for (const group of groups) {
    for (const width of group.widths) {
      longestCell = Math.max(longestCell, width);
    }
  }

  if (groups.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_NO_GROUPS,
      `No output groups remain for column "${column}".${uncachedFormulaHint([...uncachedFormulas])}`,
      {
        details: {
          uncachedFormulas: [...uncachedFormulas],
          column,
          includeBlank: selection.includeBlank ?? true,
          ...context.details,
        },
      },
    );
  }

  let definition: ExcelTableDefinition | undefined;
  if (preserveWorkbook) {
    definition = workbook.tables.find(
      (candidate) =>
        candidate.name.toLocaleLowerCase() ===
          selection.table?.toLocaleLowerCase() &&
        candidate.sheet.toLocaleLowerCase() ===
          chosen.sheet.name.toLocaleLowerCase(),
    );
    if (!definition) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_SPLIT_PRESERVE_TABLE_NOT_FOUND,
        `Excel Table "${selection.table}" could not be located in the workbook package.`,
        {
          details: {
            ...context.details,
            sheet: chosen.sheet.name,
            table: selection.table,
          },
        },
      );
    }
  }

  return {
    workbook,
    sheet: chosen.sheet,
    gathered: read.gathered,
    clip: chosen.kind === "worksheet",
    columns,
    names: names!,
    column: column!,
    first,
    last,
    groupOfRow,
    groups: groups.map((group) => ({
      value: group.value,
      rows: group.rows,
      widths: group.widths.map(tableColumnWidth),
    })),
    longestCell,
    label: context.label,
    details: context.details,
    inputRows,
    skippedRows,
    uncachedFormulas,
    sheetName: chosen.sheet.name,
    definition,
  };
}

/** Where one output's bytes go. */
export interface RegionOutputTarget {
  write(chunk: Uint8Array): void | Promise<void>;
  close(): void | Promise<void>;
  abort?(): void | Promise<void>;
}

/** Groups written per read of the region in a compact split. */
const COMPACT_BATCH = 16;

/**
 * Write each group as a compact workbook of the region's values, in batches,
 * one read of the region per batch. The bytes are `buildTableWorkbookBytes`'s.
 */
/**
 * Refuse groups an Excel worksheet cannot hold, before any output is opened.
 * Workbook input always fits; a CSV file need not.
 */
export function assertGroupsFitWorksheet(resolved: ResolvedRegionSplit): void {
  let rows = 0;
  for (const group of resolved.groups) rows = Math.max(rows, group.rows + 1);
  const columns = resolved.names.length;
  const problem =
    resolved.longestCell > 32_767
      ? `a cell of ${resolved.longestCell.toLocaleString("en-US")} characters, more than the 32,767 an Excel cell holds`
      : rows > 1_048_576
        ? `a group of ${rows.toLocaleString("en-US")} rows with its header, more than the 1,048,576 a worksheet holds`
        : columns > 16_384
          ? `${columns.toLocaleString("en-US")} columns, more than the 16,384 a worksheet holds`
          : undefined;
  if (problem === undefined) return;
  throw new ConsultChimpsError(
    XLSX_ERRORS.XLSX_OUTPUT_TOO_LARGE,
    `${resolved.label} has ${problem}, so nothing was written.`,
    {
      details: {
        ...resolved.details,
        longestCell: resolved.longestCell,
        rows,
        columns,
      },
    },
  );
}

export async function writeCompactGroups(
  resolved: ResolvedRegionSplit,
  open: (index: number) => Promise<RegionOutputTarget> | RegionOutputTarget,
  between: () => Promise<void>,
  format: TableOutputFormat = XLSX_TABLE_OUTPUT,
): Promise<void> {
  // A worksheet's limits bind a workbook only; a CSV file has none.
  if (format.kind === "xlsx") assertGroupsFitWorksheet(resolved);
  for (let start = 0; start < resolved.groups.length; start += COMPACT_BATCH) {
    const indexes = resolved.groups
      .map((_, index) => index)
      .slice(start, start + COMPACT_BATCH);
    const queues = new Map<number, Uint8Array[]>();
    const targets = new Map<number, RegionOutputTarget>();
    const writers = new Map<number, TableRowWriter>();
    try {
      for (const index of indexes) {
        const group = resolved.groups[index]!;
        const queue: Uint8Array[] = [];
        queues.set(index, queue);
        targets.set(index, await open(index));
        writers.set(
          index,
          openTableWriter(format, {
            sheetName: resolved.sheetName,
            columns: resolved.names,
            widths: group.widths,
            rowCount: group.rows,
            onChunk: (chunk) => queue.push(chunk),
          }),
        );
      }
      let begun = false;
      const drain = async (): Promise<void> => {
        for (const [index, queue] of queues) {
          const target = targets.get(index)!;
          for (const chunk of queue.splice(0)) await target.write(chunk);
        }
      };
      await resolved.workbook.readWorksheet(
        resolved.sheet,
        {
          begin: () => {
            // The grouping read settled the order, so a read that starts over
            // means the worksheet changed between the two.
            if (begun) {
              throw new ConsultChimpsError(
                XLSX_ERRORS.XLSX_READ_FAILED,
                `Worksheet "${resolved.sheet.name}" changed while it was being split.`,
                { details: { worksheet: resolved.sheet.name } },
              );
            }
            begun = true;
          },
          row: (row, cells) => {
            if (row < resolved.first || row > resolved.last) return;
            const writer = writers.get(
              resolved.groupOfRow[row - resolved.first]!,
            );
            if (!writer) return;
            const byColumn = new Map<number, CellValue>();
            for (const cell of cells) {
              byColumn.set(cell.column, tableValue(cell.value));
            }
            writer.writeRow(
              resolved.columns.map((index) =>
                writableCellValue(byColumn.get(index) ?? null),
              ),
            );
          },
        },
        {
          clip: resolved.clip,
          gather: resolved.gathered,
          between: async () => {
            await drain();
            await between();
          },
        },
      );
      for (const writer of writers.values()) writer.finish();
      await drain();
      for (const index of indexes) await targets.get(index)!.close();
    } catch (error) {
      for (const target of targets.values()) {
        await Promise.resolve(target.abort?.()).catch(() => undefined);
      }
      throw error;
    }
  }
}

/** The light package a table-keeping split writes from, read once. */
export async function openRegionPackage(
  source: RandomAccessSource,
): Promise<StreamedSplitPackage> {
  return openStreamedSplitPackage(
    await ZipReader.open(source),
    () => () => undefined,
  );
}

/**
 * The cells a values conversion of the whole workbook finds with no cached
 * value, in the order the conversion finds them, read without changing it.
 */
export async function uncachedForValues(
  splitPackage: StreamedSplitPackage,
): Promise<string[]> {
  const base = splitPackage.base;
  const names = worksheetNamesByPart(base);
  const locations: string[] = [];
  for (const part of base.partNames()) {
    if (!isConvertedWorksheetPart(part)) continue;
    const located = (xml: string): void => {
      for (const missing of removeWorksheetFormulas(xml, part)
        .formulasWithoutCachedValues) {
        locations.push(`${names.get(part) ?? part}!${missing.cell}`);
      }
    };
    const stub = splitPackage.stubs.get(part);
    if (!stub) {
      located(base.requireText(part));
      continue;
    }
    let suffix = "";
    await readWorksheetPart(stub.feed(), {
      prefix: located,
      row: (row) => located(row.text),
      text: located,
      suffix: (text) => {
        suffix = text;
      },
    });
    located(suffix);
  }
  return locations;
}

/**
 * Write one group of a table-keeping split: the workbook as it is, its values
 * converted when asked, with the table's rows narrowed to the group's, as
 * `preserveWorkbookWithFilteredExcelTable` writes it, and its pivot tables
 * removed. Returns how many pivot tables were removed.
 */
export async function writePreservedTableGroup(
  splitPackage: StreamedSplitPackage,
  resolved: ResolvedRegionSplit,
  groupIndex: number,
  values: boolean,
  target: RegionOutputTarget,
  between: () => Promise<void>,
): Promise<number> {
  const definition = resolved.definition!;
  const workbookPackage = splitPackage.base.clone();
  const converted = new Set<string>();
  if (values) {
    for (const part of workbookPackage.partNames()) {
      if (!isConvertedWorksheetPart(part)) continue;
      // Rows are converted as they are written; the text around them now.
      workbookPackage.writeText(
        part,
        removeWorksheetFormulas(workbookPackage.requireText(part), part).xml,
      );
      if (splitPackage.stubs.has(part)) converted.add(part);
    }
    convertTablesAndCalcChain(workbookPackage);
  }
  const worksheetXml = workbookPackage.readText(definition.worksheetPart);
  const tableXml = workbookPackage.readText(definition.tablePart);
  if (worksheetXml === undefined || tableXml === undefined) {
    throw new Error(
      `Excel Table "${definition.name}" is missing workbook package parts.`,
    );
  }
  const sourceRows: number[] = [];
  resolved.groupOfRow.forEach((group, offset) => {
    if (group === groupIndex) sourceRows.push(resolved.first + offset + 1);
  });
  const rewrite = tableRewrite(definition, sourceRows.length);
  workbookPackage.writeText(
    definition.tablePart,
    rewriteTablePart(tableXml, rewrite),
  );
  // Written back whole, so the part takes the edited time.
  workbookPackage.writeText(definition.worksheetPart, worksheetXml);
  const stripped = stripPivotPartsIn(workbookPackage);

  const writer = new JsZipWriter((chunk) => target.write(chunk));
  for (const part of workbookPackage.partNames()) {
    const stub = splitPackage.stubs.get(part);
    const writing = writer.part(part, workbookPackage.partDate(part));
    if (!stub) {
      await writeHeldPart(
        splitPackage,
        workbookPackage,
        part,
        writing,
        between,
      );
      continue;
    }
    const isTable = part === definition.worksheetPart;
    const convert = converted.has(part);
    if (
      !isTable &&
      !convert &&
      workbookPackage.partBytes(part) === stub.bytes
    ) {
      await stub.feed(between)((chunk) => writing.push(chunk));
      await writing.close();
      continue;
    }
    const current = workbookPackage.readText(part)!;
    const asConverted = (text: string): string =>
      convert ? removeWorksheetFormulas(text, part).xml : text;
    if (!isTable) {
      const pieces = stubSplit(current);
      writing.text(pieces.prefix);
      await readWorksheetPart(stub.feed(between), {
        prefix: () => undefined,
        suffix: () => undefined,
        row: (row) => writing.text(asConverted(row.text)),
        text: (text) => writing.text(asConverted(text)),
      });
      writing.text(pieces.suffix);
      await writing.close();
      continue;
    }
    if (resolved.gathered) {
      // Rows out of order: the rewrite sorts them, so the part is read whole.
      let whole = "";
      const decoder = new TextDecoder();
      await stub.feed(between)((chunk) => {
        whole += decoder.decode(chunk, { stream: true });
      });
      whole += decoder.decode();
      const filtered = filterWholeWorksheetRows(
        asConverted(whole),
        definition,
        sourceRows,
        false,
      );
      writing.text(filtered.worksheetXml);
      await writing.close();
      continue;
    }
    // Rows in order: each row is kept, moved or dropped as it is read, which
    // gives the rewrite's sorted rows without holding them.
    const bounds = tableSheetDataBounds(current, definition.sheet);
    const keptAt = new Map(
      sourceRows.map((row, index) => [row, index] as const),
    );
    let rowElementName: string | undefined;
    writing.text(current.slice(0, bounds.start));
    await readWorksheetPart(stub.feed(between), {
      prefix: () => undefined,
      suffix: () => undefined,
      // The rewrite keeps rows only.
      text: () => undefined,
      row: (row: StreamedRow) => {
        const text = asConverted(row.text);
        rowElementName ??= tableRowElementName(text, bounds.openingTag);
        const number = tableRowNumber(text);
        if (
          number < rewrite.firstDataRow ||
          number > rewrite.originalTableEndRow
        ) {
          writing.text(text);
          return;
        }
        const index = keptAt.get(number);
        if (index !== undefined) {
          writing.text(
            relocateTableRow(
              text,
              rewrite.firstDataRow + index,
              false,
              rowElementName,
            ),
          );
        } else if (
          definition.totalsRow &&
          number === rewrite.originalTableEndRow
        ) {
          writing.text(
            relocateTableRow(
              text,
              rewrite.newTableEndRow,
              false,
              rowElementName,
            ),
          );
        }
      },
    });
    writing.text(current.slice(bounds.end));
    await writing.close();
  }
  await writer.finish();
  return stripped?.removedPivotTables ?? 0;
}

/** A stub's text around its empty sheetData, as the model splits it. */
function stubSplit(xml: string): { prefix: string; suffix: string } {
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
 * Write every group of a region split, in group order: compact outputs in
 * batches, table-keeping outputs one at a time. Returns the pivot tables the
 * table-keeping outputs removed.
 */
export async function writeRegionGroups(
  source: RandomAccessSource,
  resolved: ResolvedRegionSplit,
  values: boolean,
  open: (index: number) => Promise<RegionOutputTarget> | RegionOutputTarget,
  between: () => Promise<void>,
  splitPackage?: StreamedSplitPackage,
  format: TableOutputFormat = XLSX_TABLE_OUTPUT,
): Promise<number> {
  // A CSV output holds the rows only, so it is always written compactly.
  if (!resolved.definition || format.kind === "csv") {
    await writeCompactGroups(resolved, open, between, format);
    return 0;
  }
  const workbookPackage = splitPackage ?? (await openRegionPackage(source));
  // The rewrite reads every row's number before it moves any row, so a row
  // without one is refused first, whatever else the table holds.
  const tableStub = workbookPackage.stubs.get(
    resolved.definition.worksheetPart,
  );
  if (tableStub && !resolved.gathered) {
    await readWorksheetPart(tableStub.feed(between), {
      prefix: () => undefined,
      suffix: () => undefined,
      text: () => undefined,
      row: (row) => {
        tableRowNumber(row.text);
      },
    });
  }
  let pivotTablesRemoved = 0;
  for (const [index] of resolved.groups.entries()) {
    await between();
    const target = await open(index);
    try {
      pivotTablesRemoved += await writePreservedTableGroup(
        workbookPackage,
        resolved,
        index,
        values,
        target,
        between,
      );
    } catch (error) {
      await Promise.resolve(target.abort?.()).catch(() => undefined);
      throw error;
    }
    await target.close();
  }
  return pivotTablesRemoved;
}
