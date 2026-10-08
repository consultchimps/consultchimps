/**
 * L3 operation: workbook inspection. ADR 0002 makes this a first-class
 * operation on every surface, with `pptx.inspect-template` as the
 * create-nothing precedent.
 *
 * The description answers "what is in this file, and what would an operation
 * see in it". So it reads cells through the streaming reader the table readers
 * and consolidation read through (ADR 0006), and settles the header row and
 * the columns by the shared rule in `src/region/header-detection.ts`: the
 * header row an inspection reports is the row a consolidation reads from, and
 * the columns are the ones it keeps.
 *
 * Each worksheet is read once, as a stream, and nothing proportional to its
 * rows is kept: per row, whether it holds a value; per column, the last row
 * holding one and a bounded sample; and the rows the header rule may still
 * need, which on an ordinary worksheet is a dozen. Samples are bounded because
 * a description is what a picker renders, never a copy of the data; unbounded
 * values and inferred types are excluded by the ADR deliberately.
 */
import {
  ConsultChimpsError,
  throwIfAborted,
  type AbortOutputContext,
  type OperationControlOptions,
  type OperationResult,
} from "@consultchimps/core";
import { type CellValue, uniqueHeaders } from "@consultchimps/tabular";

import { XLSX_ERRORS } from "../errors.js";
import { decodeRange } from "../model/references.js";
import type { CellRange } from "../model/types.js";
import { CellError } from "../package/cell-error.js";
import {
  cellKey,
  detectHeaderRow,
  isBlankValue,
  settleHeaderCandidate,
  type RowValueCount,
} from "../region/header-detection.js";
import { formatCellRef, parseSheetRange } from "../region/values.js";
import {
  INSPECT_OPERATION,
  yieldToEventLoop,
  type ReadWorkbookOptions,
} from "../shared.js";
import { uncachedFormulaWarnings } from "../uncached-formulas.js";
import type {
  StreamedCell,
  StreamedDefinedName,
  StreamedSheet,
  StreamedValue,
  StreamedWorkbook,
  WorksheetConsumer,
  WorksheetRead,
} from "./consolidate/reader.js";
import { tableValue } from "./sheet-grid.js";

/**
 * The hard ceiling on per-column sample values. ADR 0002 requires samples to
 * be bounded: a description is a summary a picker renders, never a copy of the
 * data. A caller may ask for fewer, never for more.
 */
export const MAX_COLUMN_SAMPLE_VALUES = 5;

/** Excel's own reserved defined names (print areas and the like). */
const BUILTIN_DEFINED_NAME_PREFIX = "_xlnm.";

export type DescribeWorkbookMetric =
  | "dataRows"
  | "excelTables"
  | "formulaCellsWithoutCachedValues"
  | "headerColumns"
  | "hiddenWorksheets"
  | "namedRanges"
  | "worksheets";

/**
 * How Excel presents a worksheet. "very-hidden" is the state only the VBA
 * editor can reverse, which is why an inspection distinguishes it from an
 * ordinary hidden sheet a reader can unhide from the tab bar.
 */
export type WorksheetVisibility = "visible" | "hidden" | "very-hidden";

export interface DescribeWorkbookOptions
  extends ReadWorkbookOptions, OperationControlOptions {
  /**
   * Distinct sample values to collect per column, from 0 to
   * `MAX_COLUMN_SAMPLE_VALUES`. Defaults to the maximum.
   */
  sampleValues?: number | undefined;
}

/** One column of a worksheet's effective header row. */
export interface WorkbookColumnDescription {
  /** The header as the union readers would spell it, blanks filled in. */
  header: string;
  /** Zero-based position within the header row. */
  index: number;
  /**
   * The first few distinct non-empty stored values below the header, in the
   * order the rows carry them. Bounded by the sample limit; never the whole
   * column, and never a type this package inferred.
   */
  sampleValues: CellValue[];
}

export interface WorkbookSheetDescription {
  name: string;
  visibility: WorksheetVisibility;
  /** Rows in the worksheet's used range, header row included; 0 when empty. */
  rowCount: number;
  /** Columns in the worksheet's used range; 0 when empty. */
  columnCount: number;
  /**
   * The one-based effective header row - the row `resolveRegions` resolves for
   * this worksheet - or undefined when the worksheet holds no values for one
   * to be found in.
   */
  headerRow: number | undefined;
  /** The header preview: one entry per column of the effective header row. */
  columns: WorkbookColumnDescription[];
  /** Non-empty rows below the header row. */
  dataRowCount: number;
}

export interface WorkbookExcelTableDescription {
  name: string;
  range: string;
  sheet: string;
  /** The column names the table part declares, in table order. */
  headers: string[];
}

export interface WorkbookNamedRangeDescription {
  name: string;
  /** The range this name points at, as the workbook stores it. */
  ref: string;
  sheet: string;
}

export interface WorkbookDescription {
  /** The workbook this description was read from: a filename or input name. */
  source: string;
  sheets: WorkbookSheetDescription[];
  excelTables: WorkbookExcelTableDescription[];
  namedRanges: WorkbookNamedRangeDescription[];
}

/**
 * The outcome of a workbook inspection: the structured operation result every
 * completed operation reports, plus the description it summarizes. The two
 * travel side by side for the same reason `ByteOperationOutcome` keeps
 * `outputs` beside `result`: metrics are counts, and sheet names, headers and
 * sample values are not.
 */
export interface WorkbookDescriptionOutcome {
  description: WorkbookDescription;
  result: OperationResult<DescribeWorkbookMetric>;
}

/**
 * Validate the sample bound. Silently clamping a caller's 50 to 5 would make
 * the same call mean different things on either side of the cap, so an
 * out-of-range request is a stable refusal instead.
 */
function resolveSampleLimit(requested: number | undefined): number {
  if (requested === undefined) {
    return MAX_COLUMN_SAMPLE_VALUES;
  }
  if (
    !Number.isInteger(requested) ||
    requested < 0 ||
    requested > MAX_COLUMN_SAMPLE_VALUES
  ) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_SAMPLE_LIMIT,
      `The sample value count must be a whole number from 0 to ${MAX_COLUMN_SAMPLE_VALUES}.`,
      {
        details: {
          maximum: MAX_COLUMN_SAMPLE_VALUES,
          sampleValues: requested,
        },
      },
    );
  }
  return requested;
}

/**
 * Refuse an invalid `headerRow` before anything reads the workbook.
 *
 * The region resolver validates this too, but it only ever sees the option for
 * a worksheet that has content to resolve. Leaving the check to it made the
 * same option valid or invalid depending on what the workbook happened to
 * contain - an empty worksheet, or a selection that filtered every sheet out
 * as hidden, would accept `0` or `1.5` in silence. Option validation belongs
 * to the operation, before the work starts; the code and wording match the
 * resolver's so a caller sees one refusal however the sheet is shaped.
 */
function validateHeaderRow(headerRow: number | undefined): void {
  if (
    headerRow !== undefined &&
    (!Number.isInteger(headerRow) || headerRow < 1)
  ) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_HEADER_ROW,
      "The header row must be a positive whole number counted from 1.",
      { details: { headerRow } },
    );
  }
}

/**
 * A sample value's identity for the distinctness test. The type is part of the
 * key so the number 1 and the text "1", which Excel very much distinguishes
 * and which a mapping review needs to see separately, do not collapse into
 * one sample.
 */
function sampleKey(value: CellValue): string {
  return `${typeof value}:${String(value)}`;
}

function isEmptyValue(value: CellValue): boolean {
  return value === null || value === "";
}

/** An A1 range, written the way the workbook itself writes it. */
function formatRange(range: CellRange): string {
  const start = formatCellRef(range.start);
  const end = formatCellRef(range.end);
  return start === end ? start : `${start}:${end}`;
}

const EMPTY_SHEET = {
  columnCount: 0,
  columns: [] as WorkbookColumnDescription[],
  dataRowCount: 0,
  headerRow: undefined,
  rowCount: 0,
} as const;

/** The public spelling of how Excel presents a worksheet. */
function publicVisibility(
  visibility: StreamedSheet["visibility"],
): WorksheetVisibility {
  return visibility === "veryHidden" ? "very-hidden" : visibility;
}

/**
 * A cell's value exactly as the workbook stores it. ADR 0002 promises samples
 * are stored values with no inferred types: a date-formatted cell stores the
 * serial, not a date, and a declared date its own text, which the reader
 * carries as `stored`. An error cell stores its text.
 */
function storedValue(cell: StreamedCell): CellValue {
  return tableValue(cell.stored ?? cell.value);
}

/**
 * The name a header cell gives its column, spelled as the table readers spell
 * it: a number by its value, a boolean as `true` or `false`, a date as its ISO
 * timestamp, an error by its text, and a blank cell as an empty name for
 * `uniqueHeaders` to fill in.
 */
function headerName(value: StreamedValue | undefined): string {
  if (value === undefined || isBlankValue(value)) return "";
  return value instanceof CellError ? value.text : String(value);
}

/** An Excel Table's range, one-based, as the workbook declares it. */
function tableRange(reference: string): CellRange {
  const range = decodeRange(reference);
  return {
    start: { row: range.startRow + 1, column: range.startColumn },
    end: { row: range.endRow + 1, column: range.endColumn },
  };
}

/**
 * The first few distinct non-empty stored values of each column, in the order
 * they are offered.
 */
class ColumnSamples {
  readonly #limit: number;
  readonly #values = new Map<number, CellValue[]>();
  readonly #seen = new Map<number, Set<string>>();

  constructor(limit: number) {
    this.#limit = limit;
  }

  offer(column: number, value: CellValue): void {
    if (isEmptyValue(value)) return;
    let values = this.#values.get(column);
    if (values === undefined) {
      values = [];
      this.#values.set(column, values);
      this.#seen.set(column, new Set());
    }
    if (values.length >= this.#limit) return;
    const key = sampleKey(value);
    const seen = this.#seen.get(column)!;
    if (seen.has(key)) return;
    seen.add(key);
    values.push(value);
  }

  /** This column's samples, then `later`'s that are new, to the limit. */
  merged(column: number, later: ColumnSamples): CellValue[] {
    const values = [...(this.#values.get(column) ?? [])];
    const seen = new Set(values.map(sampleKey));
    for (const value of later.#values.get(column) ?? []) {
      if (values.length >= this.#limit) break;
      const key = sampleKey(value);
      if (seen.has(key)) continue;
      seen.add(key);
      values.push(value);
    }
    return values;
  }
}

/** A row the header rule may still need. */
interface HeldRow {
  readonly row: number;
  readonly cells: readonly StreamedCell[];
  /** The row's position in `counts`. */
  readonly count: number;
}

/**
 * One pass over one worksheet, zero-based throughout.
 *
 * Rows holding a value are counted for the header rule and held until the
 * rule's answer can no longer change (`settleHeaderCandidate`); the header row
 * is then the candidate or the first populated row, so every row after that
 * point lies below it either way, and only feeds the counts and samples. A
 * row of cells that hold something but no value, such as formulas with no
 * cached value, is remembered by its columns, since whether it is a data row
 * depends on which columns turn out to be kept.
 */
class SheetScan implements WorksheetConsumer {
  readonly #declared: number | undefined;
  readonly #limit: number;
  #counts: RowValueCount[] = [];
  #held: HeldRow[] = [];
  #candidate = 0;
  #settled = false;
  #declaredCells: readonly StreamedCell[] | undefined;
  #lastValueRow = new Map<number, number>();
  #later: ColumnSamples;
  #rowsLater = 0;
  #valuelessRows: Array<{ row: number; columns: readonly number[] }> = [];

  /** `headerRow` is the one-based row a caller declared, if any. */
  constructor(headerRow: number | undefined, sampleLimit: number) {
    this.#declared = headerRow === undefined ? undefined : headerRow - 1;
    this.#limit = sampleLimit;
    this.#later = new ColumnSamples(sampleLimit);
  }

  begin(): void {
    this.#counts = [];
    this.#held = [];
    this.#candidate = 0;
    this.#settled = false;
    this.#declaredCells = undefined;
    this.#lastValueRow = new Map();
    this.#later = new ColumnSamples(this.#limit);
    this.#rowsLater = 0;
    this.#valuelessRows = [];
  }

  row(
    row: number,
    cells: readonly StreamedCell[],
    occupied: readonly number[] = [],
  ): void {
    let values = 0;
    for (const cell of cells) {
      if (!isBlankValue(cell.value)) {
        values += 1;
        this.#lastValueRow.set(cell.column, row);
      }
    }
    if (values === 0) {
      if (occupied.length > 0) {
        this.#valuelessRows.push({ row, columns: occupied });
      }
      return;
    }
    if (this.#declared !== undefined) {
      if (row === this.#declared) this.#declaredCells = cells;
      else if (row > this.#declared) this.#feedLater(cells);
      return;
    }
    if (this.#settled) {
      this.#feedLater(cells);
      return;
    }
    this.#counts.push({ row, values, bannerValues: 0 });
    this.#held.push({ row, cells, count: this.#counts.length - 1 });
    const { index, settled } = settleHeaderCandidate(
      this.#counts,
      this.#candidate,
    );
    this.#candidate = index;
    this.#settled = settled;
  }

  #feedLater(cells: readonly StreamedCell[]): void {
    this.#rowsLater += 1;
    for (const cell of cells) {
      this.#later.offer(cell.column, storedValue(cell));
    }
  }

  /** The description, once the read is complete. */
  finish(
    read: WorksheetRead,
    name: string,
    visibility: WorksheetVisibility,
  ): WorkbookSheetDescription {
    const range = read.range;
    if (range === undefined || read.occupied !== true) {
      return { ...EMPTY_SHEET, columns: [], name, visibility };
    }
    const rowCount = range.endRow - range.startRow + 1;
    const columnCount = range.endColumn - range.startColumn + 1;
    const declared = this.#declared;
    if (
      declared !== undefined &&
      (declared < range.startRow || declared > range.endRow)
    ) {
      // A declared header row outside the used range leaves nothing to
      // preview, as the table readers yield no table for it.
      return {
        ...EMPTY_SHEET,
        columnCount,
        columns: [],
        name,
        rowCount,
        visibility,
      };
    }

    // Merged ranges follow the cells in the part, so the held rows are only
    // now measured for banners; the rule reads no other row's.
    const banners = new Set<string>();
    for (const merge of read.merges) {
      if (merge.endColumn > merge.startColumn) {
        banners.add(cellKey(merge.startRow, merge.startColumn));
      }
    }
    if (banners.size > 0) {
      for (const held of this.#held) {
        let bannerValues = 0;
        for (const cell of held.cells) {
          if (
            !isBlankValue(cell.value) &&
            banners.has(cellKey(held.row, cell.column))
          ) {
            bannerValues += 1;
          }
        }
        this.#counts[held.count] = {
          ...this.#counts[held.count]!,
          bannerValues,
        };
      }
    }

    const detected = declared ?? detectHeaderRow(this.#counts);
    // A worksheet with no value in any row, one of uncalculated formulas say,
    // has no row to pick and no column to leave out: it reads from its first
    // used row, across every used column.
    const header = detected ?? range.startRow;
    const columns: number[] = [];
    for (
      let column = range.startColumn;
      column <= range.endColumn;
      column += 1
    ) {
      const last = this.#lastValueRow.get(column);
      if (detected === undefined || (last !== undefined && last >= header)) {
        columns.push(column);
      }
    }
    const headerCells =
      declared === undefined
        ? this.#held.find((held) => held.row === header)?.cells
        : this.#declaredCells;
    const headerValues = new Map<number, StreamedValue>();
    for (const cell of headerCells ?? []) {
      headerValues.set(cell.column, cell.value);
    }

    // Every row holding a value below the header holds it in a kept column,
    // since a column with a value below the header is never a spacer.
    const kept = new Set(columns);
    const earlier = new ColumnSamples(this.#limit);
    let dataRowCount = this.#rowsLater;
    for (const held of this.#held) {
      if (held.row <= header) continue;
      dataRowCount += 1;
      for (const cell of held.cells) {
        earlier.offer(cell.column, storedValue(cell));
      }
    }
    for (const valueless of this.#valuelessRows) {
      if (
        valueless.row > header &&
        valueless.columns.some((column) => kept.has(column))
      ) {
        dataRowCount += 1;
      }
    }

    const headers = uniqueHeaders(
      columns.map((column) => headerName(headerValues.get(column)) || null),
    );
    return {
      columnCount,
      columns: headers.map((headerText, index) => ({
        header: headerText,
        index,
        sampleValues: earlier.merged(columns[index]!, this.#later),
      })),
      dataRowCount,
      headerRow: header + 1,
      name,
      rowCount,
      visibility,
    };
  }
}

/**
 * Describe one worksheet in one streaming pass. The read hands back control
 * between chunks of the part, so a cancellation posted while a long worksheet
 * is read is collected rather than observed after the answer is built.
 */
async function describeWorksheet(
  workbook: StreamedWorkbook,
  sheet: StreamedSheet,
  options: DescribeWorkbookOptions,
  sampleLimit: number,
  outputContext: AbortOutputContext,
  uncachedFormulas: string[],
): Promise<WorkbookSheetDescription> {
  const scan = new SheetScan(options.headerRow, sampleLimit);
  const read = await workbook.readWorksheet(sheet, scan, {
    occupancy: true,
    between: async () => {
      await yieldToEventLoop();
      throwIfAborted(options.signal, INSPECT_OPERATION, outputContext);
    },
  });
  // Every one on the sheet: none gives a sample, and one standing alone in a
  // column is the reason that column shows no values.
  if (read.range !== undefined) {
    for (const position of read.uncachedFormulas) {
      uncachedFormulas.push(
        `${sheet.name}!${formatCellRef({
          column: position.column,
          row: position.row + 1,
        })}`,
      );
    }
  }
  return scan.finish(read, sheet.name, publicVisibility(sheet.visibility));
}

function describeExcelTables(
  workbook: StreamedWorkbook,
  describedSheets: ReadonlySet<string>,
): WorkbookExcelTableDescription[] {
  return workbook.tables
    .filter((table) => describedSheets.has(table.sheet))
    .map((table) => ({
      name: table.name,
      range: formatRange(tableRange(table.range)),
      sheet: table.sheet,
      // The declared column names, not the cells: a table with no data rows
      // still has headers, and a picker needs to offer them.
      headers: [...table.columns],
    }));
}

function describeNamedRanges(
  definedNames: readonly StreamedDefinedName[],
  describedSheets: ReadonlySet<string>,
): WorkbookNamedRangeDescription[] {
  const ranges: WorkbookNamedRangeDescription[] = [];

  for (const definedName of definedNames) {
    if (definedName.name.startsWith(BUILTIN_DEFINED_NAME_PREFIX)) {
      continue;
    }
    // The reference arrives with its entities resolved, so a worksheet called
    // `Review & Log` matches the decoded sheet names compared against below.
    const parsed = parseSheetRange(definedName.reference);
    if (!parsed || !describedSheets.has(parsed.sheet)) {
      continue;
    }
    ranges.push({
      name: definedName.name,
      ref: formatRange(parsed.range),
      sheet: parsed.sheet,
    });
  }

  return ranges;
}

/**
 * The worksheets this description covers, in workbook order: the sheets whose
 * cells the workbook says where to find, so not a chart sheet.
 *
 * `includeHiddenSheets` and `sheets` mean exactly what they mean to every
 * other reader in this package: an option that reads the same everywhere is
 * worth more than an inspection-specific default. Naming a worksheet the
 * workbook does not have is a refusal rather than an empty answer, because a
 * picker asking about a sheet that is not there has a mistake to report.
 */
function selectSheets(
  workbook: StreamedWorkbook,
  options: DescribeWorkbookOptions,
): { hiddenExcluded: number; selected: StreamedSheet[] } {
  const worksheets = workbook.sheets.filter(
    (sheet) => sheet.part !== undefined && sheet.part !== "",
  );
  if (worksheets.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_SHEETS,
      "The workbook does not contain any worksheets.",
      { details: { availableWorksheets: [] } },
    );
  }

  const requested = options.sheets;
  if (requested) {
    const available = new Set(
      worksheets.map((sheet) => sheet.name.toLocaleLowerCase()),
    );
    const missing = requested.filter(
      (name) => !available.has(name.toLocaleLowerCase()),
    );
    if (missing.length > 0) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_WORKSHEET_NOT_FOUND,
        `Worksheet "${missing[0]}" was not found in the workbook.`,
        {
          details: {
            availableWorksheets: worksheets.map((sheet) => sheet.name),
            missingWorksheets: missing,
          },
        },
      );
    }
  }
  const selectedNames = requested
    ? new Set(requested.map((name) => name.toLocaleLowerCase()))
    : undefined;

  const selected: StreamedSheet[] = [];
  let hiddenExcluded = 0;

  for (const sheet of worksheets) {
    if (selectedNames && !selectedNames.has(sheet.name.toLocaleLowerCase())) {
      continue;
    }
    if (options.includeHiddenSheets !== true && !sheet.visible) {
      hiddenExcluded += 1;
      continue;
    }
    selected.push(sheet);
  }

  return { hiddenExcluded, selected };
}

/**
 * The conditions an inspection reports as warnings. Order is fixed rather than
 * derived from the workbook, so identical inputs produce an identical result.
 */
function descriptionWarnings(
  description: WorkbookDescription,
  hiddenExcluded: number,
): string[] {
  const warnings: string[] = [];

  if (hiddenExcluded > 0) {
    warnings.push(
      `${hiddenExcluded} worksheet${
        hiddenExcluded === 1 ? " is" : "s are"
      } hidden and ${
        hiddenExcluded === 1 ? "was" : "were"
      } not described. Include hidden worksheets to describe ${
        hiddenExcluded === 1 ? "it" : "them"
      }.`,
    );
  }

  const withoutHeaderRow = description.sheets
    .filter((sheet) => sheet.headerRow === undefined)
    .map((sheet) => sheet.name);
  if (withoutHeaderRow.length > 0) {
    warnings.push(
      `No header row was found in ${withoutHeaderRow
        .map((name) => `"${name}"`)
        .join(
          ", ",
        )}. An operation that matches columns by header would find nothing to match in ${
        withoutHeaderRow.length === 1 ? "it" : "them"
      }.`,
    );
  }

  if (description.sheets.length === 0) {
    warnings.push(
      "No worksheets matched the selection, so the description is empty.",
    );
  }

  return warnings;
}

/** Present a description as the structured result every operation reports. */
export function workbookDescriptionResult(
  description: WorkbookDescription,
  hiddenExcluded: number,
  uncachedFormulas: readonly string[] = [],
): OperationResult<DescribeWorkbookMetric> {
  return {
    operation: INSPECT_OPERATION,
    // An inspection creates nothing, so it has no artifacts. The structure
    // itself travels beside this result: metrics are counts, and names are not.
    artifacts: [],
    warnings: [
      ...descriptionWarnings(description, hiddenExcluded),
      ...uncachedFormulaWarnings(
        uncachedFormulas,
        "they give no sample values",
      ),
    ],
    metrics: {
      dataRows: description.sheets.reduce(
        (total, sheet) => total + sheet.dataRowCount,
        0,
      ),
      excelTables: description.excelTables.length,
      formulaCellsWithoutCachedValues: uncachedFormulas.length,
      headerColumns: description.sheets.reduce(
        (total, sheet) => total + sheet.columns.length,
        0,
      ),
      hiddenWorksheets: description.sheets.filter(
        (sheet) => sheet.visibility !== "visible",
      ).length,
      namedRanges: description.namedRanges.length,
      worksheets: description.sheets.length,
    },
  };
}

/**
 * Describe a workbook opened for streaming. Both surfaces call this, the
 * command line over a file and the byte surface over bytes or a browser
 * `File`, so one workbook gives one description however it is read.
 *
 * Every worksheet boundary carries an abort check and a yield, and each read
 * yields between chunks, so a cancellation posted while the operation runs is
 * actually collected rather than observed after the answer is already built.
 */
export async function describeStreamedWorkbook(
  workbook: StreamedWorkbook,
  source: string,
  options: DescribeWorkbookOptions = {},
  outputContext: AbortOutputContext = "files",
): Promise<WorkbookDescriptionOutcome> {
  throwIfAborted(options.signal, INSPECT_OPERATION, outputContext);
  // Options are validated before any worksheet is read, so a bad option is
  // refused the same way whatever the workbook contains.
  const sampleLimit = resolveSampleLimit(options.sampleValues);
  validateHeaderRow(options.headerRow);
  const { hiddenExcluded, selected } = selectSheets(workbook, options);

  const sheets: WorkbookSheetDescription[] = [];
  const uncachedFormulas: string[] = [];
  try {
    for (const [index, sheet] of selected.entries()) {
      if (index > 0) {
        await yieldToEventLoop();
      }
      throwIfAborted(options.signal, INSPECT_OPERATION, outputContext);
      sheets.push(
        await describeWorksheet(
          workbook,
          sheet,
          options,
          sampleLimit,
          outputContext,
          uncachedFormulas,
        ),
      );
      options.onProgress?.({
        operation: INSPECT_OPERATION,
        stage: "describing-worksheets",
        completed: index + 1,
        total: selected.length,
        detail: sheet.name,
      });
    }
  } finally {
    workbook.releaseStrings();
  }

  throwIfAborted(options.signal, INSPECT_OPERATION, outputContext);
  const describedSheets = new Set(selected.map((sheet) => sheet.name));
  const description: WorkbookDescription = {
    source,
    sheets,
    excelTables: describeExcelTables(workbook, describedSheets),
    namedRanges: describeNamedRanges(workbook.names, describedSheets),
  };

  throwIfAborted(options.signal, INSPECT_OPERATION, outputContext);
  options.onProgress?.({
    operation: INSPECT_OPERATION,
    stage: "describing-structures",
    completed: 1,
    total: 1,
  });

  return {
    description,
    result: workbookDescriptionResult(
      description,
      hiddenExcluded,
      uncachedFormulas,
    ),
  };
}
