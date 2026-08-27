/**
 * L3 operation — workbook inspection. ADR 0002 makes this a first-class
 * operation on every surface, with `pptx.inspect-template` as the
 * create-nothing precedent.
 *
 * The description answers "what is in this file, and what would an operation
 * see in it". That second half is why this is composed from the L1 model and
 * the L2 region resolver rather than read off SheetJS objects: the header row
 * an inspection reports is the row `resolveRegions` would resolve, the columns
 * are the region's boundary columns, and the visibility is the model's own
 * `SheetInfo`. A second implementation of header detection here would let the
 * promise drift the moment a model-layer fix lands — which is exactly the
 * drift ARCHITECTURE.md's layering exists to prevent.
 *
 * Samples are bounded because a description is what a picker renders, never a
 * copy of the data; unbounded values and inferred types are excluded by the
 * ADR deliberately.
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
import { WorkbookModel } from "../model/index.js";
import type {
  CellModel,
  CellRange,
  CellRef,
  DefinedNameEntry,
  RowModel,
  SheetInfo,
  WorkbookTableInfo,
  WorksheetModel,
} from "../model/types.js";
import { decodeXmlText } from "../model/xml.js";
import { resolveRegions } from "../region/resolve.js";
import type { DataRegion } from "../region/types.js";
import { formatCellRef, parseSheetRange } from "../region/values.js";
import {
  INSPECT_OPERATION,
  yieldToEventLoop,
  type ReadWorkbookOptions,
} from "../shared.js";

/**
 * The hard ceiling on per-column sample values. ADR 0002 requires samples to
 * be bounded: a description is a summary a picker renders, never a copy of the
 * data. A caller may ask for fewer, never for more.
 */
export const MAX_COLUMN_SAMPLE_VALUES = 5;

/**
 * Rows scanned between yields. A worksheet scan is synchronous work, so
 * without a macrotask boundary a cancellation posted to a worker cannot be
 * dequeued until the whole sheet is done - the same reason `consolidate`
 * yields between inputs. See `yieldToEventLoop`.
 */
const ROWS_PER_YIELD = 1024;

/** Excel's own reserved defined names (print areas and the like). */
const BUILTIN_DEFINED_NAME_PREFIX = "_xlnm.";

export type DescribeWorkbookMetric =
  | "dataRows"
  | "excelTables"
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
 * `outputs` beside `result` — metrics are counts, and sheet names, headers and
 * sample values are not.
 */
export interface WorkbookDescriptionOutcome {
  description: WorkbookDescription;
  result: OperationResult<DescribeWorkbookMetric>;
}

/**
 * Load the document model, reporting an unreadable package as the stable read
 * error every other workbook reader raises.
 */
export async function loadWorkbookModelForDescribe(
  bytes: Uint8Array,
  source: string,
  details: Record<string, unknown>,
): Promise<WorkbookModel> {
  try {
    return await WorkbookModel.load(bytes);
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${source}`,
      { cause: error, details },
    );
  }
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

/** The model's visibility vocabulary in this operation's public spelling. */
function publicVisibility(
  visibility: SheetInfo["visibility"],
): WorksheetVisibility {
  return visibility === "veryHidden" ? "very-hidden" : visibility;
}

/**
 * A cell's value exactly as the workbook stores it.
 *
 * `cellValue` is the grouping-oriented view: it reads the cell's style and
 * infers a `Date` from a date number format, which is the right answer for a
 * split key and the wrong one here. ADR 0002 promises samples are stored
 * values with no inferred types, and a date-formatted cell stores a *number* -
 * the serial - not a date. Reporting an ISO string would show a mapping review
 * a type and a value the cell does not contain.
 *
 * So a `Date` from the model is re-read from the cell's stored text: a numeric
 * serial comes back as that number, and a genuine ISO date cell (`t="d"`,
 * whose stored text is the ISO string) comes back as that text. Both are what
 * the workbook holds, and both are deterministic.
 */
function storedValue(worksheet: WorksheetModel, ref: CellRef): CellValue {
  const value = worksheet.cellValue(ref);
  if (value === undefined || value === null) {
    return null;
  }
  if (value instanceof Date) {
    const text = worksheet.cellText(ref);
    if (text === undefined || text.trim() === "") {
      return null;
    }
    const numeric = Number(text.trim());
    return Number.isFinite(numeric) ? numeric : text;
  }
  if (
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return value;
  }
  return String(value);
}

/**
 * A sample value's identity for the distinctness test. The type is part of the
 * key so the number 1 and the text "1" — which Excel very much distinguishes,
 * and which a mapping review needs to see separately — do not collapse into
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

/**
 * Whether the worksheet carries any value at all inside its used range.
 *
 * A worksheet's used range comes from the stored `<dimension>` hint, which
 * Excel and every writer set to `A1` for a sheet that holds nothing — so the
 * range alone cannot tell "one blank cell" from "no cells". An inspection has
 * to answer "there is no header row here" for a genuinely empty sheet rather
 * than invent a `column_1`, so emptiness is decided by looking.
 *
 * It looks at the cells the sheet actually stores rather than at every
 * coordinate the dimension claims. A blank formatted template can declare an
 * extent spanning hundreds of thousands of rows while storing nothing at all;
 * walking the declared grid made answering "empty" proportional to the claim
 * instead of to the contents.
 *
 * A template can still *store* a great many rows - one styled cell per row is
 * enough - and finding the first occupied cell among them is a long scan when
 * the answer is "none". So it yields on the same cadence as every other scan
 * here rather than running to the end uninterrupted.
 */
async function hasAnyContent(
  rows: readonly RowModel[],
  options: DescribeWorkbookOptions,
  outputContext: AbortOutputContext,
): Promise<boolean> {
  let rowsSinceYield = 0;
  for (const row of rows) {
    if (row.cells.some(isOccupiedCell)) {
      return true;
    }
    rowsSinceYield += 1;
    if (rowsSinceYield >= ROWS_PER_YIELD) {
      rowsSinceYield = 0;
      await yieldToEventLoop();
      throwIfAborted(options.signal, INSPECT_OPERATION, outputContext);
    }
  }
  return false;
}

/**
 * Whether a cell counts as content.
 *
 * A stored value obviously counts. So does a formula with no cached value: the
 * cell is not blank, the worksheet holds it, and reporting the row as absent
 * would tell a reader the sheet is emptier than it is. Sampling stays separate
 * and stays stored-values-only - an uncached formula contributes no sample,
 * because there is no stored value to report and the inspection never computes
 * one - so occupancy and samples answer their own questions.
 *
 * A cell carrying only a style is formatting, not content, and does not count.
 */
function isOccupiedCell(cell: CellModel): boolean {
  return (
    (cell.value !== undefined && cell.value !== "") ||
    cell.formula !== undefined
  );
}

const EMPTY_SHEET = {
  columnCount: 0,
  columns: [] as WorkbookColumnDescription[],
  dataRowCount: 0,
  headerRow: undefined,
  rowCount: 0,
} as const;

/**
 * Describe one worksheet from its resolved region: the dimensions of its used
 * range, the header row an operation would key on, and a bounded sample of
 * each column's values.
 *
 * The scan walks the rows the sheet stores, in order, and stops sampling as
 * soon as every column has filled its quota, but keeps counting occupied rows
 * to the end. Because the order is row order and the bound is fixed, the same
 * worksheet always yields the same samples; the periodic yield changes when the
 * scan runs, never what it produces.
 *
 * Rows are read from the model's row store once and looked up by number, so
 * the cost is proportional to what the sheet contains rather than to the extent
 * it declares - and never to the square of it.
 */
async function describeWorksheet(
  workbook: WorkbookModel,
  sheet: SheetInfo,
  worksheet: WorksheetModel,
  options: DescribeWorkbookOptions,
  sampleLimit: number,
  outputContext: AbortOutputContext,
): Promise<WorkbookSheetDescription> {
  const visibility = publicVisibility(sheet.visibility);
  const used = worksheet.usedRange;
  if (!used) {
    return { ...EMPTY_SHEET, columns: [], name: sheet.name, visibility };
  }

  // Materializing the row store is one synchronous burst - the model parses
  // and allocates every row and cell in a single call - so the signal is
  // checked on a fresh macrotask immediately before it, where a cancellation
  // queued while the previous worksheet was scanned can still be collected
  // without paying for this sheet at all.
  await yieldToEventLoop();
  throwIfAborted(options.signal, INSPECT_OPERATION, outputContext);
  // One pass over the stored rows answers both "is this sheet empty" and every
  // per-row occupancy question below.
  const storedRows = worksheet.rows();
  throwIfAborted(options.signal, INSPECT_OPERATION, outputContext);

  if (!(await hasAnyContent(storedRows, options, outputContext))) {
    return { ...EMPTY_SHEET, columns: [], name: sheet.name, visibility };
  }

  const rowCount = used.end.row - used.start.row + 1;
  const columnCount = used.end.column - used.start.column + 1;

  // The region resolver owns header detection and the headerRow override, so
  // the row reported here is the row an operation on this sheet would use.
  const regions = await resolveRegions(workbook, {
    headerRow: options.headerRow,
    sheet: sheet.name,
  });
  const region: DataRegion | undefined = regions[0];
  if (!region || region.headerRow > used.end.row) {
    // A declared header row below the last used row leaves nothing to preview.
    return {
      ...EMPTY_SHEET,
      columnCount,
      columns: [],
      name: sheet.name,
      rowCount,
      visibility,
    };
  }

  // Blank and repeated header cells are filled in and de-duplicated here
  // rather than in the region layer, which deliberately keeps boundary columns
  // raw: naming is a `tabular` concern, and this is where the description
  // promises the names an operation would actually produce.
  const headers = uniqueHeaders(
    region.columns.map((column) => column.name || null),
  );
  const columnIndexes = region.columns.map((column) => column.index);
  const regionColumns = new Set(columnIndexes);

  const samples = headers.map(() => [] as CellValue[]);
  const seen = headers.map(() => new Set<string>());
  let dataRowCount = 0;
  let satisfiedColumns = sampleLimit === 0 ? headers.length : 0;
  let rowsSinceYield = 0;

  // Only the rows the sheet stores, in document order, clipped to the body.
  // A declared extent far larger than the contents costs nothing here.
  const bodyRows = storedRows.filter(
    (row) =>
      row.number >= region.body.start.row && row.number <= region.body.end.row,
  );

  for (const row of bodyRows) {
    rowsSinceYield += 1;
    if (rowsSinceYield >= ROWS_PER_YIELD) {
      rowsSinceYield = 0;
      // A long scan must hand back a macrotask, or a cancellation posted while
      // it runs cannot be dequeued until the sheet is finished.
      await yieldToEventLoop();
      throwIfAborted(options.signal, INSPECT_OPERATION, outputContext);
    }

    // Occupancy comes from the cells the row holds - a formula with no cached
    // value still occupies its row - while the samples below stay
    // stored-values-only.
    if (
      !row.cells.some(
        (cell) => regionColumns.has(cell.ref.column) && isOccupiedCell(cell),
      )
    ) {
      continue;
    }
    dataRowCount += 1;

    if (satisfiedColumns >= headers.length) {
      continue;
    }

    const values = columnIndexes.map((column) =>
      storedValue(region.worksheet, { column, row: row.number }),
    );

    values.forEach((value, index) => {
      const columnSamples = samples[index]!;
      if (isEmptyValue(value) || columnSamples.length >= sampleLimit) {
        return;
      }
      const key = sampleKey(value);
      const columnSeen = seen[index]!;
      if (columnSeen.has(key)) {
        return;
      }
      columnSeen.add(key);
      columnSamples.push(value);
      if (columnSamples.length === sampleLimit) {
        satisfiedColumns += 1;
      }
    });
  }

  return {
    columnCount,
    columns: headers.map((header, index) => ({
      header,
      index,
      sampleValues: samples[index]!,
    })),
    dataRowCount,
    headerRow: region.headerRow,
    name: sheet.name,
    rowCount,
    visibility,
  };
}

function describeExcelTables(
  tables: readonly WorkbookTableInfo[],
  describedSheets: ReadonlySet<string>,
): WorkbookExcelTableDescription[] {
  return tables
    .filter((table) => describedSheets.has(table.sheetName))
    .map((table) => ({
      name: table.name,
      range: formatRange(table.range),
      sheet: table.sheetName,
      // The declared column names, not the cells: a table with no data rows
      // still has headers, and a picker needs to offer them. The model reads
      // these from the table part, so an empty table is described in full.
      headers: [...table.columnNames],
    }));
}

function describeNamedRanges(
  definedNames: readonly DefinedNameEntry[],
  describedSheets: ReadonlySet<string>,
): WorkbookNamedRangeDescription[] {
  const ranges: WorkbookNamedRangeDescription[] = [];

  for (const definedName of definedNames) {
    if (definedName.name.startsWith(BUILTIN_DEFINED_NAME_PREFIX)) {
      continue;
    }
    // The model hands back the reference as the workbook part stores it, which
    // is escaped XML text: a worksheet called `Review & Log` arrives as
    // `'Review &amp; Log'!$A$1`. Sheet names elsewhere in the description are
    // decoded, so without this the membership test below would silently drop
    // the range and undercount the `namedRanges` metric.
    const parsed = parseSheetRange(decodeXmlText(definedName.reference));
    if (!parsed || !describedSheets.has(parsed.sheet)) {
      continue;
    }
    ranges.push({
      name: decodeXmlText(definedName.name),
      ref: formatRange(parsed.range),
      sheet: parsed.sheet,
    });
  }

  return ranges;
}

/**
 * The worksheets this description covers, in workbook order.
 *
 * `includeHiddenSheets` and `sheets` mean exactly what they mean to every
 * other reader in this package — an option that reads the same everywhere is
 * worth more than an inspection-specific default. Naming a worksheet the
 * workbook does not have is a refusal rather than an empty answer, because a
 * picker asking about a sheet that is not there has a mistake to report.
 */
function selectSheets(
  workbook: WorkbookModel,
  options: DescribeWorkbookOptions,
): { hiddenExcluded: number; selected: SheetInfo[] } {
  if (workbook.sheets.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_SHEETS,
      "The workbook does not contain any worksheets.",
      { details: { availableWorksheets: [] } },
    );
  }

  const requested = options.sheets;
  if (requested) {
    const available = new Set(
      workbook.sheets.map((sheet) => sheet.name.toLocaleLowerCase()),
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
            availableWorksheets: workbook.sheets.map((sheet) => sheet.name),
            missingWorksheets: missing,
          },
        },
      );
    }
  }
  const selectedNames = requested
    ? new Set(requested.map((name) => name.toLocaleLowerCase()))
    : undefined;

  const selected: SheetInfo[] = [];
  let hiddenExcluded = 0;

  for (const sheet of workbook.sheets) {
    if (selectedNames && !selectedNames.has(sheet.name.toLocaleLowerCase())) {
      continue;
    }
    if (
      options.includeHiddenSheets !== true &&
      sheet.visibility !== "visible"
    ) {
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
): OperationResult<DescribeWorkbookMetric> {
  return {
    operation: INSPECT_OPERATION,
    // An inspection creates nothing, so it has no artifacts. The structure
    // itself travels beside this result: metrics are counts, and names are not.
    artifacts: [],
    warnings: descriptionWarnings(description, hiddenExcluded),
    metrics: {
      dataRows: description.sheets.reduce(
        (total, sheet) => total + sheet.dataRowCount,
        0,
      ),
      excelTables: description.excelTables.length,
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
 * Describe a loaded workbook. Both surfaces call this with the same model, so
 * the file and byte descriptions of one workbook are structurally identical.
 *
 * Every worksheet boundary carries an abort check and a yield, and long scans
 * yield inside themselves, so a cancellation posted while the operation runs
 * is actually collected rather than observed after the answer is already built.
 */
export async function describeWorkbookModel(
  workbook: WorkbookModel,
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
  for (const [index, sheet] of selected.entries()) {
    if (index > 0) {
      await yieldToEventLoop();
    }
    throwIfAborted(options.signal, INSPECT_OPERATION, outputContext);
    const worksheet = workbook.worksheet(sheet.name);
    sheets.push(
      worksheet
        ? await describeWorksheet(
            workbook,
            sheet,
            worksheet,
            options,
            sampleLimit,
            outputContext,
          )
        : {
            ...EMPTY_SHEET,
            columns: [],
            name: sheet.name,
            visibility: publicVisibility(sheet.visibility),
          },
    );
    options.onProgress?.({
      operation: INSPECT_OPERATION,
      stage: "describing-worksheets",
      completed: index + 1,
      total: selected.length,
      detail: sheet.name,
    });
  }

  throwIfAborted(options.signal, INSPECT_OPERATION, outputContext);
  const describedSheets = new Set(selected.map((sheet) => sheet.name));
  const description: WorkbookDescription = {
    source,
    sheets,
    excelTables: describeExcelTables(await workbook.tables(), describedSheets),
    namedRanges: describeNamedRanges(
      await workbook.definedNames(),
      describedSheets,
    ),
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
    result: workbookDescriptionResult(description, hiddenExcluded),
  };
}
