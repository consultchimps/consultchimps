/**
 * The SheetJS-backed cell reader the streaming reader replaced, kept only as a
 * test oracle until SheetJS leaves the repository (ADR 0006).
 */
import { ConsultChimpsError } from "@consultchimps/core";
import type { CellValue } from "@consultchimps/tabular";
import * as XLSX from "xlsx";

import { XLSX_ERRORS } from "../src/errors.js";
import {
  calendarIsoText,
  isComponentsInRange,
  utcCalendarParts,
} from "../src/model/calendar.js";
import type { WorkbookModel } from "../src/model/index.js";
import type { WorksheetModel } from "../src/model/types.js";
import { WorkbookRead } from "../src/operations/read-model.js";

export interface ParseWorkbookOptions {
  /** Machine-readable context added to a read failure. */
  details?: Record<string, unknown> | undefined;
  /** Keep cached display text, which worksheet records report verbatim. */
  cellText?: boolean | undefined;
}

/**
 * Parse workbook bytes, reporting an unreadable workbook as a stable error.
 * The source label appears in the message, so callers pass a file path or an
 * in-memory input name.
 */
export function parseWorkbookBytes(
  workbookBytes: Uint8Array,
  source: string,
  options: ParseWorkbookOptions = {},
): XLSX.WorkBook {
  try {
    return XLSX.read(workbookBytes, {
      // The serial is kept rather than turned into a `Date`. A `Date` has a
      // local face as well as a UTC one, and which of them is the workbook's
      // depends on how the engine composed it, so text built from one is text
      // built on somebody else's convention. Which cells are dates, and what
      // they hold, is the document model's answer: see `WorkbookDates`.
      cellDates: false,
      cellText: options.cellText ?? false,
      dense: false,
      type: "array",
    });
  } catch (error) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${source}`,
      {
        cause: error,
        details: options.details ?? { source },
      },
    );
  }
}

/**
 * What a worksheet holds in the cells it declares or formats as dates, read
 * from the document model.
 *
 * The engine cannot answer this. A worksheet may say a cell is a date in two
 * ways: by wearing a date number format, or by declaring `t="d"` and writing
 * ISO 8601 text. Reading with `cellDates` off, which is what keeps a serial a
 * serial, turns the second kind into a plain number and drops the declaration
 * entirely: measured on this engine, `<c t="d"><v>2024-01-01</v></c>` arrives
 * as the number 45292 with no field left saying what it was. So the reader
 * asked the style alone, and an unstyled date cell became an integer, in the
 * table and in the schema an import inferred from it. Reading with `cellDates`
 * on keeps the declaration but hands back the engine's own parse of the text,
 * which remaps a year of 0099 to 1999 and normalises a month of 13 into
 * January of the next year: two defects the model refuses.
 *
 * So the model answers, for both kinds. It is the reader that sees the
 * declared type, it owns the style table, and its values come through the one
 * calendar route. The engine keeps the work it is good at: the used range, the
 * header row, and every cell that is not a date.
 */
export interface WorkbookDates {
  /** The dates one worksheet holds. */
  forSheet(sheet: string): SheetDates;
}

/**
 * What the model read in a date cell, and how the cell said it was one.
 *
 * The two ways differ in what the engine makes of the cell, so they differ in
 * how far this has to override it. A cell wearing a date format is a number the
 * engine formats correctly for display and reports as a serial: its stored
 * value comes from here, and its displayed text is the engine's, which is the
 * text the worksheet shows. A cell that declares `t="d"` is one the engine has
 * already turned into a serial, so both its value and its displayed text come
 * from here; the engine's display for it is that serial, which is a number
 * nobody wrote.
 */
interface CellDate {
  /**
   * One of the two answers a date cell has: the canonical timestamp, or the
   * cell's own text where that names no moment. Never blank - a cell holding
   * nothing is not a date cell, and `WorksheetModel.cellValue` reads it as no
   * value at all, so it never reaches this map and the reader's ordinary blank
   * rule answers for it.
   */
  readonly value: string;
  readonly declared: boolean;
}

/** A worksheet's date cells, by row and column. */
type WorksheetDates = ReadonlyMap<string, CellDate>;

/**
 * One worksheet's dates, asked for the way the engine indexes its cells: a
 * zero-based row, as `getCell` takes, rather than the one-based row a worksheet
 * writes. Converting here means no reader has to remember which of the two it
 * is holding.
 */
export type SheetDates = (
  rowIndex: number,
  columnIndex: number,
) => CellDate | undefined;

function dateKey(row: number, column: number): string {
  return `${row},${column}`;
}

/**
 * Collect a worksheet's date cells in one pass.
 *
 * A cell is a date when it declares itself one or when its style says so; both
 * questions are the model's, and the value is the model's too, so the reader
 * has one rule and one spelling rather than a second set of its own. A cell
 * the model reads as text - a `t="d"` cell whose text names no moment - is
 * kept as that text, because that is what the file says it is.
 */
function collectWorksheetDates(
  model: WorkbookModel,
  worksheet: WorksheetModel | undefined,
): WorksheetDates {
  const dates = new Map<string, CellDate>();
  if (!worksheet) {
    return dates;
  }
  for (const row of worksheet.rows()) {
    for (const cell of row.cells) {
      const declared = cell.type === DECLARED_DATE_TYPE;
      if (!declared && !model.isDateStyle(cell.styleIndex)) {
        continue;
      }
      const value = worksheet.cellValue(cell.ref);
      if (value instanceof Date) {
        dates.set(dateKey(cell.ref.row, cell.ref.column), {
          declared,
          value: calendarIsoText(utcCalendarParts(value)),
        });
      } else if (declared && typeof value === "string") {
        dates.set(dateKey(cell.ref.row, cell.ref.column), {
          declared,
          value,
        });
      }
    }
  }
  return dates;
}

/** The OOXML cell type a worksheet declares a date with. */
const DECLARED_DATE_TYPE = "d";

/**
 * The dates a workbook's bytes hold, for a reader that has bytes rather than a
 * loaded model. The worksheets are parsed on demand, so a reader that touches
 * one worksheet pays for one.
 */
export async function readWorkbookDates(
  bytes: Uint8Array,
  source: string,
  details: Record<string, unknown>,
): Promise<WorkbookDates> {
  return workbookDatesFrom(await WorkbookRead.load(bytes, { source, details }));
}

/**
 * The dates a workbook holds, read once per worksheet and only when a reader
 * asks for that worksheet.
 */
export function workbookDatesFrom(read: WorkbookRead): WorkbookDates {
  const bySheet = new Map<string, WorksheetDates>();
  return {
    forSheet(sheet) {
      return (rowIndex, columnIndex) => {
        let dates = bySheet.get(sheet);
        if (dates === undefined) {
          dates = collectWorksheetDates(read.workbook, read.worksheet(sheet));
          bySheet.set(sheet, dates);
        }
        return dates.get(dateKey(rowIndex + 1, columnIndex));
      };
    },
  };
}

/**
 * The same spelling for a moment that arrived already parsed.
 *
 * Unreachable from this package's own reading: `parseWorkbookBytes` keeps
 * serials, so no cell it produces carries a `Date`. It is here so a cell from a
 * workbook somebody else parsed with `cellDates` cannot fall through to
 * `String(v)` and arrive as a platform date string, and it goes through
 * `calendarIsoText` like everything else rather than through `toISOString`, so
 * there is one spelling of a date in this package and not two.
 *
 * The UTC face is the one read, because a moment composed from a workbook's
 * epoch and whole days wears the calendar date on that face and the local face
 * is that shifted by wherever the reader is sitting. A face outside the years
 * that can be written has no spelling, so it comes back undefined and the
 * caller falls through, the same rule the two paths above follow.
 */
function parsedDateText(value: Date): string | undefined {
  const parts = utcCalendarParts(value);
  return isComponentsInRange(parts) ? calendarIsoText(parts) : undefined;
}

export function cellToPrimitive(
  cell: XLSX.CellObject | undefined,
  date: CellDate | undefined,
): CellValue {
  // A date the model read is the stored value, whether or not the engine saw a
  // cell here at all: a `t="d"` cell whose text names no moment reaches the
  // engine as nothing, and the text is what the file holds.
  if (date !== undefined) {
    return date.value;
  }
  if (!cell || cell.v === null || cell.v === undefined) {
    return null;
  }

  if (cell.v instanceof Date) {
    const parsed = parsedDateText(cell.v);
    if (parsed !== undefined) {
      return parsed;
    }
  }

  if (typeof cell.v === "number") {
    // A cell cannot hold one of these; the engine makes them out of text it
    // could not read, such as the spaces in a declared date cell. Blank is what
    // the cell holds, and blank is what every other reader here calls it.
    return Number.isFinite(cell.v) ? cell.v : null;
  }

  if (typeof cell.v === "string" || typeof cell.v === "boolean") {
    return cell.v;
  }

  return String(cell.w ?? cell.v);
}
