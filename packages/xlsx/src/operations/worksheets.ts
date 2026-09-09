/**
 * L3: reading a workbook's worksheets, and what reading them could not say.
 *
 * The tables themselves come from the worksheet reader, which resolves a header
 * row and takes the rows under it. Two things no `Table` can carry travel
 * beside it, because in both the cell the table holds is not the cell the
 * worksheet holds.
 *
 * A formula the workbook carries no calculated value for reads as *empty*.
 * Excel writes a formula and its last result side by side, and a file written
 * by a generator, or saved with calculation switched off, carries the formula
 * alone; every reader downstream then sees an empty cell, because empty is all
 * the file says, and the table has holes where the numbers belong.
 *
 * An error cell reads as a *number*. `#REF!`, `#DIV/0!` and their kind are
 * stored as `t="e"`, and the spreadsheet engine reports the internal code Excel
 * numbers each error by, so `#REF!` arrives as 23 and `#DIV/0!` as 7. That is
 * worse than a hole: nothing downstream can tell those from data, so a column
 * of amounts infers a numeric type and the table reads as complete.
 *
 * Both counts are read from the document model, because the engine the table
 * reader is built on drops a numeric formula cell with no cached value while
 * parsing and flattens an error into that code, so by the time a `Table` exists
 * the evidence is gone. They are scoped to the rectangle the table reader
 * reported rather than to a region resolved a second time: this package
 * resolves a worksheet's header two ways, and they disagree on exactly the rows
 * that read as blank, which is every row holding nothing but uncalculated
 * formulas. Given the region rather than finding one, a count can only ever
 * describe the read the caller got.
 *
 * This is the operation both surfaces call. It lives here, rather than in the
 * byte surface where the region happened to be in reach, because deciding what
 * a worksheet's cells mean is semantics, and L5 adapts inputs and outputs and
 * nothing else.
 */
import { ConsultChimpsError } from "@consultchimps/core";

import { XLSX_ERRORS } from "../errors.js";
import type { CellModel, WorksheetModel } from "../model/types.js";
import {
  parseWorkbookBytes,
  workbookWorksheetReports,
  type ReadWorkbookOptions,
  type WorksheetRegion,
  type WorksheetTableReport,
} from "../shared.js";
import { WorkbookRead } from "./read-model.js";

/**
 * One worksheet as the table reader saw it, with what the reader alone cannot
 * say about it.
 */
export interface WorksheetImportReport extends WorksheetTableReport {
  /**
   * Cells the read covered that hold a formula the workbook carries no
   * calculated value for, counted over the table's own region, or over the
   * whole worksheet when it yielded no table.
   *
   * Zero for a workbook Excel has calculated and saved. Anything above zero
   * means the reader saw those cells as empty, because the value they would
   * produce is not in the file.
   */
  uncachedFormulaCells: number;
  /**
   * Cells the read covered that hold an error value: `#REF!`, `#DIV/0!`,
   * `#N/A` and the rest, whether typed straight into the cell or left there by
   * a formula whose last calculation failed.
   *
   * Zero for a workbook whose values are all values. Anything above zero means
   * the table beside this report carries a number in those cells that the
   * worksheet does not: the spreadsheet engine reports an error cell as the
   * internal code Excel numbers it by, so a `#REF!` arrives as 23 and a
   * `#DIV/0!` as 7, and nothing downstream can tell those from data.
   */
  errorCells: number;
}

/**
 * The OOXML cell type of an error value. It is the one type whose stored value
 * is not what the cell means: every other type a worksheet writes carries the
 * text, number, boolean or date the cell shows, and this one carries Excel's
 * internal numbering of `#REF!`, `#DIV/0!` and their kind.
 */
const ERROR_CELL_TYPE = "e";

/**
 * Whether a cell holds a formula the workbook carries no calculated value for.
 *
 * The question is whether a cached value element is there, never what it holds:
 * a formula that evaluated to an empty string, or to zero, has been calculated.
 */
function isUncachedFormula(cell: CellModel): boolean {
  return cell.formula !== undefined && !cell.hasCachedValue;
}

/**
 * Whether a cell holds an error value.
 *
 * A cell someone typed `#N/A` into and a cell whose formula last evaluated to
 * `#DIV/0!` are the same cell to the format and the same problem to a reader,
 * so the formula is not part of the question.
 */
function isErrorCell(cell: CellModel): boolean {
  return cell.type === ERROR_CELL_TYPE;
}

/** What one read of one worksheet holds that a `Table` cannot express. */
interface UnreadableCells {
  uncachedFormulaCells: number;
  errorCells: number;
}

/**
 * Count both conditions inside one read, in one walk.
 *
 * One walk because they are one question - what does this rectangle hold that
 * the table beside it misrepresents - and two walks over two region rules is
 * how a report comes to describe two different reads.
 *
 * A worksheet that yielded no table has no rectangle, so the whole of it is
 * counted: there is no region to be wrong about, and the counts are what
 * explain why the worksheet looks the way it does.
 */
function countUnreadableCells(
  worksheet: WorksheetModel,
  region: WorksheetRegion | undefined,
): UnreadableCells {
  const counts: UnreadableCells = { errorCells: 0, uncachedFormulaCells: 0 };
  for (const row of worksheet.rows()) {
    if (
      region !== undefined &&
      (row.number < region.headerRow || row.number > region.lastRow)
    ) {
      continue;
    }
    for (const cell of row.cells) {
      if (
        region !== undefined &&
        (cell.ref.column < region.startColumn ||
          cell.ref.column > region.endColumn)
      ) {
        continue;
      }
      if (isUncachedFormula(cell)) {
        counts.uncachedFormulaCells += 1;
      }
      if (isErrorCell(cell)) {
        counts.errorCells += 1;
      }
    }
  }
  return counts;
}

/**
 * Read every selected worksheet, reporting each one whether or not it yielded a
 * table, the rectangle the read covered, and how many cells of that read hold a
 * formula the workbook carries no calculated value for or an error value.
 */
export async function readWorksheetReports(
  bytes: Uint8Array,
  source: string,
  options: ReadWorkbookOptions = {},
): Promise<WorksheetImportReport[]> {
  const details = { source };
  const reports = workbookWorksheetReports(
    parseWorkbookBytes(bytes, source, { details }),
    source,
    options,
  );
  const read = await WorkbookRead.load(bytes, { source, details });
  return reports.map((report) => {
    const worksheet = read.worksheet(report.sheet);
    if (worksheet === undefined) {
      // A zero here would say "this worksheet holds nothing of the kind",
      // which is a claim about a worksheet nothing looked at. The whole reason
      // these counts exist is that a reader silently reporting nothing is
      // indistinguishable from a reader reporting nothing is there.
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_READ_FAILED,
        `Worksheet "${report.sheet}" could not be read from ${source} to check its cells, so whether it holds values the workbook never calculated, or values that are errors, is unknown.`,
        { details: { ...details, worksheet: report.sheet } },
      );
    }
    return { ...report, ...countUnreadableCells(worksheet, report.region) };
  });
}
