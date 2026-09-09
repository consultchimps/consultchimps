/**
 * L3: reading a workbook's worksheets, and what reading them had to treat as
 * empty.
 *
 * The tables themselves come from the worksheet reader, which resolves a header
 * row and takes the rows under it. What no `Table` can carry is the difference
 * between a cell that is empty and a cell holding a formula the workbook carries
 * no calculated value for: Excel writes a formula and its last result side by
 * side, and a file written by a generator, or saved with calculation switched
 * off, carries the formula alone. Every reader downstream then sees an empty
 * cell, because empty is all the file says, and a table read from that worksheet
 * has holes where the numbers belong.
 *
 * The count is read from the document model, because the spreadsheet engine the
 * table reader is built on drops a numeric formula cell with no cached value
 * while parsing, so by the time a `Table` exists the evidence is gone. It is
 * scoped to the rectangle the table reader reported rather than to a region
 * resolved a second time: this package resolves a worksheet's header two ways,
 * and they disagree on exactly the rows that read as blank, which is every row
 * holding nothing but uncalculated formulas. Given the region rather than
 * finding one, the count can only ever describe the read the caller got.
 *
 * This is the operation both surfaces call. It lives here, rather than in the
 * byte surface where the region happened to be in reach, because deciding what
 * a worksheet's formulas mean is semantics, and L5 adapts inputs and outputs
 * and nothing else.
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
import { loadWorkbookModelForDescribe } from "./describe.js";

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
}

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
 * Count the uncalculated formulas inside one read.
 *
 * A worksheet that yielded no table has no rectangle, so the whole of it is
 * counted: there is no region to be wrong about, and the count is what explains
 * why the worksheet looks empty.
 */
function countUncachedFormulas(
  worksheet: WorksheetModel,
  region: WorksheetRegion | undefined,
): number {
  let total = 0;
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
        total += 1;
      }
    }
  }
  return total;
}

/**
 * Read every selected worksheet, reporting each one whether or not it yielded a
 * table, the rectangle the read covered, and how many cells of that read hold a
 * formula the workbook carries no calculated value for.
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
  const model = await loadWorkbookModelForDescribe(bytes, source, details);
  return reports.map((report) => {
    const worksheet = model.worksheet(report.sheet);
    if (worksheet === undefined) {
      // A zero here would say "this worksheet has no uncalculated formulas",
      // which is a claim about a worksheet nothing looked at. The whole reason
      // this count exists is that a reader silently reporting nothing is
      // indistinguishable from a reader reporting nothing is there.
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_READ_FAILED,
        `Worksheet "${report.sheet}" could not be read from ${source} to check its formulas, so whether it holds values the workbook never calculated is unknown.`,
        { details: { ...details, worksheet: report.sheet } },
      );
    }
    return {
      ...report,
      uncachedFormulaCells: countUncachedFormulas(worksheet, report.region),
    };
  });
}
