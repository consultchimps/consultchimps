/**
 * L2 region layer: the Excel Table binding.
 *
 * `TableBinding` owns everything that is true only of Excel Tables: the body
 * stops above a totals row, the columns come from the table definition rather
 * than from reading header cells, and a filtered table compacts *as a table* -
 * the surviving data rows close up, the totals row follows them, and whatever
 * sits below the table stays exactly where the author put it. A plain
 * "delete these rows and close every gap" cannot say that, which is why the
 * edit goes through an explicit `RowRelocation` rather than `deleteRows`.
 *
 * Resizing the table part's `ref`, its autoFilter, and every other dependent
 * structure is the model's job: the invariant pass performs all of it. This
 * binding only decides *where* each row lands and whether the table part
 * follows.
 */

import { RowRelocation } from "../model/references.js";
import type {
  CellRange,
  RowNumber,
  WorkbookTableInfo,
  WorksheetModel,
} from "../model/types.js";
import {
  evaluateFormulaGuard,
  type FormulaGuardVerdict,
} from "./range-binding.js";
import type {
  ColumnInfo,
  DataRegion,
  MatchingPolicy,
  RegionEditReport,
  RegionOrigin,
} from "./types.js";
import { readRowKeys } from "./values.js";

/**
 * A `RegionEditReport` widened with what the table did. `tableResized` is
 * false when the sheet holds formulas written against row positions: the rows
 * still go and every reference still follows them, but the table part is left
 * claiming its original range rather than being silently narrowed around
 * formulas the author placed by hand.
 */
export interface TableEditReport extends RegionEditReport {
  readonly formulaGuard: FormulaGuardVerdict;
  readonly tableResized: boolean;
}

export function isTableEditReport(
  report: RegionEditReport,
): report is TableEditReport {
  return "tableResized" in report;
}

export class TableBinding implements DataRegion {
  readonly body: CellRange;
  readonly columns: readonly ColumnInfo[];
  readonly headerRow: RowNumber;
  readonly origin: RegionOrigin;
  readonly sheetName: string;
  /** True when the last row of the table range is a totals row. */
  readonly totalsRow: boolean;
  readonly worksheet: WorksheetModel;
  /** The last row the table range covers, totals row included. */
  readonly #tableEndRow: RowNumber;

  constructor(worksheet: WorksheetModel, table: WorkbookTableInfo) {
    this.worksheet = worksheet;
    this.sheetName = table.sheetName;
    this.origin = { kind: "table", tableName: table.name };
    this.headerRow = table.headerRow;
    this.totalsRow = table.totalsRow;
    this.#tableEndRow = table.range.end.row;
    // The body starts below the header and stops above the totals row. An
    // empty table yields end.row === start.row - 1, which every loop below
    // treats as "no rows" without a special case.
    this.body = {
      end: {
        column: table.range.end.column,
        row: table.range.end.row - (table.totalsRow ? 1 : 0),
      },
      start: { column: table.range.start.column, row: table.headerRow + 1 },
    };
    this.columns = table.columnNames.map((name, offset) => ({
      index: table.range.start.column + offset,
      name,
    }));
  }

  rowKeys(
    column: ColumnInfo,
    matching: MatchingPolicy,
  ): ReadonlyMap<RowNumber, string | undefined> {
    return readRowKeys(
      this.worksheet,
      this.body.start.row,
      this.body.end.row,
      column.index,
      matching,
    );
  }

  /**
   * Whether the sheet holds a formula written against row positions. Excel
   * Table formulas are structured references, which move with the table, so a
   * verdict of `canRenumber` is the normal case; anything else is a hand-
   * written A1, shared or array formula somewhere on the same sheet.
   */
  formulaGuard(): FormulaGuardVerdict {
    return evaluateFormulaGuard(this.worksheet);
  }

  filterRows(keep: (row: RowNumber) => boolean): TableEditReport {
    const kept: RowNumber[] = [];
    const doomed = new Set<RowNumber>();
    for (let row = this.body.start.row; row <= this.body.end.row; row += 1) {
      if (keep(row)) {
        kept.push(row);
      } else {
        doomed.add(row);
      }
    }

    const formulaGuard = this.formulaGuard();
    const report: TableEditReport = {
      deletedRows: doomed.size,
      formulaGuard,
      retainedRows: kept.length,
      tableResized: formulaGuard.canRenumber,
    };
    if (doomed.size === 0) {
      return report;
    }

    if (!formulaGuard.canRenumber) {
      // The fallback: rows leave and everything that described them follows,
      // but the table part keeps the range it declared, so a reviewer sees a
      // table that still covers the block the author drew.
      this.worksheet.applyRowRelocation(
        RowRelocation.compacting(
          doomed,
          Math.max(this.worksheet.lastRow, this.#tableEndRow),
          true,
        ),
        { resizeTables: false },
      );
      return report;
    }

    this.worksheet.applyRowRelocation(this.#compactionPlan(kept, doomed));
    return report;
  }

  /**
   * Where every row of the table goes: survivors close up under the header,
   * the totals row lands directly beneath them, and rows outside the table
   * range are absent from the plan, which leaves them untouched.
   */
  #compactionPlan(
    kept: readonly RowNumber[],
    doomed: ReadonlySet<RowNumber>,
  ): RowRelocation {
    const entries: Array<[RowNumber, RowNumber | null]> = [];
    kept.forEach((row, index) => {
      entries.push([row, this.body.start.row + index]);
    });
    for (const row of doomed) {
      entries.push([row, null]);
    }
    if (this.totalsRow) {
      entries.push([this.#tableEndRow, this.body.start.row + kept.length]);
    }
    return RowRelocation.explicit(entries);
  }
}
