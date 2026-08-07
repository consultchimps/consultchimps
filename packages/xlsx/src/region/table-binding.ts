/**
 * L2 region layer — the Excel Table binding.
 *
 * `TableBinding` owns everything that is true only of Excel Tables: the body
 * stops above a totals row, the columns come from the table definition rather
 * than from reading header cells, and deletions always compact because a
 * Table's formulas are structured references (`Table1[Amount]`), which do not
 * move when rows are renumbered.
 *
 * Resizing the table part's `ref`, its autoFilter, and every other dependent
 * structure is the model's job: `WorksheetModel.deleteRows` performs the whole
 * invariant pass. This binding only decides *which* rows go and whether the
 * survivors are renumbered.
 */

import type {
  CellRange,
  RowNumber,
  WorkbookTableInfo,
  WorksheetModel,
} from "../model/types.js";
import type {
  ColumnInfo,
  DataRegion,
  MatchingPolicy,
  RegionEditReport,
  RegionOrigin,
} from "./types.js";
import { readRowKeys } from "./values.js";

export class TableBinding implements DataRegion {
  readonly body: CellRange;
  readonly columns: readonly ColumnInfo[];
  readonly headerRow: RowNumber;
  readonly origin: RegionOrigin;
  readonly sheetName: string;
  /** True when the last row of the table range is a totals row. */
  readonly totalsRow: boolean;
  readonly worksheet: WorksheetModel;

  constructor(worksheet: WorksheetModel, table: WorkbookTableInfo) {
    this.worksheet = worksheet;
    this.sheetName = table.sheetName;
    this.origin = { kind: "table", tableName: table.name };
    this.headerRow = table.headerRow;
    this.totalsRow = table.totalsRow;
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

  filterRows(keep: (row: RowNumber) => boolean): RegionEditReport {
    const doomed = new Set<RowNumber>();
    let retainedRows = 0;
    for (let row = this.body.start.row; row <= this.body.end.row; row += 1) {
      if (keep(row)) {
        retainedRows += 1;
      } else {
        doomed.add(row);
      }
    }
    if (doomed.size > 0) {
      this.worksheet.deleteRows(doomed, { renumber: true });
    }
    return { deletedRows: doomed.size, retainedRows };
  }
}
