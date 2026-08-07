/**
 * L2 region layer — the worksheet-range binding.
 *
 * `RangeBinding` owns everything that is true only of plain ranges: the
 * boundaries come from detection, a declared header row, a defined name, or an
 * explicit `Sheet!A1:F200`; the columns are read from the header cells; and
 * deletion carries the formula-safety guard inherited from the split engine.
 *
 * The guard is the port of `tableCanBeCompacted` in
 * `src/workbook-column-split.ts`. That function decided whether rows could be
 * removed with compaction by scanning worksheet XML for `t="shared"` /
 * `t="array"` formulas and for A1-style references in formula text. The same
 * decision is made here over the model's structured formula data instead of
 * over raw XML, which removes two sources of error: XML entity escaping, and
 * string literals that had to be stripped with a regex before the reference
 * scan could run.
 *
 * Once L1's invariant pass adjusts formula references natively this guard is
 * belt and braces. It stays as the conservative default, and the decision is
 * reported (`RangeEditReport.formulaGuard`) so an operation can raise the same
 * warning the split engine raised.
 */

import type {
  CellRange,
  ColumnIndex,
  RowNumber,
  WorksheetModel,
} from "../model/types.js";
import type {
  ColumnInfo,
  DataRegion,
  MatchingPolicy,
  RegionEditReport,
  RegionOrigin,
} from "./types.js";
import { formatCellRef, readRowKeys } from "./values.js";

/**
 * An A1-style reference: optional `$`, one to three letters, optional `$`,
 * digits — not preceded by an identifier character (so `Table1[#Data]` and
 * `_xlfn.IFS` are not references) and not followed by one or by `(` (so
 * `LOG10(x)` is a function call, not a reference). Ported verbatim.
 */
const A1_REFERENCE_PATTERN =
  /(?<![A-Za-z0-9_.$])\$?[A-Za-z]{1,3}\$?\d+(?![\dA-Za-z_(])/u;

/** Quoted text inside a formula never contains a live reference. */
const STRING_LITERAL_PATTERN = /"[^"]*"/gu;

/** Why renumbering was refused, or `undefined` when it was allowed. */
export type FormulaGuardReason =
  "array-formula" | "a1-reference" | "shared-formula";

export interface FormulaGuardVerdict {
  readonly canRenumber: boolean;
  readonly reason?: FormulaGuardReason | undefined;
  /** A1 address of the first formula that blocked renumbering. */
  readonly cell?: string | undefined;
}

/**
 * A `RegionEditReport` widened with the guard decision. Returning a subtype
 * keeps `DataRegion.filterRows` satisfied while giving operations enough to
 * warn that a sheet was filtered without compaction.
 */
export interface RangeEditReport extends RegionEditReport {
  readonly formulaGuard: FormulaGuardVerdict;
  readonly renumbered: boolean;
}

export function isRangeEditReport(
  report: RegionEditReport,
): report is RangeEditReport {
  return "formulaGuard" in report;
}

export interface RangeBindingOptions {
  readonly body: CellRange;
  /** Defaults to the header-row cells across the body's column span. */
  readonly columns?: readonly ColumnInfo[] | undefined;
  readonly headerRow: RowNumber;
  readonly origin: RegionOrigin;
  readonly worksheet: WorksheetModel;
}

/**
 * Column names read straight from the header cells. Blank headers keep an
 * empty name and their position: deduplication and renaming are a `tabular`
 * concern, not a boundary concern.
 */
export function readHeaderColumns(
  worksheet: WorksheetModel,
  headerRow: RowNumber,
  startColumn: ColumnIndex,
  endColumn: ColumnIndex,
): ColumnInfo[] {
  const columns: ColumnInfo[] = [];
  for (let column = startColumn; column <= endColumn; column += 1) {
    columns.push({
      index: column,
      name: worksheet.cellText({ column, row: headerRow }) ?? "",
    });
  }
  return columns;
}

/**
 * Whether every formula on the sheet survives row renumbering. Shared and
 * array formulas carry a range attribute that ties them to row positions, and
 * a formula holding an A1 reference means some cell points at a row by number.
 */
export function evaluateFormulaGuard(
  worksheet: WorksheetModel,
): FormulaGuardVerdict {
  for (const row of worksheet.rows()) {
    for (const cell of row.cells) {
      const formula = cell.formula;
      if (!formula) {
        continue;
      }
      if (formula.kind === "shared" || formula.kind === "array") {
        return {
          canRenumber: false,
          cell: formatCellRef(cell.ref),
          reason:
            formula.kind === "shared" ? "shared-formula" : "array-formula",
        };
      }
      const expression = formula.text.replace(STRING_LITERAL_PATTERN, "");
      if (A1_REFERENCE_PATTERN.test(expression)) {
        return {
          canRenumber: false,
          cell: formatCellRef(cell.ref),
          reason: "a1-reference",
        };
      }
    }
  }
  return { canRenumber: true };
}

export class RangeBinding implements DataRegion {
  readonly body: CellRange;
  readonly columns: readonly ColumnInfo[];
  readonly headerRow: RowNumber;
  readonly origin: RegionOrigin;
  readonly sheetName: string;
  readonly worksheet: WorksheetModel;

  constructor(options: RangeBindingOptions) {
    this.worksheet = options.worksheet;
    this.sheetName = options.worksheet.info.name;
    this.origin = options.origin;
    this.headerRow = options.headerRow;
    this.body = options.body;
    this.columns =
      options.columns ??
      readHeaderColumns(
        options.worksheet,
        options.headerRow,
        options.body.start.column,
        options.body.end.column,
      );
  }

  /**
   * The guard decision for the sheet as it stands now. Recomputed on demand
   * rather than cached, because the worksheet is mutable.
   */
  formulaGuard(): FormulaGuardVerdict {
    return evaluateFormulaGuard(this.worksheet);
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

  filterRows(keep: (row: RowNumber) => boolean): RangeEditReport {
    const formulaGuard = this.formulaGuard();
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
      this.worksheet.deleteRows(doomed, { renumber: formulaGuard.canRenumber });
    }
    return {
      deletedRows: doomed.size,
      formulaGuard,
      renumbered: formulaGuard.canRenumber,
      retainedRows,
    };
  }
}
