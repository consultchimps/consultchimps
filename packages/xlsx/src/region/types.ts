/**
 * L2 seam: data regions. See packages/xlsx/ARCHITECTURE.md.
 *
 * This file freezes the interface between parallel implementation streams.
 * Deviations require justification at integration time; do not edit casually.
 */

import type {
  CellRange,
  RowNumber,
  WorkbookModel,
  WorksheetModel,
} from "../model/types.js";

export interface ColumnInfo {
  readonly name: string;
  readonly index: number;
}

/** How a region was located; carried for provenance and warnings. */
export type RegionOrigin =
  | { readonly kind: "table"; readonly tableName: string }
  | { readonly kind: "named-range"; readonly rangeName: string }
  | { readonly kind: "explicit-range"; readonly reference: string }
  | { readonly kind: "declared-header" }
  | { readonly kind: "detected-header" };

export type RegionSelector =
  | { readonly table: string }
  | { readonly range: string }
  | { readonly sheet: string; readonly headerRow?: number | undefined }
  | { readonly find: string }
  | "all-worksheets";

export interface MatchingPolicy {
  readonly strict: boolean;
}

export interface BlankPolicy {
  readonly includeBlank: boolean;
}

export interface RegionEditReport {
  readonly deletedRows: number;
  readonly retainedRows: number;
}

/**
 * A bounded, header-carrying data region. Exactly two implementations exist:
 * TableBinding and RangeBinding. Operations program against this interface
 * only; the Table/range difference lives behind it.
 */
export interface DataRegion {
  readonly sheetName: string;
  readonly origin: RegionOrigin;
  readonly headerRow: RowNumber;
  readonly body: CellRange;
  readonly columns: readonly ColumnInfo[];
  readonly worksheet: WorksheetModel;
  /**
   * Normalized key (per MatchingPolicy) of the given column for each body
   * row, undefined for blank values.
   */
  rowKeys(
    column: ColumnInfo,
    matching: MatchingPolicy,
  ): ReadonlyMap<RowNumber, string | undefined>;
  /**
   * Keep only rows whose key passes `keep`; maintain all binding-specific
   * invariants (table ref/totals for TableBinding, formula-safety rules for
   * RangeBinding) plus the model's invariant pass.
   */
  filterRows(keep: (row: RowNumber) => boolean): RegionEditReport;
}

/**
 * Resolve regions for a selector. Owns all discovery heuristics (NFKC/trim/
 * case-insensitive header search, headerRow overrides, table association).
 * `column` is required for `{ find }` and `"all-worksheets"` selectors.
 */
export type ResolveRegions = (
  workbook: WorkbookModel,
  selector: RegionSelector,
  column?: string,
) => Promise<readonly DataRegion[]>;
