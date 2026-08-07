/**
 * L1 seam — the document model. See packages/xlsx/ARCHITECTURE.md.
 *
 * This file freezes the interface between parallel implementation streams.
 * Deviations require justification at integration time; do not edit casually.
 *
 * Phase 1 integration added three members the split operation and the region
 * bindings genuinely need and could not express through `deleteRows` alone:
 *
 *  - `WorksheetModel.cellValue` — typed cell values (a date-formatted number
 *    comes back as a `Date`), so grouping keys and therefore output filenames
 *    match the SheetJS `cellDates` reads the previous engine made.
 *  - `WorksheetModel.lastRow` and `applyRowRelocation` — an Excel Table
 *    compacts its data rows while its totals row follows and the block below
 *    the table stays put. That is a row *move*, not a gap closure, so
 *    `deleteRows` cannot express it.
 *  - `RelocateRowsOptions.resizeTables` — the one case where a table binding
 *    deliberately leaves its table part claiming the original range.
 */

import type { RowRelocation } from "./references.js";

/** One-based row number as used by OOXML `r` attributes. */
export type RowNumber = number;
/** Zero-based column index (A = 0). */
export type ColumnIndex = number;

export interface CellRef {
  readonly row: RowNumber;
  readonly column: ColumnIndex;
}

export interface CellRange {
  readonly start: CellRef;
  readonly end: CellRef;
}

export type FormulaKind = "normal" | "shared" | "array";

export interface CellFormula {
  readonly kind: FormulaKind;
  readonly text: string;
  /** Range attribute for shared/array formulas, when present. */
  readonly range?: CellRange | undefined;
  readonly sharedIndex?: number | undefined;
}

export interface CellModel {
  readonly ref: CellRef;
  readonly type?: string | undefined;
  readonly styleIndex?: number | undefined;
  readonly value?: string | undefined;
  readonly formula?: CellFormula | undefined;
}

export interface RowModel {
  readonly number: RowNumber;
  readonly cells: readonly CellModel[];
  /** Raw non-cell row attributes preserved verbatim (height, hidden, ...). */
  readonly attributes: Readonly<Record<string, string>>;
}

export interface SheetInfo {
  readonly name: string;
  readonly partPath: string;
  readonly visibility: "visible" | "hidden" | "veryHidden";
}

export interface DefinedNameEntry {
  readonly name: string;
  readonly reference: string;
  /** Sheet-scoped names carry the sheet index; workbook scope is undefined. */
  readonly localSheetId?: number | undefined;
}

export interface DeleteRowsOptions {
  /** Renumber surviving rows to be consecutive (compaction). */
  readonly renumber: boolean;
}

export interface RelocateRowsOptions {
  /**
   * Excel Tables anchored on the worksheet follow the edit. Defaults to true;
   * `false` leaves every table part claiming the range it already claimed.
   */
  readonly resizeTables?: boolean | undefined;
}

/** A cell value in the shapes a grouping key can be built from. */
export type CellValue = Date | number | boolean | string | undefined;

export interface DeleteRowsReport {
  readonly deletedRows: number;
  readonly retainedRows: number;
  /** Structures adjusted during the invariant pass, for warnings/metrics. */
  readonly adjusted: {
    readonly mergedRanges: number;
    readonly conditionalFormatting: number;
    readonly dataValidations: number;
    readonly hyperlinks: number;
    readonly formulaReferences: number;
    readonly tableRefs: number;
    readonly calcChainEntries: number;
  };
}

/**
 * Structured view over one worksheet part. Edits go through methods; the
 * model re-serializes only when something changed.
 */
export interface WorksheetModel {
  readonly info: SheetInfo;
  readonly usedRange: CellRange | undefined;
  /** The highest row number present in the sheet, or 0 when it has none. */
  readonly lastRow: RowNumber;
  rows(): readonly RowModel[];
  row(number: RowNumber): RowModel | undefined;
  cellText(ref: CellRef): string | undefined;
  /**
   * A cell's value typed the way a grouping key needs it: a number whose
   * style formats it as a date comes back as a `Date`, everything else keeps
   * the type the OOXML cell declares.
   */
  cellValue(ref: CellRef): CellValue;
  /**
   * The general form of `deleteRows`: an explicit plan saying where each row
   * lands. Both funnel into the same invariant pass.
   */
  applyRowRelocation(
    relocation: RowRelocation,
    options?: RelocateRowsOptions | undefined,
  ): DeleteRowsReport;
  /**
   * Delete the given rows and maintain every dependent invariant in one
   * pass: row/cell renumbering (when `renumber`), formula-text reference
   * adjustment (A1 and ranges, shared-formula ref attributes), merged
   * ranges, conditional-formatting and data-validation sqrefs, hyperlink
   * refs, table ref/autoFilter resize, and calcChain maintenance.
   */
  deleteRows(
    rows: ReadonlySet<RowNumber>,
    options: DeleteRowsOptions,
  ): DeleteRowsReport;
}

export interface WorkbookModel {
  readonly sheets: readonly SheetInfo[];
  worksheet(name: string): WorksheetModel | undefined;
  definedNames(): Promise<readonly DefinedNameEntry[]>;
  /** Excel Table definitions, from the existing excel-tables reader. */
  tables(): Promise<readonly WorkbookTableInfo[]>;
  /** Serialize: changed parts rewritten, untouched parts byte-identical. */
  save(): Promise<Uint8Array>;
}

export interface WorkbookTableInfo {
  readonly name: string;
  readonly sheetName: string;
  readonly partPath: string;
  readonly range: CellRange;
  readonly headerRow: RowNumber;
  readonly totalsRow: boolean;
  readonly columnNames: readonly string[];
}
