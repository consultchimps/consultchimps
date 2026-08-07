/**
 * L1 seam — the document model. See packages/xlsx/ARCHITECTURE.md.
 *
 * This file freezes the interface between parallel implementation streams.
 * Deviations require justification at integration time; do not edit casually.
 */

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
  rows(): readonly RowModel[];
  row(number: RowNumber): RowModel | undefined;
  cellText(ref: CellRef): string | undefined;
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
