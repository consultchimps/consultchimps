/**
 * L2 region layer — value and reference normalization.
 *
 * The split-key normalization here is a faithful port of `normalizeSplitValue`
 * and `normalizeHeader` from `src/workbook-column-split.ts`. The old engine
 * read values through SheetJS (which handed back JavaScript `Date`, `number`,
 * `boolean`, and `string` values); the region layer reads them through the L1
 * document model instead, which types cells the same way. Key shapes (`date:`,
 * `number:`, `boolean:`, `string:`) are unchanged so grouping stays stable
 * across the migration.
 *
 * A1 reference helpers live here too because both `resolve.ts` (parsing user
 * references) and the bindings (naming cells in reports) need them.
 */

import type {
  CellRange,
  CellRef,
  ColumnIndex,
  RowNumber,
  WorksheetModel,
} from "../model/types.js";
import type { MatchingPolicy } from "./types.js";

/**
 * Text that Excel would round-trip as a number. Tolerant matching coerces
 * such text so "100" and 100 group together.
 */
const NUMERIC_TEXT_PATTERN =
  /^[+-]?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+))(?:e[+-]?\d+)?$/iu;

/**
 * `Sheet!A1:F200`, with the quoted-sheet form Excel writes when a sheet name
 * contains spaces. Ported verbatim from `NAMED_RANGE_REF_PATTERN` in
 * `src/shared.ts` so defined names resolve identically.
 */
const SHEET_RANGE_PATTERN =
  /^(?:'(?<quotedSheet>(?:[^']|'')+)'|(?<sheet>[^'!,:]+))!(?<range>\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$/u;

const CELL_REF_PATTERN = /^\$?(?<letters>[A-Za-z]{1,3})\$?(?<digits>\d+)$/u;

const LETTER_COUNT = 26;
const LETTER_A = 65;

/** A grouping key plus the human-readable value it came from. */
export interface NormalizedValue {
  readonly display: string;
  readonly key: string;
}

/** NFKC + trim + case-fold, the header comparison used everywhere in L2. */
export function normalizeHeader(value: string): string {
  return value.normalize("NFKC").trim().toLowerCase();
}

/**
 * Normalize a cell value into a grouping key. `undefined` means "blank" and
 * the caller (a `BlankPolicy` at L3) decides what to do with it.
 *
 * Strict matching keys strings by their raw text, so "North" and "north " stay
 * apart. Tolerant matching trims, coerces numeric-looking text to numbers, and
 * folds case after NFKC.
 */
export function normalizeSplitValue(
  value: unknown,
  strict: boolean,
): NormalizedValue | undefined {
  if (value === null || value === undefined) {
    return undefined;
  }
  if (value instanceof Date) {
    const display = value.toISOString();
    return { display, key: `date:${display}` };
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      return undefined;
    }
    const display = String(Object.is(value, -0) ? 0 : value);
    return { display, key: `number:${display}` };
  }
  if (typeof value === "boolean") {
    return { display: String(value), key: `boolean:${String(value)}` };
  }

  const display = String(value).trim();
  if (!display) {
    return undefined;
  }
  if (strict) {
    return { display, key: `string:${String(value)}` };
  }
  if (NUMERIC_TEXT_PATTERN.test(display)) {
    const numericValue = Number(display);
    if (Number.isFinite(numericValue)) {
      return {
        display,
        key: `number:${String(Object.is(numericValue, -0) ? 0 : numericValue)}`,
      };
    }
  }
  return {
    display,
    key: `string:${display.normalize("NFKC").toLowerCase()}`,
  };
}

/**
 * Typed value of one cell, or `undefined` when the cell is absent or blank.
 *
 * Typing is L1's job: the model resolves shared strings, reads the OOXML `t`
 * attribute, and consults the cell's number format so a date-formatted serial
 * arrives as a `Date`. That last part is why the region layer does not do this
 * itself - a `date:` key and a `number:` key produce different output
 * filenames, and the previous engine read dates through SheetJS `cellDates`.
 */
export function cellValue(worksheet: WorksheetModel, ref: CellRef): unknown {
  return worksheet.cellValue(ref);
}

/**
 * Normalized value of one column for every row in `[firstRow, lastRow]`.
 * Every row gets an entry, blank ones mapping to `undefined`, so callers can
 * count retained and skipped rows without re-walking the sheet.
 */
export function readRowValues(
  worksheet: WorksheetModel,
  firstRow: RowNumber,
  lastRow: RowNumber,
  column: ColumnIndex,
  matching: MatchingPolicy,
): Map<RowNumber, NormalizedValue | undefined> {
  const values = new Map<RowNumber, NormalizedValue | undefined>();
  for (let row = firstRow; row <= lastRow; row += 1) {
    values.set(
      row,
      normalizeSplitValue(
        cellValue(worksheet, { column, row }),
        matching.strict,
      ),
    );
  }
  return values;
}

/** `readRowValues` reduced to the keys the `DataRegion` interface returns. */
export function readRowKeys(
  worksheet: WorksheetModel,
  firstRow: RowNumber,
  lastRow: RowNumber,
  column: ColumnIndex,
  matching: MatchingPolicy,
): Map<RowNumber, string | undefined> {
  const keys = new Map<RowNumber, string | undefined>();
  for (const [row, value] of readRowValues(
    worksheet,
    firstRow,
    lastRow,
    column,
    matching,
  )) {
    keys.set(row, value?.key);
  }
  return keys;
}

/** Zero-based column index to its Excel letters (0 -> "A", 26 -> "AA"). */
export function columnLetters(column: ColumnIndex): string {
  let remaining = column;
  let letters = "";
  do {
    letters = `${String.fromCharCode(LETTER_A + (remaining % LETTER_COUNT))}${letters}`;
    remaining = Math.floor(remaining / LETTER_COUNT) - 1;
  } while (remaining >= 0);
  return letters;
}

/** Excel column letters to a zero-based index ("A" -> 0, "AA" -> 26). */
export function parseColumnLetters(letters: string): ColumnIndex {
  let column = 0;
  for (const letter of letters.toUpperCase()) {
    column = column * LETTER_COUNT + (letter.charCodeAt(0) - LETTER_A + 1);
  }
  return column - 1;
}

/** A cell reference as Excel writes it, for messages and warnings. */
export function formatCellRef(ref: CellRef): string {
  return `${columnLetters(ref.column)}${String(ref.row)}`;
}

/** Parse "A1" or "$A$1"; `undefined` when the text is not a cell reference. */
export function parseCellRef(text: string): CellRef | undefined {
  const match = CELL_REF_PATTERN.exec(text.trim());
  const letters = match?.groups?.letters;
  const digits = match?.groups?.digits;
  if (!letters || !digits) {
    return undefined;
  }
  const row = Number(digits);
  if (!Number.isInteger(row) || row < 1) {
    return undefined;
  }
  return { column: parseColumnLetters(letters), row };
}

/**
 * Parse "A1:F200" (or the single-cell "A1"). Corners are normalized so the
 * start is always the top-left cell, matching what `decode_range` produced.
 */
export function parseCellRange(text: string): CellRange | undefined {
  const [first, second, ...rest] = text.trim().split(":");
  if (first === undefined || rest.length > 0) {
    return undefined;
  }
  const start = parseCellRef(first);
  if (!start) {
    return undefined;
  }
  if (second === undefined) {
    return { end: start, start };
  }
  const end = parseCellRef(second);
  if (!end) {
    return undefined;
  }
  return {
    end: {
      column: Math.max(start.column, end.column),
      row: Math.max(start.row, end.row),
    },
    start: {
      column: Math.min(start.column, end.column),
      row: Math.min(start.row, end.row),
    },
  };
}

/** Parse "Sheet1!A1:F200" or "'My Sheet'!A1:F200". */
export function parseSheetRange(
  text: string,
): { readonly range: CellRange; readonly sheet: string } | undefined {
  const match = SHEET_RANGE_PATTERN.exec(text.trim());
  const sheet =
    match?.groups?.quotedSheet?.replaceAll("''", "'") ?? match?.groups?.sheet;
  const reference = match?.groups?.range;
  if (!sheet || !reference) {
    return undefined;
  }
  const range = parseCellRange(reference.replaceAll("$", ""));
  if (!range) {
    return undefined;
  }
  return { range, sheet };
}
