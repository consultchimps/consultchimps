/**
 * L2 region layer: where a worksheet's header row is, and which of its columns
 * hold anything at all.
 *
 * Two readers in this package resolve a worksheet's header without being told
 * where it is: the SheetJS-backed table reader in `src/shared.ts`, which the
 * consolidation and the worksheet-records reader are built on, and the region
 * resolver in `resolve.ts`, which the inspection reports through. The rule
 * both apply lives here, once, so that the header row an inspection reports is
 * the header row a consolidation reads from. Each reader supplies what it can
 * see - for every row, how many values it holds and how many of those are not
 * text - and this module answers.
 *
 * The rule. A row is measured against the fullest of the rows that follow it,
 * up to `HEADER_LOOKAHEAD_ROWS` populated rows down: it qualifies as the
 * header when it holds more than half as many values as that fullest row, or
 * at most one value fewer. The header row is the first row that qualifies. A
 * report title in A1, a merged banner, or a "Prepared by" line with a name
 * beside it holds one or two values above a table that holds five, so those
 * rows are skipped, while a header row that leaves a column unnamed over data
 * still wins, because it is nearly as full as the rows beneath it. Measuring
 * against the rows that follow, rather than against the fullest row anywhere
 * on the sheet, is what keeps a wider block further down the sheet from
 * turning a whole table above it into "title rows". The "one fewer" clause
 * matters only on a two-column sheet, where one value is both what a title
 * holds and what a header missing a name holds: there the first row holding a
 * value wins, which is what every reader did before.
 *
 * The guard. A count cannot tell a header that leaves half its columns
 * unnamed from a title line above a full header: both are a sparse row over
 * fuller ones. Column names are text and data rows usually are not, so rows
 * are skipped as titles only when the row that would become the header holds
 * nothing but text, or when every skipped row holds a single value - a title,
 * a banner - and the header holds at least three, which is what lets a title
 * be skipped above year or month columns. When the guard refuses, the first
 * row holding a value is the header, as it always was, and nothing is lost:
 * the worst case is the old one, a title read as a header, which the
 * inspection shows and `headerRow` overrides. A declared header row is never
 * second-guessed: the rule applies only when no row was declared.
 *
 * The same idea, turned on its side, decides the columns. A column whose
 * header cell is blank and whose every cell under the header is blank is a
 * spacer - a gap left between two blocks - and holds nothing a table could
 * carry, so it is left out. A column with a blank header over values is kept:
 * the values are data, and the missing name is a naming concern that
 * `uniqueHeaders` fills in.
 *
 * "Value" here is what the readers mean by it everywhere else: a cell holding
 * `null`, `undefined`, or the empty string holds nothing. A formula the
 * workbook carries no calculated value for holds nothing, the same as it does
 * to every reader downstream, so a title cell nobody calculated does not push
 * the header down and a column of uncalculated formulas is a spacer. An error
 * value is a value: the cell holds something, however unwelcome. "Not text"
 * means a number, a boolean, or an error: the kinds a column name is never
 * made of. A date is text for this purpose on both sides, because the table
 * reader hands dates back as ISO text and month columns are named by dates.
 */

import type {
  CellModel,
  CellRange,
  ColumnIndex,
  RowModel,
  RowNumber,
  WorksheetModel,
} from "../model/types.js";
import type { ColumnInfo } from "./types.js";

/** How many populated rows below a candidate the rule measures it against. */
export const HEADER_LOOKAHEAD_ROWS = 10;

/** How a title block is told from a header: at most this many values. */
const TITLE_LINE_VALUES = 1;

/** The narrowest header a title may be skipped above when it is not all text. */
const NON_TEXT_HEADER_MIN_VALUES = 3;

/** The OOXML cell type of an error value. */
const ERROR_CELL_TYPE = "e";

/** What one row of a worksheet holds, as far as the header rule is concerned. */
export interface RowValueCount {
  /** The row, in whatever numbering the caller uses; only its order matters. */
  readonly row: RowNumber;
  /** Cells on the row holding a value, within the range the caller examined. */
  readonly values: number;
  /** Of those, the values that are a number, a boolean, or an error. */
  readonly nonTextValues: number;
}

/** Whether a cell holds nothing, by the definition every reader here shares. */
export function isBlankValue(value: unknown): boolean {
  return value === undefined || value === null || value === "";
}

/**
 * Whether a value is one a column name is never made of. Error cells are
 * classified by their cell type, because the two readers hand an error back
 * differently (one as the engine's numeric code, one as its text).
 */
export function isNonTextValue(
  value: unknown,
  cellType: string | undefined,
): boolean {
  return (
    cellType === ERROR_CELL_TYPE ||
    typeof value === "number" ||
    typeof value === "boolean"
  );
}

/**
 * Whether a row holding `values` values qualifies as the header when the
 * fullest row it is measured against holds `fullest`. Exported so the rule
 * can be tested as the property it is, over every width, rather than by
 * example.
 */
export function qualifiesAsHeader(values: number, fullest: number): boolean {
  // Strictly more than half, in integers, so a two-value "Prepared by" line
  // above a four-column table is a title row and a four-column header row
  // with one unnamed column above five-column rows is still the header. The
  // second clause admits a header one name short on a two-column sheet, and
  // nowhere else changes the answer; a row holding nothing never qualifies.
  return values > 0 && (values * 2 > fullest || values + 1 >= fullest);
}

/**
 * Whether the rows above a header may be skipped as titles, given what the
 * header holds and what the skipped rows hold. See the guard in the module
 * comment.
 */
function maySkipTitles(
  header: RowValueCount,
  skipped: readonly RowValueCount[],
): boolean {
  if (header.nonTextValues === 0) {
    return true;
  }
  return (
    header.values >= NON_TEXT_HEADER_MIN_VALUES &&
    skipped.every((row) => row.values <= TITLE_LINE_VALUES)
  );
}

/**
 * The header row among `counts`, or undefined when no row holds a value.
 *
 * `counts` is read in the order given, which callers keep top to bottom: the
 * answer is the first row that qualifies against the fullest of itself and
 * the populated rows that follow it within the lookahead, provided the guard
 * allows the rows above it to be skipped; otherwise the first populated row.
 */
export function detectHeaderRow(
  counts: readonly RowValueCount[],
): RowNumber | undefined {
  const populated = counts.filter((count) => count.values > 0);
  const first = populated[0];
  if (first === undefined) {
    return undefined;
  }
  for (let index = 0; index < populated.length; index += 1) {
    const candidate = populated[index]!;
    let fullest = candidate.values;
    const end = Math.min(populated.length, index + 1 + HEADER_LOOKAHEAD_ROWS);
    for (let ahead = index + 1; ahead < end; ahead += 1) {
      fullest = Math.max(fullest, populated[ahead]!.values);
    }
    if (!qualifiesAsHeader(candidate.values, fullest)) {
      continue;
    }
    return maySkipTitles(candidate, populated.slice(0, index))
      ? candidate.row
      : first.row;
  }
  // Unreachable: the last populated row is measured against itself alone.
  return first.row;
}

/**
 * Rows above `headerRow` that hold at least one value: the title rows a read
 * from `headerRow` leaves out. Blank rows above the header are not counted,
 * since nothing was skipped in them.
 */
export function countTitleRows(
  counts: readonly RowValueCount[],
  headerRow: RowNumber,
): number {
  let skipped = 0;
  for (const count of counts) {
    if (count.row < headerRow && count.values > 0) {
      skipped += 1;
    }
  }
  return skipped;
}

/**
 * One walk over a worksheet, answering both questions the readers ask: what
 * each row holds, for the header rule, and the last row each column holds a
 * value on, for the spacer rule. A column is a spacer for a header row when
 * its last value, if any, sits above that row.
 */
export interface WorksheetProfile {
  /** Every stored row inside the used range, in row order. */
  readonly counts: readonly RowValueCount[];
  /** The last row holding a value, per column that holds one anywhere. */
  readonly lastValueRow: ReadonlyMap<ColumnIndex, RowNumber>;
}

/**
 * Profile every stored row of `worksheet` inside `used`. Only the rows the
 * sheet stores are visited, so a declared extent far larger than the contents
 * costs nothing here.
 */
export function profileWorksheet(
  worksheet: WorksheetModel,
  rows: readonly RowModel[],
  used: CellRange,
): WorksheetProfile {
  const counts: RowValueCount[] = [];
  const lastValueRow = new Map<ColumnIndex, RowNumber>();
  for (const row of rows) {
    if (row.number < used.start.row || row.number > used.end.row) {
      continue;
    }
    let values = 0;
    let nonTextValues = 0;
    for (const cell of row.cells) {
      const column = cell.ref.column;
      if (column < used.start.column || column > used.end.column) {
        continue;
      }
      const value = worksheet.cellValue(cell.ref);
      if (isBlankValue(value)) {
        continue;
      }
      values += 1;
      if (isNonTextValue(value, cellType(cell))) {
        nonTextValues += 1;
      }
      const last = lastValueRow.get(column);
      if (last === undefined || last < row.number) {
        lastValueRow.set(column, row.number);
      }
    }
    counts.push({ nonTextValues, row: row.number, values });
  }
  // Stored order is document order, which a hand-written part may leave
  // unsorted; the rule reads rows top to bottom.
  counts.sort((left, right) => left.row - right.row);
  return { counts, lastValueRow };
}

function cellType(cell: CellModel): string | undefined {
  return cell.type;
}

/**
 * The columns of a worksheet region that hold anything: every column across
 * the body's span except the spacers, each named by its header cell (blank
 * when the header cell is). The header row itself counts as part of the
 * region, so a named column with no values under it is still a column.
 */
export function regionColumns(
  worksheet: WorksheetModel,
  headerRow: RowNumber,
  body: CellRange,
  lastValueRow: ReadonlyMap<ColumnIndex, RowNumber>,
): ColumnInfo[] {
  const columns: ColumnInfo[] = [];
  for (let column = body.start.column; column <= body.end.column; column += 1) {
    const last = lastValueRow.get(column);
    if (last !== undefined && last >= headerRow && last <= body.end.row) {
      columns.push({
        index: column,
        name: worksheet.cellText({ column, row: headerRow }) ?? "",
      });
    }
  }
  return columns;
}
