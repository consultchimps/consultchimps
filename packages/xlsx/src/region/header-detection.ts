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
 * see - for every populated row, how many values it holds and how many of them
 * sit in a cell merged across columns - and this module answers.
 *
 * The rule. A row is measured against the fullest of the rows that follow it,
 * up to `HEADER_LOOKAHEAD_ROWS` populated rows down: it qualifies as the
 * header when it holds more than half as many values as that fullest row, or
 * at most one value fewer. The header row is the first row that qualifies,
 * provided the rows above it are titles by the evidence below; otherwise it is
 * the first row holding a value, which is what every reader did before.
 * Measuring against the rows that follow, rather than against the fullest row
 * anywhere on the sheet, keeps a wider block further down the sheet from
 * turning a whole table above it into "title rows". The "one fewer" clause
 * matters only on a two-column sheet, where one value is both what a title
 * holds and what a header missing a name holds.
 *
 * The evidence. By count alone a title line cannot be told from a header that
 * leaves most of its columns unnamed: `Name | | ` above `Alice | North | Open`
 * is one value above three, and so is `Report` above `Case_ID | Region |
 * Status`; and `Name` alone above four or more columns is no different, since
 * a count is a likelihood and never a proof. Reading a sparse header as a
 * title would take its first record for the header row, and losing a record
 * silently is the one outcome these readers must never produce, so a row
 * above the header is skipped only on structural evidence that it cannot be
 * the header of the block below:
 *
 * - a blank row lies between it and the header, since a header is never set
 *   apart from its records; or
 * - it is a banner: every value on it sits in a cell merged across two or
 *   more columns, and a cell spanning columns names none of them.
 *
 * And rows that themselves form a table - two adjacent populated rows,
 * neither a banner, the would-be header included - are never titles,
 * whatever wider block follows them. No count enters into it: a one-column
 * list of a header and one record looks exactly like a title and a
 * subtitle, so an unmerged title block of adjacent lines reads as the table
 * it might be, while lines set apart by blank rows and merged banners are
 * skipped. a small summary table above a wide detail block is a
 * table, and the detail block reads as its data, as it always did. When the
 * evidence is missing, the first row holding a value is the header, as it
 * always was, and nothing is lost: the worst case is the old one, a title
 * read as a header, which the inspection shows and `headerRow` overrides. A
 * declared header row is never second-guessed: the rule applies only when no
 * row was declared.
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
 * value is a value: the cell holds something, however unwelcome.
 */

import {
  calendarIsoText,
  isComponentsInRange,
  utcCalendarParts,
} from "../model/calendar.js";
import type {
  CellRange,
  CellRef,
  ColumnIndex,
  RowModel,
  RowNumber,
  WorksheetModel,
} from "../model/types.js";
import type { ColumnInfo } from "./types.js";

/** How many populated rows below a candidate the rule measures it against. */
export const HEADER_LOOKAHEAD_ROWS = 10;

/** What one row of a worksheet holds, as far as the header rule is concerned. */
export interface RowValueCount {
  /**
   * The row, in whatever numbering the caller uses, provided physically
   * adjacent rows number consecutively: the rule reads gaps as blank rows.
   */
  readonly row: RowNumber;
  /** Cells on the row holding a value, within the range the caller examined. */
  readonly values: number;
  /** Of those, the values sitting in a cell merged across two or more columns. */
  readonly bannerValues: number;
}

/** Whether a cell holds nothing, by the definition every reader here shares. */
export function isBlankValue(value: unknown): boolean {
  return value === undefined || value === null || value === "";
}

/**
 * The `row:column` keys of the top-left cells of every merged range that
 * spans two or more columns, in the caller's own numbering. Only that cell
 * holds the merge's value, so a value cell whose key is here is a banner
 * value.
 */
export function bannerAnchors(
  ranges: readonly CellRange[],
): ReadonlySet<string> {
  const anchors = new Set<string>();
  for (const range of ranges) {
    if (range.end.column > range.start.column) {
      anchors.add(cellKey(range.start.row, range.start.column));
    }
  }
  return anchors;
}

/** The key `bannerAnchors` uses, for callers building their own set. */
export function cellKey(row: number, column: number): string {
  return `${String(row)}:${String(column)}`;
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
 * Whether a populated row above the header is a title line by the evidence
 * the module comment describes: set apart from the header by a blank row, or
 * a banner whose every value sits in a cell merged across columns. Exported
 * for the same reason as `qualifiesAsHeader`.
 */
export function isTitleLine(
  line: RowValueCount,
  blankRowBetween: boolean,
): boolean {
  return (
    blankRowBetween || (line.values > 0 && line.bannerValues === line.values)
  );
}

/**
 * Whether rows, in row order, contain a table: two physically adjacent
 * populated rows, neither a banner (a cell spanning columns is no record).
 * A title block is banners, or lines set apart from each other and from the
 * header by blank rows; two adjacent populated rows are a table, and a table
 * is never skipped as titles, whatever fuller block follows it. The row that would become the
 * header is part of the question: a three-value line directly above it may
 * be a three-name header over its first record, which is the sparse-header
 * case nothing can tell from a title, so it is read as one.
 */
/** A row that could be a record: populated, and not a banner. */
function isTableRow(row: RowValueCount): boolean {
  return row.values > 0 && row.bannerValues < row.values;
}

function formsTable(rows: readonly RowValueCount[]): boolean {
  for (let index = 1; index < rows.length; index += 1) {
    const above = rows[index - 1]!;
    const below = rows[index]!;
    if (below.row === above.row + 1 && isTableRow(above) && isTableRow(below)) {
      return true;
    }
  }
  return false;
}

/**
 * Whether every populated row above `populated[headerIndex]` is a title line
 * by the evidence, and none of them form a table with each other or with the
 * header.
 */
function maySkipTitles(
  populated: readonly RowValueCount[],
  headerIndex: number,
): boolean {
  const header = populated[headerIndex]!;
  for (let index = 0; index < headerIndex; index += 1) {
    const line = populated[index]!;
    // Populated rows number consecutively when no blank row lies between
    // them, so a larger gap in row numbers than in positions is a blank row.
    const blankRowBetween = header.row - line.row > headerIndex - index;
    if (!isTitleLine(line, blankRowBetween)) {
      return false;
    }
  }
  return !formsTable(populated.slice(0, headerIndex + 1));
}

/**
 * The header row among `counts`, or undefined when no row holds a value.
 *
 * `counts` is read in the order given, which callers keep top to bottom: the
 * answer is the first row that qualifies against the fullest of itself and
 * the populated rows that follow it within the lookahead, provided the rows
 * above it are titles by the evidence; otherwise the first populated row.
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
    return maySkipTitles(populated, index) ? candidate.row : first.row;
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
  /**
   * Every stored row inside the used range that holds a value, in row order.
   * Rows holding nothing are left out; the rule reads blank rows from the
   * gaps in row numbers, so nothing is kept per empty row.
   */
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
  const banners = bannerAnchors(worksheet.mergedRanges());
  for (const row of rows) {
    if (row.number < used.start.row || row.number > used.end.row) {
      continue;
    }
    let values = 0;
    let bannerValues = 0;
    for (const cell of row.cells) {
      const column = cell.ref.column;
      if (column < used.start.column || column > used.end.column) {
        continue;
      }
      if (isBlankValue(worksheet.cellValue(cell.ref))) {
        continue;
      }
      values += 1;
      if (banners.has(cellKey(row.number, column))) {
        bannerValues += 1;
      }
      const last = lastValueRow.get(column);
      if (last === undefined || last < row.number) {
        lastValueRow.set(column, row.number);
      }
    }
    if (values > 0) {
      counts.push({ bannerValues, row: row.number, values });
    }
  }
  // Stored order is document order, which a hand-written part may leave
  // unsorted; the rule reads rows top to bottom.
  counts.sort((left, right) => left.row - right.row);
  return { counts, lastValueRow };
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
        name: headerCellName(worksheet, { column, row: headerRow }),
      });
    }
  }
  return columns;
}

/**
 * The name a header cell gives its column, spelled the way the SheetJS table
 * reader spells it, so the columns an inspection reports are the columns a
 * consolidation produces whatever the cell's type: a number by its value
 * rather than its stored token (`1E3` is `1000`), a boolean as `true` or
 * `false`, a date as the ISO timestamp every reader here writes dates as, an
 * error by its text, and a blank cell as an empty name for `uniqueHeaders`
 * to fill in.
 */
export function headerCellName(
  worksheet: WorksheetModel,
  ref: CellRef,
): string {
  const value = worksheet.cellValue(ref);
  if (isBlankValue(value)) {
    return "";
  }
  if (value instanceof Date) {
    const parts = utcCalendarParts(value);
    return isComponentsInRange(parts)
      ? calendarIsoText(parts)
      : (worksheet.cellText(ref) ?? "");
  }
  return String(value);
}
