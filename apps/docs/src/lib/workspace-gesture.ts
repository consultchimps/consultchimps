/**
 * What a paste or a fill would write, decided before anything is sent.
 *
 * A gesture is one movement that changes many cells, so the whole of it is
 * planned first: which cells, what goes in them, or one sentence saying why the
 * gesture is refused and nothing is written at all. Planning it here, as pure
 * functions over a grid described as indices and text, is what lets every rule
 * below be pinned by a test with no browser, no grid library, and no database.
 *
 * A plan is total: it either writes or refuses. There is no third outcome where
 * part of a gesture lands, because the visitor read the whole movement as one
 * action. A refusal a person can act on beats a partial write they have to
 * reconstruct, so the refusals are:
 *
 * - **More than one selected rectangle.** A clipboard gesture and a fill each
 *   work on one rectangle, the way a spreadsheet refuses a multiple selection,
 *   rather than quietly acting on whichever was selected last.
 * - **Past the grid's edge.** A paste that needs rows or columns the table does
 *   not have is refused whole and says how many are missing. Adding records is
 *   not something this grid does, and truncating a block reads as a paste that
 *   worked.
 * - **The Record ID.** It is assigned once (ADR 0003 Decision 8), so it is
 *   never a write target. Writing the other columns and dropping that one would
 *   read as a block that landed.
 * - **A source cell with a refusal still standing.** Its value on screen is the
 *   one the database kept, not the one the visitor asked for, so filling from it
 *   would spread a value nobody chose.
 * - **More cells than one step applies.** `WORKSPACE_MAX_GESTURE_CELLS` is the
 *   cap, so a runaway drag cannot enqueue a transaction the tab will not finish.
 *
 * The values themselves come from `workspace-series` for a fill and from the
 * clipboard block for a paste. Both are text, and the library's one conversion
 * point decides what each column will hold, so nothing here validates a value.
 */
import { fillLine } from "./workspace-series";
import { WORKSPACE_MAX_GESTURE_CELLS } from "./workspace-protocol";

/** One column of the grid, as a gesture needs to know it. */
export interface GestureColumn {
  /** The column name, which is the field the grid addresses it by. */
  readonly field: string;
  /** Whether a gesture may write here. False for the Record ID. */
  readonly writable: boolean;
  /**
   * Whether a fill may read a series from this column's values. False for a
   * foreign key, whose values name records rather than counting anything.
   */
  readonly series: boolean;
}

/** The grid a gesture acts on: its columns, its records, and its values. */
export interface GestureGrid {
  /** Left to right, as shown. The Record ID is the first. */
  readonly columns: readonly GestureColumn[];
  /** Top to bottom, as shown. */
  readonly recordIds: readonly string[];
  /** The text of one cell as the database last confirmed it. */
  readonly text: (recordId: string, field: string) => string;
  /** Whether a refusal is still standing against one cell. */
  readonly refused?: (recordId: string, field: string) => boolean;
}

/** A rectangle of the grid, in inclusive row and column indices. */
export interface GestureRect {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

/** One cell a gesture writes, with the text to write into it. */
export interface PlannedWrite {
  readonly recordId: string;
  readonly column: string;
  /** The text to store. Empty means an empty cell. */
  readonly value: string;
}

export type GesturePlan =
  | {
      readonly kind: "writes";
      readonly writes: readonly PlannedWrite[];
      /**
       * What the gesture covered, for the selection to follow: the pasted block,
       * or a fill's source and everything it filled. Null when it covered
       * nothing. A spreadsheet leaves the selection on what a gesture just did,
       * which is also where the next drag starts from.
       */
      readonly covered: GestureRect | null;
    }
  | { readonly kind: "refused"; readonly reason: string };

/** What a gesture is called in the sentence explaining its refusal. */
export type GestureKind = "paste" | "fill";

/** One rectangle at a time, for a copy, a paste, and a fill alike. */
export const ONE_RECTANGLE_ONLY =
  "That works on one selected rectangle at a time. Select a single rectangle and try again";

/** Refused because the workspace was busy, so the gesture was never sent. */
export function busyRefusal(kind: GestureKind): string {
  return `That ${kind} was not made because the workspace was busy. Make it again now the workspace is ready`;
}

function edgeRefusal(kind: GestureKind, missing: number, axis: string): string {
  const plural = missing === 1 ? axis : `${axis}s`;
  const done = kind === "paste" ? "pasted" : "filled";
  return `That ${kind} needs ${String(missing)} more ${plural} than this table has, so nothing was ${done}. Records are not added by a ${kind}, so start nearer the top left or use a smaller block`;
}

function recordIdRefusal(kind: GestureKind): string {
  return `The Record ID is assigned once when a record is created, so a ${kind} cannot write to it. Start the ${kind} in one of the columns beside it`;
}

function capRefusal(kind: GestureKind, cells: number): string {
  return `That ${kind} covers ${String(cells)} cells, and one step applies at most ${String(WORKSPACE_MAX_GESTURE_CELLS)}, so nothing was changed. Use a smaller range`;
}

const REFUSED_SOURCE =
  "That fill would spread a value the workspace refused, which is not the value the cell holds. Deal with the refused cell first, then fill again";

function refused(reason: string): GesturePlan {
  return { kind: "refused", reason };
}

function writes(
  planned: readonly PlannedWrite[],
  covered: GestureRect | null = null,
): GesturePlan {
  return { kind: "writes", writes: planned, covered };
}

/** The smallest rectangle holding both, which is what a fill leaves selected. */
function union(one: GestureRect, other: GestureRect): GestureRect {
  return {
    top: Math.min(one.top, other.top),
    left: Math.min(one.left, other.left),
    bottom: Math.max(one.bottom, other.bottom),
    right: Math.max(one.right, other.right),
  };
}

/** The one rectangle a gesture may act on, or a refusal. */
function onlyRect(ranges: readonly GestureRect[]): GestureRect | null {
  return ranges.length === 1 ? (ranges[0] as GestureRect) : null;
}

function height(rect: GestureRect): number {
  return rect.bottom - rect.top + 1;
}

function width(rect: GestureRect): number {
  return rect.right - rect.left + 1;
}

/**
 * Everything both gestures check once a target rectangle is known: it lies
 * inside the grid, every column of it may be written, and it is not more than
 * one step applies.
 */
function checkTarget(
  grid: GestureGrid,
  target: GestureRect,
  kind: GestureKind,
): string | null {
  const missingRows = target.bottom - (grid.recordIds.length - 1);
  if (missingRows > 0) {
    return edgeRefusal(kind, missingRows, "row");
  }
  const missingColumns = target.right - (grid.columns.length - 1);
  if (missingColumns > 0) {
    return edgeRefusal(kind, missingColumns, "column");
  }
  for (let index = target.left; index <= target.right; index += 1) {
    if ((grid.columns[index] as GestureColumn).writable !== true) {
      return recordIdRefusal(kind);
    }
  }
  const cells = height(target) * width(target);
  if (cells > WORKSPACE_MAX_GESTURE_CELLS) {
    return capRefusal(kind, cells);
  }
  return null;
}

export interface PastePlanOptions {
  readonly grid: GestureGrid;
  /** The selection, which must be one rectangle. */
  readonly ranges: readonly GestureRect[];
  /** The clipboard block, rows top to bottom and cells left to right. */
  readonly block: ReadonlyArray<readonly string[]>;
}

/**
 * Plan a paste of a clipboard block into the selection.
 *
 * The shape rules, which are a spreadsheet's:
 *
 * - A single copied value fills the whole selection.
 * - A block whose height and width both divide the selection tiles across it.
 *   A selection exactly the block's size is the same rule with one tile.
 * - Anything else is written from the anchor, the selection's top left,
 *   overflowing the selection when the block is larger and covering part of it
 *   when the block is smaller.
 *
 * A block with a short row is read as ending in empty cells, which is how a
 * spreadsheet reads one, rather than refusing a clipboard a person did not
 * compose by hand.
 */
export function planPaste(options: PastePlanOptions): GesturePlan {
  const { grid, block } = options;
  const selection = onlyRect(options.ranges);
  if (selection === null) {
    return refused(ONE_RECTANGLE_ONLY);
  }
  const blockHeight = block.length;
  const blockWidth = Math.max(0, ...block.map((row) => row.length));
  if (blockHeight === 0 || blockWidth === 0) {
    // Nothing on the clipboard is nothing to do, not a failure to explain.
    return writes([]);
  }

  const tiles =
    height(selection) % blockHeight === 0 &&
    width(selection) % blockWidth === 0;
  const target: GestureRect = tiles
    ? selection
    : {
        top: selection.top,
        left: selection.left,
        bottom: selection.top + blockHeight - 1,
        right: selection.left + blockWidth - 1,
      };

  const problem = checkTarget(grid, target, "paste");
  if (problem !== null) {
    return refused(problem);
  }

  const planned: PlannedWrite[] = [];
  for (let row = target.top; row <= target.bottom; row += 1) {
    const recordId = grid.recordIds[row] as string;
    for (let column = target.left; column <= target.right; column += 1) {
      const blockRow = block[(row - target.top) % blockHeight] ?? [];
      planned.push({
        recordId,
        column: (grid.columns[column] as GestureColumn).field,
        value: blockRow[(column - target.left) % blockWidth] ?? "",
      });
    }
  }
  return writes(planned, target);
}

/** A cell the pointer is over, as grid indices. */
export interface GesturePointer {
  readonly row: number;
  readonly column: number;
}

export interface FillPlanOptions {
  readonly grid: GestureGrid;
  /** The selection, which must be one rectangle: the fill's source. */
  readonly ranges: readonly GestureRect[];
  /** Where the drag has reached. Clamped into the grid. */
  readonly pointer: GesturePointer;
}

/**
 * Plan a fill from the selected rectangle to where the drag has reached.
 *
 * One axis at a time, the one the pointer has moved furthest along beyond the
 * source, with a tie going to the vertical because a column is the fill people
 * mean. Down and right continue the series forwards, up and left continue it
 * backwards, from the same rule: the source occupies indices 0 to n-1 along the
 * axis and a target cell is asked for by its own index, which is negative
 * behind the source and n or greater ahead of it.
 *
 * Each line is filled on its own: a column of the source for a vertical fill, a
 * row of it for a horizontal one. A line that crosses a column whose values
 * name records rather than counting them copies instead of extending, because a
 * Record ID with a trailing integer is not a counter.
 */
export function planFill(options: FillPlanOptions): GesturePlan {
  const { grid, pointer } = options;
  const source = onlyRect(options.ranges);
  if (source === null) {
    return refused(ONE_RECTANGLE_ONLY);
  }

  const row = clamp(pointer.row, 0, grid.recordIds.length - 1);
  const column = clamp(pointer.column, 0, grid.columns.length - 1);
  const below = row - source.bottom;
  const above = source.top - row;
  const right = column - source.right;
  const left = source.left - column;
  const vertical = Math.max(below, above);
  const horizontal = Math.max(right, left);
  if (vertical <= 0 && horizontal <= 0) {
    // The drag never left the source, so there is nothing to fill.
    return writes([]);
  }

  const target: GestureRect =
    // A tie goes to the vertical: a column is the fill a drag usually means.
    vertical >= horizontal
      ? {
          top: below > 0 ? source.bottom + 1 : row,
          bottom: below > 0 ? row : source.top - 1,
          left: source.left,
          right: source.right,
        }
      : {
          top: source.top,
          bottom: source.bottom,
          left: right > 0 ? source.right + 1 : column,
          right: right > 0 ? column : source.left - 1,
        };

  const problem = checkTarget(grid, target, "fill");
  if (problem !== null) {
    return refused(problem);
  }
  if (sourceWasRefused(grid, source)) {
    return refused(REFUSED_SOURCE);
  }

  return writes(
    vertical >= horizontal
      ? fillDown(grid, source, target)
      : fillAcross(grid, source, target),
    // The source too: a fill leaves the whole of what it made selected, so the
    // handle is at the corner of it and a second drag carries on from there.
    union(source, target),
  );
}

function clamp(value: number, low: number, high: number): number {
  return Math.min(Math.max(value, low), high);
}

/** Whether any cell of the source still has a refusal standing against it. */
function sourceWasRefused(grid: GestureGrid, source: GestureRect): boolean {
  const wasRefused = grid.refused;
  if (wasRefused === undefined) {
    return false;
  }
  for (let row = source.top; row <= source.bottom; row += 1) {
    const recordId = grid.recordIds[row] as string;
    for (let column = source.left; column <= source.right; column += 1) {
      if (wasRefused(recordId, (grid.columns[column] as GestureColumn).field)) {
        return true;
      }
    }
  }
  return false;
}

/** One line per column, extended along the rows. */
function fillDown(
  grid: GestureGrid,
  source: GestureRect,
  target: GestureRect,
): PlannedWrite[] {
  const rows = indicesBetween(target.top, target.bottom);
  const planned: PlannedWrite[] = [];
  for (let column = source.left; column <= source.right; column += 1) {
    const { field, series } = grid.columns[column] as GestureColumn;
    const line = indicesBetween(source.top, source.bottom).map((row) =>
      grid.text(grid.recordIds[row] as string, field),
    );
    const values = fillLine(
      line,
      rows.map((row) => row - source.top),
      { copyOnly: !series },
    );
    rows.forEach((row, position) => {
      planned.push({
        recordId: grid.recordIds[row] as string,
        column: field,
        value: values[position] as string,
      });
    });
  }
  return sortedByCell(planned, grid);
}

/** One line per row, extended along the columns. */
function fillAcross(
  grid: GestureGrid,
  source: GestureRect,
  target: GestureRect,
): PlannedWrite[] {
  const columns = indicesBetween(target.left, target.right);
  // A line that crosses a column whose values name records copies throughout: a
  // series read across mixed columns is not a series anyone asked for.
  const copyOnly = indicesBetween(
    Math.min(source.left, target.left),
    Math.max(source.right, target.right),
  ).some((column) => !(grid.columns[column] as GestureColumn).series);
  const planned: PlannedWrite[] = [];
  for (let row = source.top; row <= source.bottom; row += 1) {
    const recordId = grid.recordIds[row] as string;
    const line = indicesBetween(source.left, source.right).map((column) =>
      grid.text(recordId, (grid.columns[column] as GestureColumn).field),
    );
    const values = fillLine(
      line,
      columns.map((column) => column - source.left),
      { copyOnly },
    );
    columns.forEach((column, position) => {
      planned.push({
        recordId,
        column: (grid.columns[column] as GestureColumn).field,
        value: values[position] as string,
      });
    });
  }
  return planned;
}

function indicesBetween(from: number, to: number): number[] {
  return Array.from({ length: to - from + 1 }, (_unused, step) => from + step);
}

/**
 * Row by row, then column by column, which is the order the grid reads in and
 * the order a report about the writes should follow. A vertical fill builds its
 * writes a column at a time, so they are put back into reading order here.
 */
function sortedByCell(
  planned: readonly PlannedWrite[],
  grid: GestureGrid,
): PlannedWrite[] {
  const rowOf = new Map(
    grid.recordIds.map((recordId, index) => [recordId, index]),
  );
  const columnOf = new Map(
    grid.columns.map((column, index) => [column.field, index]),
  );
  return [...planned].sort(
    (one, other) =>
      (rowOf.get(one.recordId) ?? 0) - (rowOf.get(other.recordId) ?? 0) ||
      (columnOf.get(one.column) ?? 0) - (columnOf.get(other.column) ?? 0),
  );
}
