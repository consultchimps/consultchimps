"use client";

/**
 * The fill handle: the corner of a selection you drag to extend it.
 *
 * Tabulator has no fill handle at all, not even a copying one (the 2026-09-07
 * spike on 6.5.2 confirmed it), so the whole gesture is ours. This module is the
 * part that belongs to the screen: an element on the active range's bottom right
 * corner, a drag that follows the pointer, and an outline showing what the drag
 * would cover. It decides no values and writes nothing. What the cells become is
 * `lib/workspace-series`, which cells are written is `lib/workspace-gesture`, and
 * the writing is the grid's one batched command. Keeping those apart is what
 * lets every rule of a fill be tested without a browser.
 *
 * It is built on Tabulator's public Range API (`getRanges`, `getBounds`), on the
 * cell elements' own rectangles, and on the Record ID the grid writes onto each
 * row, so nothing here reaches into the library's internals.
 *
 * Two details are worth saying out loud.
 *
 * The handle is shown only when exactly one rectangle is selected and at least
 * one of its columns can be written. A fill acts on one rectangle, so offering a
 * handle on a disjoint selection would be offering a gesture that is refused,
 * and a selection sitting only in the Record ID column has nothing to fill.
 *
 * The cell under the pointer is found with `elementsFromPoint`, the plural one,
 * because the handle itself is under the pointer for the whole drag and the
 * singular call would only ever return the handle.
 */
import type { GesturePointer, GestureRect } from "@/lib/workspace-gesture";
import type { RangeComponent, Tabulator } from "tabulator-tables";

/** The handle's side, in pixels: big enough to grab, small enough to sit on a corner. */
const HANDLE_SIZE = 9;

/** What the grid tells the handle about itself, read whenever it is needed. */
export interface FillHandleGeometry {
  /** Record IDs in the order the rows are shown. */
  readonly rows: readonly string[];
  /** Column names in the order they are shown, the Record ID first. */
  readonly columns: readonly string[];
  /** Whether a gesture may write to a column. */
  readonly writable: (column: string) => boolean;
}

/** One finished drag, as grid indices. */
export interface FillHandleDrag {
  /**
   * The rectangles selected when the drag began, the first of them the fill's
   * source. All of them, rather than the one, so the plan is the only thing that
   * decides what a disjoint selection means.
   */
  readonly ranges: readonly GestureRect[];
  readonly pointer: GesturePointer;
}

export interface FillHandleOptions {
  readonly table: Tabulator;
  /** The element the grid was built into. The handle is positioned inside it. */
  readonly container: HTMLElement;
  readonly geometry: () => FillHandleGeometry;
  /** A drag has finished, and this is what it asked for. */
  readonly onFill: (drag: FillHandleDrag) => void;
}

/**
 * One selected range as grid indices, or null when it names a row or a column
 * this grid is not showing.
 *
 * Read from the range's own rows and columns rather than from `getBounds`. The
 * published types describe that as returning two Cell Components, and at 6.5.2
 * it returns the library's internal cells instead, which have none of a
 * component's methods. `getRows` and `getColumns` are components as documented,
 * and they are what the range is defined by anyway: the rows it spans, in the
 * order they are shown, and the columns it spans, in the order they are.
 */
export function rangeRect(
  range: RangeComponent,
  geometry: Pick<FillHandleGeometry, "columns" | "rows">,
): GestureRect | null {
  const { columns, rows } = geometry;
  const spannedRows = range.getRows();
  const spannedColumns = range.getColumns();
  const first = spannedRows[0];
  const last = spannedRows[spannedRows.length - 1];
  const leftmost = spannedColumns[0];
  const rightmost = spannedColumns[spannedColumns.length - 1];
  if (
    first === undefined ||
    last === undefined ||
    leftmost === undefined ||
    rightmost === undefined
  ) {
    return null;
  }
  const top = rows.indexOf(String(first.getIndex()));
  const bottom = rows.indexOf(String(last.getIndex()));
  const left = columns.indexOf(leftmost.getField());
  const right = columns.indexOf(rightmost.getField());
  if (top < 0 || bottom < 0 || left < 0 || right < 0) {
    return null;
  }
  return { top, left, bottom, right };
}

/** Attach the handle to a built grid. The returned function removes it again. */
export function attachFillHandle(options: FillHandleOptions): () => void {
  const { container, geometry, onFill, table } = options;

  const handle = document.createElement("div");
  handle.dataset["testid"] = "workspace-fill-handle";
  // Decoration with a gesture on it, not a control with a name: everything it
  // does is available from the keyboard through the cells themselves.
  handle.setAttribute("aria-hidden", "true");
  handle.style.cssText = [
    "position:absolute",
    `width:${String(HANDLE_SIZE)}px`,
    `height:${String(HANDLE_SIZE)}px`,
    "box-sizing:border-box",
    "border:1px solid var(--color-fd-background)",
    "background:var(--color-fd-primary)",
    "transform:translate(-50%,-50%)",
    "cursor:crosshair",
    // So a touch drag is a drag rather than a scroll.
    "touch-action:none",
    "z-index:11",
    "display:none",
  ].join(";");

  const preview = document.createElement("div");
  preview.setAttribute("aria-hidden", "true");
  preview.style.cssText = [
    "position:absolute",
    "pointer-events:none",
    "border:2px dashed var(--color-fd-primary)",
    "z-index:10",
    "display:none",
  ].join(";");

  container.append(preview, handle);

  let dragging: FillHandleDrag | null = null;
  // Ranges cannot be asked for before the table has finished building, and
  // Tabulator says so rather than answering, so nothing is placed until then.
  let built = false;

  const holder = (): HTMLElement | null =>
    table.element.querySelector(".tabulator-tableholder");

  /** The one selected rectangle, or null when there is not exactly one. */
  const onlyRange = (): RangeComponent | null => {
    if (!built) {
      return null;
    }
    const ranges = table.getRanges();
    return ranges.length === 1 ? (ranges[0] as RangeComponent) : null;
  };

  /** One cell's element, by grid indices, or null when the grid has no such cell. */
  const cellElementAt = (row: number, column: number): HTMLElement | null => {
    const element = table
      .getRows("active")
      [row]?.getCells()
      [column]?.getElement();
    return element instanceof HTMLElement ? element : null;
  };

  const rectOf = (range: RangeComponent): GestureRect | null =>
    rangeRect(range, geometry());

  /** Where the handle belongs, or null when it does not belong anywhere. */
  const corner = (): { left: number; top: number } | null => {
    const range = onlyRange();
    if (range === null || dragging !== null) {
      return null;
    }
    const rect = rectOf(range);
    const { columns, writable } = geometry();
    if (rect === null) {
      return null;
    }
    // A selection with nothing writable in it has nothing to fill.
    let fillable = false;
    for (let index = rect.left; index <= rect.right; index += 1) {
      const column = columns[index];
      if (column !== undefined && writable(column)) {
        fillable = true;
        break;
      }
    }
    if (!fillable) {
      return null;
    }
    const element = cellElementAt(rect.bottom, rect.right);
    const view = holder();
    if (element === null || view === null) {
      return null;
    }
    const cell = element.getBoundingClientRect();
    const visible = view.getBoundingClientRect();
    // Scrolled out of sight: the corner is not on screen, so neither is its
    // handle. Checked against the scrolling element rather than the container,
    // because that is what clips the rows.
    if (
      cell.bottom < visible.top ||
      cell.bottom > visible.bottom ||
      cell.right < visible.left ||
      cell.right > visible.right
    ) {
      return null;
    }
    const box = container.getBoundingClientRect();
    return { left: cell.right - box.left, top: cell.bottom - box.top };
  };

  const place = (): void => {
    const at = corner();
    if (at === null) {
      handle.style.display = "none";
      return;
    }
    handle.style.left = `${String(at.left)}px`;
    handle.style.top = `${String(at.top)}px`;
    handle.style.display = "block";
  };

  /** The cell under the pointer, as grid indices, or null when there is none. */
  const cellAt = (clientX: number, clientY: number): GesturePointer | null => {
    const { columns, rows } = geometry();
    // The plural call, because the handle itself is under the pointer.
    for (const element of document.elementsFromPoint(clientX, clientY)) {
      const cell = element.closest?.(".tabulator-cell");
      if (!(cell instanceof HTMLElement)) {
        continue;
      }
      const field = cell.getAttribute("tabulator-field");
      const recordId = cell
        .closest(".tabulator-row")
        ?.getAttribute("data-record-id");
      if (field === null || recordId === null || recordId === undefined) {
        continue;
      }
      const row = rows.indexOf(recordId);
      const column = columns.indexOf(field);
      if (row < 0 || column < 0) {
        continue;
      }
      return { row, column };
    }
    return null;
  };

  /**
   * Outline everything the drag would cover, from the source's top left corner
   * to the cell under the pointer. Both of those are rendered, which is what
   * makes their rectangles readable: a row further down the table may not be.
   */
  const showPreview = (pointer: GesturePointer, source: GestureRect): void => {
    const first = cellElementAt(source.top, source.left);
    const last = cellElementAt(pointer.row, pointer.column);
    if (first === null || last === null) {
      preview.style.display = "none";
      return;
    }
    const one = first.getBoundingClientRect();
    const other = last.getBoundingClientRect();
    const box = container.getBoundingClientRect();
    const left = Math.min(one.left, other.left) - box.left;
    const top = Math.min(one.top, other.top) - box.top;
    preview.style.left = `${String(left)}px`;
    preview.style.top = `${String(top)}px`;
    preview.style.width = `${String(Math.max(one.right, other.right) - box.left - left)}px`;
    preview.style.height = `${String(Math.max(one.bottom, other.bottom) - box.top - top)}px`;
    preview.style.display = "block";
  };

  const endDrag = (): void => {
    dragging = null;
    preview.style.display = "none";
    place();
  };

  const onPointerDown = (event: PointerEvent): void => {
    if (!built) {
      return;
    }
    const ranges = table.getRanges().map(rectOf);
    const source = ranges[0];
    // A selection this grid cannot read is not a source to drag from. Dropping
    // the range that failed would turn two rectangles into one, which is the
    // opposite of the rule that a fill acts on a single rectangle.
    if (
      source === null ||
      source === undefined ||
      !ranges.every((rect): rect is GestureRect => rect !== null)
    ) {
      return;
    }
    // The grid must not read this as a click on a cell, which would move the
    // selection out from under the drag that is starting.
    event.preventDefault();
    event.stopPropagation();
    dragging = {
      ranges,
      // Until the pointer moves, the drag covers the source and nothing else.
      pointer: { row: source.bottom, column: source.right },
    };
    handle.setPointerCapture(event.pointerId);
  };

  const onPointerMove = (event: PointerEvent): void => {
    if (dragging === null) {
      return;
    }
    // Outside the rows, the drag keeps the last cell it was over, so letting go
    // past the edge fills to the edge rather than cancelling the gesture.
    const pointer = cellAt(event.clientX, event.clientY) ?? dragging.pointer;
    dragging = { ...dragging, pointer };
    const source = dragging.ranges[0];
    if (source !== undefined) {
      showPreview(pointer, source);
    }
  };

  const onPointerUp = (event: PointerEvent): void => {
    const drag = dragging;
    if (drag === null) {
      return;
    }
    handle.releasePointerCapture(event.pointerId);
    endDrag();
    onFill(drag);
  };

  const onPointerCancel = (): void => {
    endDrag();
  };

  handle.addEventListener("pointerdown", onPointerDown);
  handle.addEventListener("pointermove", onPointerMove);
  handle.addEventListener("pointerup", onPointerUp);
  handle.addEventListener("pointercancel", onPointerCancel);

  // Everything that can move the corner the handle sits on. Tabulator reports
  // each of these itself, so the handle never polls.
  const events = [
    "rangeAdded",
    "rangeChanged",
    "rangeRemoved",
    "renderComplete",
    "columnResized",
    "scrollHorizontal",
    "scrollVertical",
  ] as const;
  for (const event of events) {
    // The signatures differ per event and none of the arguments are read.
    table.on(event, place as never);
  }
  const onBuilt = (): void => {
    built = true;
    place();
  };
  table.on("tableBuilt", onBuilt);
  window.addEventListener("resize", place);

  return () => {
    for (const event of events) {
      table.off(event, place as never);
    }
    table.off("tableBuilt", onBuilt);
    window.removeEventListener("resize", place);
    handle.removeEventListener("pointerdown", onPointerDown);
    handle.removeEventListener("pointermove", onPointerMove);
    handle.removeEventListener("pointerup", onPointerUp);
    handle.removeEventListener("pointercancel", onPointerCancel);
    handle.remove();
    preview.remove();
  };
}
