/**
 * Formula cells with no cached value (#162): a workbook saved by a tool that
 * never calculated it holds the formula but not its result. Every operation
 * that reads cell values counts them in `formulaCellsWithoutCachedValues` and
 * names them in one warning built here, so the reader knows to open and save
 * the file in Excel. Nothing computes or invents the missing value.
 */
import { encodeCell, type CellRectangle } from "./model/references.js";

/** How many locations a warning names before it says how many more. */
export const UNCACHED_LOCATIONS_SHOWN = 10;

/** `Sheet!B4` for a zero-based row and column. */
export function cellLocation(
  sheet: string,
  row: number,
  column: number,
): string {
  return `${sheet}!${encodeCell(column, row + 1)}`;
}

/** The `Sheet!B4` locations of the positions inside a zero-based rectangle. */
export function uncachedLocationsWithin(
  sheet: string,
  positions: ReadonlyArray<{ readonly row: number; readonly column: number }>,
  rectangle: CellRectangle | undefined,
): string[] {
  return positions
    .filter(
      (position) =>
        rectangle === undefined ||
        (position.row >= rectangle.startRow &&
          position.row <= rectangle.endRow &&
          position.column >= rectangle.startColumn &&
          position.column <= rectangle.endColumn),
    )
    .map((position) => cellLocation(sheet, position.row, position.column));
}

/**
 * The one warning for formula cells with no cached value, or none when there
 * are none. `effect` says what the operation did with them, such as "they came
 * out blank".
 */
export function uncachedFormulaWarnings(
  locations: readonly string[],
  effect: string,
): string[] {
  const count = locations.length;
  if (count === 0) return [];
  const shown = locations.slice(0, UNCACHED_LOCATIONS_SHOWN).join(", ");
  const more =
    count > UNCACHED_LOCATIONS_SHOWN
      ? `, and ${String(count - UNCACHED_LOCATIONS_SHOWN)} more`
      : "";
  return [
    `${String(count)} formula cell${count === 1 ? " has" : "s have"} no cached value, so ${effect}: ${shown}${more}. Open and recalculate the source workbook in Excel, save it, and run again if these values are required.`,
  ];
}
