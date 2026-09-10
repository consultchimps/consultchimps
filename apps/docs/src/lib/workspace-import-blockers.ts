/**
 * Why a source cannot be imported, decided once.
 *
 * Two surfaces need this answer and neither may reach the other's copy of it:
 * the page has to leave a source unticked and say why, and the worker has to
 * refuse it whatever the page allowed. The page cannot import the reader that
 * knows - the reader pulls in a spreadsheet engine, and the page must not - so
 * the rule lives in its own module that both can hold.
 *
 * What is shared is the condition, not the sentence. Each surface says it in
 * its own voice: the worker is refusing a request and names the file and the
 * worksheet, the page is annotating a row the visitor is already looking at.
 * Only the decision has to agree, and a decision made twice is the shape that
 * lets a form offer a tick the import will not honour.
 */

/** The cell counts a worksheet report carries about one read of one source. */
export interface ImportSourceCells {
  /** Cells holding a formula the workbook carries no calculated value for. */
  readonly uncachedFormulaCells: number;
  /** Cells holding an error value, typed in or left by a failed formula. */
  readonly errorCells: number;
}

/**
 * A condition that stops an import, and how many cells are in it.
 *
 * `uncalculated-formulas` reads as empty and would import holes.
 * `error-values` reads as a number - the internal code Excel numbers each error
 * by - and would import data that was never in the worksheet.
 */
export interface ImportBlocker {
  readonly kind: "uncalculated-formulas" | "error-values";
  readonly cells: number;
}

/**
 * Every condition stopping this source, in a fixed order, or nothing when it
 * can be imported.
 *
 * The order is fixed rather than taken from the counts so that a source failing
 * both conditions is always refused with the same one, and a page listing both
 * always lists them the same way round.
 */
export function importBlockers(
  source: ImportSourceCells,
): readonly ImportBlocker[] {
  const blockers: ImportBlocker[] = [];
  if (source.uncachedFormulaCells > 0) {
    blockers.push({
      kind: "uncalculated-formulas",
      cells: source.uncachedFormulaCells,
    });
  }
  if (source.errorCells > 0) {
    blockers.push({ kind: "error-values", cells: source.errorCells });
  }
  return blockers;
}

/** Whether a source can be imported at all. */
export function isImportable(source: ImportSourceCells): boolean {
  return importBlockers(source).length === 0;
}

/** "1 cell" or "4 cells", for a sentence either surface builds around it. */
export function cellCountText(cells: number): string {
  return cells === 1 ? "1 cell" : `${cells} cells`;
}
