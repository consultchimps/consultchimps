/**
 * Tier-1 correctness utility: clear cached results that were computed over rows
 * a split is about to remove.
 *
 * A values-only split replaces each formula with its cached result. When the
 * formula aggregated rows belonging to other groups -- `SUM(Data!D4:D9)` on a
 * summary sheet, a totals row, a footer block -- the cached result is the whole
 * workbook's answer, and baking it into one group's output presents another
 * group's total as this group's. Clearing the cached value first makes the
 * conversion produce a formatted blank cell, which the values-only report
 * already surfaces as a formula that lost its value.
 *
 * Conservative is correct here: an over-cleared cell is a visible blank the
 * caller is warned about, an under-cleared one is a wrong number nobody sees.
 *
 * Package access goes through L0, so this pass writes with the same
 * deterministic rules as a model edit and leaves untouched parts byte-identical.
 */
import { readWorkbookSheetsFrom } from "../excel-tables.js";
import { decodeXmlText } from "../model/xml.js";
import { WorkbookPackage } from "../package/index.js";

const CELL_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?c\b[^>]*?(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?c\s*>)/gu;
const CELL_FORMULA_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*?(?:\/\s*>|>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?f\s*>)/u;
const CACHED_VALUE_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?(?:v|is)\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?(?:v|is)\s*>)/gu;
const CELL_REFERENCE_PATTERN = /(?:^|\s)r\s*=\s*(?:"([^"]+)"|'([^']+)')/u;
const CELL_ROW_PATTERN = /^\$?[A-Za-z]{1,3}\$?(\d+)$/u;
const QUOTED_STRING_PATTERN = /"[^"]*"/gu;

/**
 * An A1 reference, optionally sheet-qualified, in one of the three shapes a
 * formula can name rows with: a cell or cell range, a whole-column range, or a
 * whole-row range. The lookbehind and lookahead are the guard
 * `evaluateFormulaGuard` uses to keep function names such as `LOG10(` and
 * defined names from reading as references.
 */
const A1_REFERENCE_PATTERN = new RegExp(
  "(?<![A-Za-z0-9_.$])" +
    "(?:(?:'((?:[^']|'')*)'|([A-Za-z_][A-Za-z0-9_.]*))!)?" +
    "(?:(\\$?[A-Za-z]{1,3}\\$?\\d+)(?::(\\$?[A-Za-z]{1,3}\\$?\\d+))?" +
    "|\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3}" +
    "|\\$?(\\d+):\\$?(\\d+))" +
    "(?![\\dA-Za-z_(])",
  "gu",
);

export interface BlankedCachedFormula {
  /** The cell whose cached result was cleared, e.g. `B2`. */
  cell: string;
  /** The sheet the cell lives on, by name. */
  sheet: string;
}

export interface StaleCachedValueReport {
  blankedCells: BlankedCachedFormula[];
  bytes: Uint8Array;
}

/** Rows a split will delete, in source row numbers, keyed by worksheet part. */
export type DeletedRowsByPart = ReadonlyMap<string, ReadonlySet<number>>;

function cellReference(cellXml: string): string | undefined {
  const match = CELL_REFERENCE_PATTERN.exec(cellXml);
  return match?.[1] ?? match?.[2];
}

function cellRow(reference: string): number | undefined {
  const row = Number(CELL_ROW_PATTERN.exec(reference)?.[1]);
  return Number.isInteger(row) ? row : undefined;
}

function intersects(
  deletedRows: ReadonlySet<number>,
  first: number,
  last: number,
): boolean {
  for (const row of deletedRows) {
    if (row >= first && row <= last) {
      return true;
    }
  }
  return false;
}

/**
 * Whether a formula names any row this split removes. Cross-sheet references
 * count: a summary total over another sheet's rows goes stale exactly when that
 * sheet loses rows.
 */
function referencesDeletedRows(
  formula: string,
  ownSheet: string,
  deletedRowsBySheet: ReadonlyMap<string, ReadonlySet<number>>,
): boolean {
  // Quoted text can hold anything that looks like a reference; the same trick
  // `evaluateFormulaGuard` uses removes it before scanning. Entities are
  // resolved first so a fixture that escapes its string delimiters is stripped
  // as reliably as one that does not.
  const scannable = decodeXmlText(formula).replace(QUOTED_STRING_PATTERN, "");
  for (const match of scannable.matchAll(A1_REFERENCE_PATTERN)) {
    const [, quotedSheet, bareSheet, first, last, firstRow, lastRow] = match;
    const sheet = (quotedSheet?.replaceAll("''", "'") ?? bareSheet ?? ownSheet)
      .trim()
      .toLowerCase();
    const deletedRows = deletedRowsBySheet.get(sheet);
    if (!deletedRows || deletedRows.size === 0) {
      continue;
    }
    if (first !== undefined) {
      const start = cellRow(first);
      const end = last === undefined ? start : cellRow(last);
      if (
        start !== undefined &&
        end !== undefined &&
        intersects(deletedRows, Math.min(start, end), Math.max(start, end))
      ) {
        return true;
      }
      continue;
    }
    if (firstRow !== undefined && lastRow !== undefined) {
      const start = Number(firstRow);
      const end = Number(lastRow);
      if (intersects(deletedRows, Math.min(start, end), Math.max(start, end))) {
        return true;
      }
      continue;
    }
    // A whole-column reference covers every row, so any deletion touches it.
    return true;
  }
  return false;
}

/**
 * Clear the cached result of every formula whose references reach into rows the
 * split deletes, so a subsequent values-only conversion produces a blank cell
 * rather than an aggregate computed over rows the recipient never receives.
 *
 * Call this *before* the values conversion, with row numbers as they are in the
 * source workbook. Formulas that sit on a deleted row of their own sheet are
 * left alone: the cell is leaving anyway. Structured-table references, defined
 * names, 3-D sheet ranges and shared-formula slaves that carry no text of their
 * own are not resolved and therefore not detected.
 */
export async function blankStaleCachedFormulas(
  workbookBytes: Uint8Array,
  deletedRowsBySheet: DeletedRowsByPart,
): Promise<StaleCachedValueReport> {
  const hasDeletions = [...deletedRowsBySheet.values()].some(
    (rows) => rows.size > 0,
  );
  if (!hasDeletions) {
    return { blankedCells: [], bytes: workbookBytes };
  }

  const workbookPackage = await WorkbookPackage.load(workbookBytes);
  const identities = readWorkbookSheetsFrom(workbookPackage);
  const deletedRowsByName = new Map<string, ReadonlySet<number>>();
  for (const identity of identities) {
    const rows = deletedRowsBySheet.get(identity.worksheetPart);
    if (rows && rows.size > 0) {
      deletedRowsByName.set(identity.name.trim().toLowerCase(), rows);
    }
  }

  const blankedCells: BlankedCachedFormula[] = [];
  for (const identity of identities) {
    const worksheetXml = workbookPackage.readText(identity.worksheetPart);
    if (worksheetXml === undefined) {
      continue;
    }
    const ownDeletedRows = deletedRowsBySheet.get(identity.worksheetPart);
    const rewritten = worksheetXml.replace(CELL_PATTERN, (cellXml) => {
      const formula = CELL_FORMULA_PATTERN.exec(cellXml)?.[1];
      if (!formula) {
        return cellXml;
      }
      const reference = cellReference(cellXml);
      const row = reference === undefined ? undefined : cellRow(reference);
      if (row !== undefined && ownDeletedRows?.has(row) === true) {
        return cellXml;
      }
      if (
        !referencesDeletedRows(
          formula,
          identity.name.trim().toLowerCase(),
          deletedRowsByName,
        )
      ) {
        return cellXml;
      }
      const blanked = cellXml.replace(CACHED_VALUE_PATTERN, "");
      if (blanked === cellXml) {
        return cellXml;
      }
      blankedCells.push({
        cell: reference ?? "unknown cell",
        sheet: identity.name,
      });
      return blanked;
    });
    if (rewritten !== worksheetXml) {
      workbookPackage.writeText(identity.worksheetPart, rewritten);
    }
  }

  if (blankedCells.length === 0) {
    return { blankedCells, bytes: workbookBytes };
  }
  return { blankedCells, bytes: await workbookPackage.save() };
}
