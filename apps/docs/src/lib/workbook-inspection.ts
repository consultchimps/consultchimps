/**
 * The wording the workbook inspector puts around a `sheets.inspect`
 * description.
 *
 * These are plain functions over the fields the operation reports, kept out of
 * the component so they can be tested without a DOM and reused by whatever
 * else renders a description later. Nothing here invents a fact about a
 * workbook: every phrase is assembled from counts and names the inspection
 * itself produced, which is the same rule the tool pages follow when they
 * render an operation's warnings verbatim.
 */

// Type-only: the runtime module is loaded inside the operation worker, so
// importing these names costs the page nothing.
import type {
  WorkbookColumnDescription,
  WorkbookSheetDescription,
  WorksheetVisibility,
} from "@consultchimps/xlsx/bytes";

/** One of the bounded stored values the inspection samples from a column. */
export type SampleValue = WorkbookColumnDescription["sampleValues"][number];

/** The fields the callout below needs, so a caller can pass anything shaped so. */
export type VisibilitySummary = Pick<WorkbookSheetDescription, "visibility">;

/** The fields the per-worksheet summary line reads. */
export type WorksheetSummary = Pick<
  WorkbookSheetDescription,
  "columnCount" | "dataRowCount" | "headerRow" | "rowCount"
>;

function counted(count: number, noun: string): string {
  return `${count} ${noun}${count === 1 ? "" : "s"}`;
}

function isOrAre(count: number): string {
  return count === 1 ? "is" : "are";
}

/**
 * The badge a worksheet carries in the report, or undefined for an ordinary
 * visible one. A "Visible" badge on every row would be noise; the two states
 * worth interrupting a reader for are the ones an operation skips by default.
 */
export function visibilityBadge(
  visibility: WorksheetVisibility,
): string | undefined {
  switch (visibility) {
    case "hidden": {
      return "Hidden";
    }
    case "very-hidden": {
      return "Very hidden";
    }
    default: {
      return undefined;
    }
  }
}

/**
 * The headline shown above the worksheet list when the description covers a
 * hidden worksheet, or undefined when every described worksheet is visible.
 *
 * The two hidden states are counted separately because Excel treats them
 * differently: one is a click away in the tab bar, the other is reachable only
 * from the VBA editor, and a reader deciding whether a column really is
 * missing needs to know which kind they are looking at.
 */
export function hiddenWorksheetCallout(
  sheets: readonly VisibilitySummary[],
): string | undefined {
  const hidden = sheets.filter((sheet) => sheet.visibility === "hidden").length;
  const veryHidden = sheets.filter(
    (sheet) => sheet.visibility === "very-hidden",
  ).length;

  if (hidden === 0 && veryHidden === 0) {
    return undefined;
  }

  let phrase: string;
  if (hidden > 0 && veryHidden > 0) {
    phrase = `${counted(hidden, "worksheet")} ${isOrAre(hidden)} hidden and ${veryHidden} ${isOrAre(veryHidden)} very hidden`;
  } else if (hidden > 0) {
    phrase = `${counted(hidden, "worksheet")} ${isOrAre(hidden)} hidden`;
  } else {
    phrase = `${counted(veryHidden, "worksheet")} ${isOrAre(veryHidden)} very hidden`;
  }

  return `${phrase}. A hidden worksheet can be unhidden from Excel's tab bar, a very hidden one only from the VBA editor. Most operations skip both unless you ask for them`;
}

/**
 * One worksheet's shape in a single line: the used range, the header row an
 * operation would key on, and how many rows sit below it.
 *
 * A worksheet the inspection found no content in reports that rather than a
 * row of zeros, because "0 rows" beside "no header row found" reads as a
 * failure to look rather than as an empty tab.
 */
export function worksheetSummary(sheet: WorksheetSummary): string {
  if (sheet.rowCount === 0) {
    return "No cells with content";
  }
  return [
    counted(sheet.rowCount, "row"),
    counted(sheet.columnCount, "column"),
    sheet.headerRow === undefined
      ? "no header row found"
      : `header row ${sheet.headerRow}`,
    counted(sheet.dataRowCount, "data row"),
  ].join(" · ");
}

/**
 * A sample value as the report prints it. The description already promises
 * stored values with nothing type-inferred, so this only turns a number or a
 * boolean into text; it never reformats, rounds, or relabels one.
 *
 * The inspection drops empty values before it samples, so a `null` cannot
 * reach here. It still maps to an empty string rather than to the text "null"
 * that `String` would otherwise print into the report.
 */
export function sampleValueText(value: SampleValue): string {
  return value === null ? "" : String(value);
}
