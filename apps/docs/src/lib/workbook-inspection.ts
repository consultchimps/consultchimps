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
 * A sample value as the report prints it: text quoted, numbers and booleans
 * bare, which is the spelling a column mapping's JSON uses for the same values.
 *
 * The quoting is the one thing this adds, and it is not decoration. The
 * inspection keeps the number 1 and the text "1" apart deliberately, because a
 * mapping review has to see which is which; printing both as `1` would throw
 * that away at the last step. `JSON.stringify` also escapes a quote or a
 * newline inside the text, so a value cannot fake the delimiter around it.
 * Nothing is reformatted, rounded, or relabelled beyond that.
 *
 * The inspection drops empty values before it samples, so a `null` cannot
 * reach here. It still prints as nothing rather than as the text "null" that
 * `String` would otherwise put into the report.
 */
export function sampleValueText(value: SampleValue): string {
  if (value === null) {
    return "";
  }
  // Numbers go through String rather than JSON.stringify, which writes a
  // non-finite number as "null": there is no JSON literal for one, and the
  // report must not print a stored value as though it were absent.
  return typeof value === "string" ? JSON.stringify(value) : String(value);
}
