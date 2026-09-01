import { describe, expect, it } from "vitest";

import {
  COLUMN_PREVIEW_LIMIT,
  columnPreviewNote,
  hiddenWorksheetCallout,
  sampleValueText,
  visibilityBadge,
  worksheetSummary,
  type VisibilitySummary,
} from "./workbook-inspection";

function sheets(
  ...visibilities: ReadonlyArray<VisibilitySummary["visibility"]>
): VisibilitySummary[] {
  return visibilities.map((visibility) => ({ visibility }));
}

describe("visibilityBadge", () => {
  it("names the two states an operation skips by default", () => {
    expect(visibilityBadge("hidden")).toBe("Hidden");
    expect(visibilityBadge("very-hidden")).toBe("Very hidden");
  });

  it("leaves an ordinary worksheet unbadged", () => {
    expect(visibilityBadge("visible")).toBeUndefined();
  });
});

describe("hiddenWorksheetCallout", () => {
  it("says nothing when every described worksheet is visible", () => {
    expect(
      hiddenWorksheetCallout(sheets("visible", "visible")),
    ).toBeUndefined();
    expect(hiddenWorksheetCallout([])).toBeUndefined();
  });

  it("counts hidden worksheets, singular and plural", () => {
    expect(hiddenWorksheetCallout(sheets("visible", "hidden"))).toContain(
      "1 worksheet is hidden",
    );
    expect(hiddenWorksheetCallout(sheets("hidden", "hidden"))).toContain(
      "2 worksheets are hidden",
    );
  });

  it("counts a very hidden worksheet on its own", () => {
    expect(hiddenWorksheetCallout(sheets("very-hidden"))).toContain(
      "1 worksheet is very hidden",
    );
  });

  // The two states are counted apart because Excel reverses them differently,
  // and a reader hunting a missing column needs to know which kind they face.
  it("keeps the two hidden states apart when a workbook has both", () => {
    const callout = hiddenWorksheetCallout(
      sheets("visible", "hidden", "very-hidden", "very-hidden"),
    );
    expect(callout).toContain("1 worksheet is hidden and 2 are very hidden");
    expect(callout).toContain("VBA editor");
  });
});

describe("worksheetSummary", () => {
  it("reports the used range, the header row, and the rows below it", () => {
    expect(
      worksheetSummary({
        columnCount: 3,
        dataRowCount: 8,
        headerRow: 1,
        rowCount: 9,
      }),
    ).toBe("9 rows · 3 columns · header row 1 · 8 data rows");
  });

  it("counts one of anything in the singular", () => {
    expect(
      worksheetSummary({
        columnCount: 1,
        dataRowCount: 1,
        headerRow: 2,
        rowCount: 1,
      }),
    ).toBe("1 row · 1 column · header row 2 · 1 data row");
  });

  it("says so when the inspection found no header row", () => {
    expect(
      worksheetSummary({
        columnCount: 4,
        dataRowCount: 0,
        headerRow: undefined,
        rowCount: 6,
      }),
    ).toContain("no header row found");
  });

  // "0 rows · 0 columns · no header row found" reads as a report that failed
  // rather than as an empty tab, so an empty worksheet says what it is.
  it("describes an empty worksheet rather than printing zeros", () => {
    expect(
      worksheetSummary({
        columnCount: 0,
        dataRowCount: 0,
        headerRow: undefined,
        rowCount: 0,
      }),
    ).toBe("No cells with content");
  });
});

describe("columnPreviewNote", () => {
  it("says nothing while the whole worksheet is on the page", () => {
    expect(columnPreviewNote(12, 12)).toBeUndefined();
    // A shown count above the total cannot happen, but reporting "the first 12
    // of 8" if it ever did would be worse than saying nothing.
    expect(columnPreviewNote(8, 12)).toBeUndefined();
  });

  it("names both counts when the list was cut", () => {
    expect(columnPreviewNote(312, COLUMN_PREVIEW_LIMIT)).toBe(
      `Showing the first ${COLUMN_PREVIEW_LIMIT} of 312 columns`,
    );
  });

  // The limit exists to bound main-thread layout on a worksheet that can carry
  // thousands of columns, not to shorten the wide unions this toolkit is for:
  // a real fourteen-sheet consolidation produced sixty-nine columns.
  it("sits above the widths ordinary workbooks reach", () => {
    expect(COLUMN_PREVIEW_LIMIT).toBeGreaterThan(69);
  });
});

describe("sampleValueText", () => {
  it("prints a stored value without reformatting it", () => {
    expect(sampleValueText("North")).toBe('"North"');
    expect(sampleValueText(0)).toBe("0");
    expect(sampleValueText(45_292)).toBe("45292");
    expect(sampleValueText(true)).toBe("true");
  });

  // The description keeps these two apart on purpose, because a mapping review
  // has to see which is which. Printing both as 1 would discard that at the
  // last step, so the report has to show the difference.
  it("keeps the number 1 and the text 1 apart on the page", () => {
    expect(sampleValueText(1)).not.toBe(sampleValueText("1"));
    expect(sampleValueText("1")).toBe('"1"');
    expect(sampleValueText("true")).toBe('"true"');
  });

  it("escapes a quote inside the text rather than closing it early", () => {
    expect(sampleValueText('Region "North"')).toBe('"Region \\"North\\""');
  });

  // JSON has no literal for a non-finite number, so JSON.stringify would write
  // one as "null" and the report would show a stored value as absent.
  it("prints a non-finite number rather than reporting it as absent", () => {
    expect(sampleValueText(Number.POSITIVE_INFINITY)).toBe("Infinity");
  });

  it("prints nothing for a value the inspection never samples", () => {
    expect(sampleValueText(null)).toBe("");
  });
});
