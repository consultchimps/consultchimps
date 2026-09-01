import { describe, expect, it } from "vitest";

import {
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

describe("sampleValueText", () => {
  it("prints a stored value without reformatting it", () => {
    expect(sampleValueText("North")).toBe("North");
    expect(sampleValueText(0)).toBe("0");
    expect(sampleValueText(45_292)).toBe("45292");
    expect(sampleValueText(true)).toBe("true");
  });

  it("prints nothing for a value the inspection never samples", () => {
    expect(sampleValueText(null)).toBe("");
  });
});
