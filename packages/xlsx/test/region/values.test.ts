import { describe, expect, it } from "vitest";

import {
  columnLetters,
  formatCellRef,
  normalizeHeader,
  normalizeSplitValue,
  parseCellRange,
  parseCellRef,
  parseColumnLetters,
  parseSheetRange,
} from "../../src/region/values.js";

// Non-ASCII inputs are written as escapes so the file stays pure ASCII.
const FULLWIDTH_REGION = "\uFF32\uFF45\uFF47\uFF49\uFF4F\uFF4E"; // fullwidth Region
const NON_BREAKING_SPACE = "\u00A0";

describe("normalizeHeader", () => {
  it("applies NFKC, trims, and folds case", () => {
    expect(normalizeHeader("  Region  ")).toBe("region");
    expect(normalizeHeader("REGION")).toBe("region");
    expect(normalizeHeader(FULLWIDTH_REGION)).toBe("region");
  });

  it("folds a non-breaking space to a normal space before trimming", () => {
    expect(normalizeHeader(`${NON_BREAKING_SPACE}Region`)).toBe("region");
  });
});

describe("normalizeSplitValue", () => {
  it("treats null, undefined, and blank text as blank", () => {
    expect(normalizeSplitValue(null, false)).toBeUndefined();
    expect(normalizeSplitValue(undefined, false)).toBeUndefined();
    expect(normalizeSplitValue("   ", false)).toBeUndefined();
    expect(normalizeSplitValue("   ", true)).toBeUndefined();
  });

  it("keys dates by ISO string", () => {
    const value = new Date(Date.UTC(2026, 0, 2));
    expect(normalizeSplitValue(value, false)).toEqual({
      display: "2026-01-02T00:00:00.000Z",
      key: "date:2026-01-02T00:00:00.000Z",
    });
  });

  it("keys numbers, normalizing negative zero and rejecting non-finite", () => {
    expect(normalizeSplitValue(100, false)?.key).toBe("number:100");
    expect(normalizeSplitValue(-0, false)).toEqual({
      display: "0",
      key: "number:0",
    });
    expect(normalizeSplitValue(Number.NaN, false)).toBeUndefined();
    expect(normalizeSplitValue(Number.POSITIVE_INFINITY, true)).toBeUndefined();
  });

  it("keys booleans", () => {
    expect(normalizeSplitValue(true, false)).toEqual({
      display: "true",
      key: "boolean:true",
    });
  });

  it("keeps untrimmed and differently cased text apart under strict matching", () => {
    expect(normalizeSplitValue(" North ", true)).toEqual({
      display: "North",
      key: "string: North ",
    });
    expect(normalizeSplitValue("north", true)?.key).toBe("string:north");
    expect(normalizeSplitValue("North", true)?.key).toBe("string:North");
  });

  it("folds trimming, case, and NFKC together under tolerant matching", () => {
    const keys = [" North ", "north", "NORTH", "North"].map(
      (value) => normalizeSplitValue(value, false)?.key,
    );
    expect(new Set(keys)).toEqual(new Set(["string:north"]));
  });

  it("coerces numeric text to a number key only under tolerant matching", () => {
    expect(normalizeSplitValue("100", false)?.key).toBe("number:100");
    expect(normalizeSplitValue("1e3", false)?.key).toBe("number:1000");
    expect(normalizeSplitValue("-0", false)?.key).toBe("number:0");
    expect(normalizeSplitValue("100", true)?.key).toBe("string:100");
  });

  it("keeps the untrimmed display out of the numeric coercion decision", () => {
    expect(normalizeSplitValue(" 100 ", false)).toEqual({
      display: "100",
      key: "number:100",
    });
  });

  it("leaves text that only looks numeric alone", () => {
    expect(normalizeSplitValue("100a", false)?.key).toBe("string:100a");
    expect(normalizeSplitValue("1,000", false)?.key).toBe("string:1,000");
  });
});

describe("A1 reference helpers", () => {
  it("round-trips column letters", () => {
    expect(columnLetters(0)).toBe("A");
    expect(columnLetters(25)).toBe("Z");
    expect(columnLetters(26)).toBe("AA");
    expect(columnLetters(701)).toBe("ZZ");
    expect(parseColumnLetters("A")).toBe(0);
    expect(parseColumnLetters("aa")).toBe(26);
    expect(parseColumnLetters("ZZ")).toBe(701);
  });

  it("formats and parses cell references", () => {
    expect(formatCellRef({ column: 2, row: 7 })).toBe("C7");
    expect(parseCellRef("$C$7")).toEqual({ column: 2, row: 7 });
    expect(parseCellRef("C0")).toBeUndefined();
    expect(parseCellRef("not a ref")).toBeUndefined();
  });

  it("parses ranges and normalizes the corners", () => {
    expect(parseCellRange("A1:C5")).toEqual({
      end: { column: 2, row: 5 },
      start: { column: 0, row: 1 },
    });
    expect(parseCellRange("C5:A1")).toEqual({
      end: { column: 2, row: 5 },
      start: { column: 0, row: 1 },
    });
    expect(parseCellRange("B2")).toEqual({
      end: { column: 1, row: 2 },
      start: { column: 1, row: 2 },
    });
    expect(parseCellRange("A1:B2:C3")).toBeUndefined();
  });

  it("parses sheet-qualified ranges, including quoted sheet names", () => {
    expect(parseSheetRange("Data!$A$1:$C$5")).toEqual({
      range: { end: { column: 2, row: 5 }, start: { column: 0, row: 1 } },
      sheet: "Data",
    });
    expect(parseSheetRange("'My Sheet'!A1:B2")?.sheet).toBe("My Sheet");
    expect(parseSheetRange("'It''s Data'!A1:B2")?.sheet).toBe("It's Data");
    expect(parseSheetRange("A1:B2")).toBeUndefined();
  });
});
