/**
 * The header rule, tested as the property it is rather than by example, plus
 * the model-side profile it reads and the columns of a region that hold
 * anything.
 */
import { describe, expect, it } from "vitest";

import {
  countTitleRows,
  detectHeaderRow,
  HEADER_LOOKAHEAD_ROWS,
  isBlankValue,
  isNonTextValue,
  profileWorksheet,
  qualifiesAsHeader,
  regionColumns,
  type RowValueCount,
} from "../../src/region/header-detection.js";
import { FakeWorkbookModel } from "./fakes.js";

/** Every (values, fullest) pair up to a width no worksheet here exceeds. */
const WIDTHS = Array.from({ length: 24 }, (_, index) => index + 1);

/** Rows of all-text values, one entry per row, top to bottom. */
function textRows(...values: number[]): RowValueCount[] {
  return values.map((count, index) => ({
    nonTextValues: 0,
    row: index + 1,
    values: count,
  }));
}

describe("qualifiesAsHeader", () => {
  it("never qualifies a row holding nothing", () => {
    for (const fullest of WIDTHS) {
      expect(qualifiesAsHeader(0, fullest)).toBe(false);
    }
  });

  it("always qualifies the fullest row itself", () => {
    for (const fullest of WIDTHS) {
      expect(qualifiesAsHeader(fullest, fullest)).toBe(true);
    }
  });

  it("qualifies a row missing one name at every width", () => {
    // A header that leaves one column unnamed over data is still the header.
    for (const fullest of WIDTHS.filter((width) => width >= 2)) {
      expect(qualifiesAsHeader(fullest - 1, fullest)).toBe(true);
    }
  });

  it("rejects a row holding half or fewer of the values, from three columns up", () => {
    // A one-value title is rejected above a three-column table and a
    // two-value "Prepared by" line above a four-column one: three columns is
    // where a title can first be told from a header by count at all.
    for (const fullest of WIDTHS.filter((width) => width >= 3)) {
      for (let values = 1; values * 2 <= fullest; values += 1) {
        expect(qualifiesAsHeader(values, fullest)).toBe(false);
      }
    }
  });

  it("is monotonic in the values a row holds", () => {
    // Once a count qualifies, every larger count qualifies too: the rule is a
    // threshold, so adding a value to a row can never turn it into a title.
    for (const fullest of WIDTHS) {
      let qualified = false;
      for (let values = 0; values <= fullest; values += 1) {
        const now = qualifiesAsHeader(values, fullest);
        expect(now || !qualified).toBe(true);
        qualified = now;
      }
    }
  });

  it("reduces to the first value on a one- or two-column sheet", () => {
    expect(qualifiesAsHeader(1, 1)).toBe(true);
    expect(qualifiesAsHeader(1, 2)).toBe(true);
    expect(qualifiesAsHeader(2, 2)).toBe(true);
  });
});

describe("detectHeaderRow", () => {
  it("finds nothing on a worksheet where no row holds a value", () => {
    expect(detectHeaderRow([])).toBeUndefined();
    expect(detectHeaderRow(textRows(0, 0, 0))).toBeUndefined();
  });

  it("skips a title and a two-value line above a four-column header", () => {
    // Title, "Prepared by | name", blank, header, data.
    expect(detectHeaderRow(textRows(1, 2, 0, 4, 4, 4))).toBe(4);
  });

  it("keeps a header with one unnamed column over fuller data rows", () => {
    expect(detectHeaderRow(textRows(1, 3, 4, 4))).toBe(2);
  });

  it("picks the first qualifying row, never a later fuller one", () => {
    expect(detectHeaderRow(textRows(0, 5, 6, 6))).toBe(2);
  });

  it("returns the caller's own row labels", () => {
    expect(
      detectHeaderRow([
        { nonTextValues: 0, row: 7, values: 1 },
        { nonTextValues: 0, row: 9, values: 5 },
        { nonTextValues: 0, row: 12, values: 5 },
      ]),
    ).toBe(9);
  });

  it("answers the first row holding a value whenever no row holds more than one", () => {
    expect(detectHeaderRow(textRows(0, 0, 1, 1, 1))).toBe(3);
  });

  it("measures a row against the rows that follow it, not the whole sheet", () => {
    // A three-column table of twenty rows, then a blank row, then a
    // six-column block. The table's header is the header; the wider block
    // further down does not turn the table into title rows.
    const table = Array.from({ length: 20 }, () => 3);
    expect(detectHeaderRow(textRows(...table, 0, 6, 6, 6))).toBe(1);
  });

  it("looks a bounded number of populated rows ahead", () => {
    // Exactly HEADER_LOOKAHEAD_ROWS one-value lines above the header are still
    // skipped, because the header is within reach of the first line...
    const inReach = Array.from({ length: HEADER_LOOKAHEAD_ROWS }, () => 1);
    expect(detectHeaderRow(textRows(...inReach, 6, 6))).toBe(
      HEADER_LOOKAHEAD_ROWS + 1,
    );
    // ...and one more line is one too many: the first line then sees nothing
    // fuller than itself and is the header, as it always was.
    const outOfReach = Array.from(
      { length: HEADER_LOOKAHEAD_ROWS + 1 },
      () => 1,
    );
    expect(detectHeaderRow(textRows(...outOfReach, 6, 6))).toBe(1);
  });

  it("does not skip rows as titles when the header they leave holds numbers", () => {
    // A header that names two of four columns over rows holding a number: by
    // count the first data row would be the header, which would lose a data
    // row and name columns after a record. The guard keeps the first row.
    expect(
      detectHeaderRow([
        { nonTextValues: 0, row: 1, values: 2 },
        { nonTextValues: 1, row: 2, values: 4 },
        { nonTextValues: 1, row: 3, values: 4 },
      ]),
    ).toBe(1);
  });

  it("still skips single-value lines above a header of at least three values that holds numbers", () => {
    // A title and a subtitle above "Region | 2023 | 2024 | 2025".
    expect(
      detectHeaderRow([
        { nonTextValues: 0, row: 1, values: 1 },
        { nonTextValues: 1, row: 2, values: 1 },
        { nonTextValues: 3, row: 4, values: 4 },
        { nonTextValues: 3, row: 5, values: 4 },
      ]),
    ).toBe(4);
    // But not a two-value line: that could as well be a sparse header.
    expect(
      detectHeaderRow([
        { nonTextValues: 0, row: 1, values: 2 },
        { nonTextValues: 3, row: 2, values: 4 },
        { nonTextValues: 3, row: 3, values: 4 },
      ]),
    ).toBe(1);
    // Nor above a two-value header holding a number.
    expect(
      detectHeaderRow([
        { nonTextValues: 0, row: 1, values: 1 },
        { nonTextValues: 1, row: 3, values: 2 },
        { nonTextValues: 1, row: 4, values: 2 },
      ]),
    ).toBe(1);
  });

  it("holds the property for every all-text count sequence in a small space", () => {
    // Every sequence of up to four rows holding zero to five text values
    // each: the answer is the first row that qualifies against the fullest of
    // itself and the rows after it, and no row above it does.
    const range = [0, 1, 2, 3, 4, 5];
    for (const a of range) {
      for (const b of range) {
        for (const c of range) {
          for (const d of range) {
            const sequence = textRows(a, b, c, d);
            const populated = sequence.filter((count) => count.values > 0);
            const header = detectHeaderRow(sequence);
            if (populated.length === 0) {
              expect(header).toBeUndefined();
              continue;
            }
            expect(header).toBeDefined();
            const fullestFrom = (index: number): number =>
              Math.max(...populated.slice(index).map((count) => count.values));
            const chosenIndex = populated.findIndex(
              (count) => count.row === header,
            );
            expect(chosenIndex).toBeGreaterThanOrEqual(0);
            expect(
              qualifiesAsHeader(
                populated[chosenIndex]!.values,
                fullestFrom(chosenIndex),
              ),
            ).toBe(true);
            for (let above = 0; above < chosenIndex; above += 1) {
              expect(
                qualifiesAsHeader(populated[above]!.values, fullestFrom(above)),
              ).toBe(false);
            }
          }
        }
      }
    }
  });

  it("never skips rows that themselves form a table", () => {
    // A three-column table of a header and two all-text rows, then an
    // eight-column all-text block within reach: by count no row of the small
    // table qualifies and the wide header would pass the text test, so the
    // whole small table would be read as title rows. Two consecutive rows of
    // three or more values are a table, and a table is never a title.
    expect(detectHeaderRow(textRows(3, 3, 3, 8, 8, 8))).toBe(1);
    expect(detectHeaderRow(textRows(3, 3, 0, 8, 8, 8))).toBe(1);
    // A three-value line set apart from the wide block by a blank row is a
    // title line and is skipped; the same line directly above the block may
    // be a three-name header over its first record, and reads as one.
    expect(detectHeaderRow(textRows(3, 0, 8, 8, 8))).toBe(3);
    expect(detectHeaderRow(textRows(3, 8, 8, 8))).toBe(1);
    // A two-column key-value block above a wide table reads as a title
    // block: two values per line is what "Prepared by | name" holds too.
    expect(detectHeaderRow(textRows(1, 2, 2, 8, 8, 8))).toBe(4);
    // The property: whenever the first two populated rows are adjacent and
    // each holds at least three values, the first is the header, whatever
    // follows.
    for (const first of [3, 4, 5]) {
      for (const second of [3, 4, 5]) {
        for (const later of [6, 8, 12]) {
          expect(
            detectHeaderRow(textRows(0, first, second, later, later, later)),
          ).toBe(2);
        }
      }
    }
  });

  it("never answers below the first populated row when the guard refuses", () => {
    // Whatever the counts, a header holding numbers with a multi-value row
    // above it means the first populated row is the answer.
    for (const above of [2, 3, 4]) {
      for (const width of [4, 5, 6, 8]) {
        expect(
          detectHeaderRow([
            { nonTextValues: 0, row: 2, values: above },
            { nonTextValues: 1, row: 5, values: width },
            { nonTextValues: 2, row: 6, values: width },
          ]),
        ).toBe(2);
      }
    }
  });
});

describe("countTitleRows", () => {
  it("counts only the rows above the header that hold a value", () => {
    const sequence = textRows(1, 0, 2, 6, 6);
    expect(countTitleRows(sequence, 4)).toBe(2);
    expect(countTitleRows(sequence, 1)).toBe(0);
    expect(countTitleRows(sequence, 5)).toBe(3);
  });
});

describe("isBlankValue and isNonTextValue", () => {
  it("treats nothing, null, and the empty string as blank and everything else as a value", () => {
    expect(isBlankValue(undefined)).toBe(true);
    expect(isBlankValue(null)).toBe(true);
    expect(isBlankValue("")).toBe(true);
    expect(isBlankValue(" ")).toBe(false);
    expect(isBlankValue(0)).toBe(false);
    expect(isBlankValue(false)).toBe(false);
    expect(isBlankValue("#REF!")).toBe(false);
  });

  it("classifies numbers, booleans, and error cells as not text", () => {
    expect(isNonTextValue(2024, undefined)).toBe(true);
    expect(isNonTextValue(true, "b")).toBe(true);
    // The two readers hand an error back differently; the cell type decides.
    expect(isNonTextValue("#REF!", "e")).toBe(true);
    expect(isNonTextValue(23, "e")).toBe(true);
    expect(isNonTextValue("Region", "s")).toBe(false);
    expect(isNonTextValue("2024-03-09", undefined)).toBe(false);
    expect(isNonTextValue(new Date(0), undefined)).toBe(false);
  });
});

describe("profileWorksheet", () => {
  it("counts values inside the used range from the rows the sheet stores", () => {
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Quarterly review log", null, null],
            // A row the sheet does not store at all: it never reaches the
            // counts, and the rule does not need it to.
            [null, null, null],
            ["Case_ID", "Region", "Failed Checks"],
            ["R-1", "north", 5],
            // A present but blank cell holds nothing.
            ["R-2", "", 7],
          ],
          name: "Review Log",
        },
      ],
    });
    const worksheet = workbook.worksheet("Review Log")!;

    const profile = profileWorksheet(
      worksheet,
      worksheet.rows(),
      worksheet.usedRange!,
    );
    expect(profile.counts).toEqual([
      { nonTextValues: 0, row: 1, values: 1 },
      { nonTextValues: 0, row: 3, values: 3 },
      { nonTextValues: 1, row: 4, values: 3 },
      { nonTextValues: 1, row: 5, values: 2 },
    ]);
    expect([...profile.lastValueRow.entries()]).toEqual([
      [0, 5],
      [1, 4],
      [2, 5],
    ]);
  });

  it("counts an uncalculated formula as nothing and an error as a non-text value", () => {
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            [
              { formula: { kind: "normal", text: 'CONCATENATE("Q","1")' } },
              { text: "#REF!", type: "e" },
            ],
            ["Case_ID", "Region"],
          ],
          name: "Data",
        },
      ],
    });
    const worksheet = workbook.worksheet("Data")!;

    expect(
      profileWorksheet(worksheet, worksheet.rows(), worksheet.usedRange!)
        .counts,
    ).toEqual([
      { nonTextValues: 1, row: 1, values: 1 },
      { nonTextValues: 0, row: 2, values: 2 },
    ]);
  });
});

describe("regionColumns", () => {
  function columnsOf(
    grid: ConstructorParameters<
      typeof FakeWorkbookModel
    >[0]["sheets"][number]["grid"],
    headerRow: number,
  ) {
    const workbook = new FakeWorkbookModel({
      sheets: [{ grid, name: "Data" }],
    });
    const worksheet = workbook.worksheet("Data")!;
    const used = worksheet.usedRange!;
    const profile = profileWorksheet(worksheet, worksheet.rows(), used);
    return regionColumns(
      worksheet,
      headerRow,
      {
        end: used.end,
        start: { column: used.start.column, row: headerRow + 1 },
      },
      profile.lastValueRow,
    );
  }

  it("leaves out a column that holds nothing in the header row or under it", () => {
    expect(
      columnsOf(
        [
          ["Case_ID", null, "Region", null],
          ["R-1", null, "north", "side note"],
          ["R-2", null, "south", null],
        ],
        1,
      ),
    ).toEqual([
      { index: 0, name: "Case_ID" },
      { index: 2, name: "Region" },
      // A blank header over values is a column with no name, not a spacer.
      { index: 3, name: "" },
    ]);
  });

  it("keeps a named column with nothing under it", () => {
    expect(
      columnsOf(
        [
          ["Case_ID", "Notes"],
          ["R-1", null],
        ],
        1,
      ),
    ).toEqual([
      { index: 0, name: "Case_ID" },
      { index: 1, name: "Notes" },
    ]);
  });

  it("ignores values above the header row", () => {
    // The title cell sits above the header in the middle column; it does not
    // make that column part of the region.
    expect(
      columnsOf(
        [
          [null, "title cell", null],
          ["Case_ID", null, "Region"],
          ["R-1", null, "north"],
        ],
        2,
      ),
    ).toEqual([
      { index: 0, name: "Case_ID" },
      { index: 2, name: "Region" },
    ]);
  });
});
