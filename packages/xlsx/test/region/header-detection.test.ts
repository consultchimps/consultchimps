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
  isTitleLine,
  profileWorksheet,
  qualifiesAsHeader,
  regionColumns,
  type RowValueCount,
} from "../../src/region/header-detection.js";
import { FakeWorkbookModel } from "./fakes.js";

/** Every (values, fullest) pair up to a width no worksheet here exceeds. */
const WIDTHS = Array.from({ length: 24 }, (_, index) => index + 1);

/** Rows numbered from 1, one entry per row, top to bottom; 0 is a blank row. */
function rows(...values: number[]): RowValueCount[] {
  return values.map((count, index) => ({ row: index + 1, values: count }));
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

describe("isTitleLine", () => {
  const header = (values: number): RowValueCount => ({ row: 9, values });
  const line = (values: number): RowValueCount => ({ row: 1, values });

  it("accepts any populated line a blank row sets apart from the header", () => {
    for (const width of WIDTHS) {
      expect(isTitleLine(line(width), header(width), true)).toBe(true);
    }
  });

  it("otherwise accepts only a line naming fewer than a third of the header", () => {
    for (const width of WIDTHS) {
      for (let values = 1; values <= width; values += 1) {
        expect(isTitleLine(line(values), header(width), false)).toBe(
          values * 3 < width,
        );
      }
    }
    // The shapes the rule is judged by: a one-value title above four columns
    // is a title; above three it may be a header naming one column.
    expect(isTitleLine(line(1), header(4), false)).toBe(true);
    expect(isTitleLine(line(1), header(3), false)).toBe(false);
    expect(isTitleLine(line(2), header(7), false)).toBe(true);
    expect(isTitleLine(line(2), header(6), false)).toBe(false);
  });
});

describe("detectHeaderRow", () => {
  it("finds nothing on a worksheet where no row holds a value", () => {
    expect(detectHeaderRow([])).toBeUndefined();
    expect(detectHeaderRow(rows(0, 0, 0))).toBeUndefined();
  });

  it("skips a title block a blank row sets apart from the header", () => {
    // Title, "Prepared by | name", blank, header, data.
    expect(detectHeaderRow(rows(1, 2, 0, 4, 4, 4))).toBe(4);
    // The same block directly above a wide header: the one-value title names
    // fewer than a third of it, and so does the two-value line.
    expect(detectHeaderRow(rows(1, 2, 7, 7, 7))).toBe(3);
  });

  it("keeps a sparse header above all-text records rather than losing one", () => {
    // `Name | |` above `Alice | North | Open`: by count a title above a
    // header, by count also a header naming one of three columns over its
    // first record. Nothing tells them apart, so the first row is the
    // header and no record is lost.
    expect(detectHeaderRow(rows(1, 3, 3, 3))).toBe(1);
    // Two names of four above records: the same.
    expect(detectHeaderRow(rows(2, 4, 4, 4))).toBe(1);
    // A blank row between them is what makes the first a title.
    expect(detectHeaderRow(rows(1, 0, 3, 3, 3))).toBe(3);
    // Four columns is where a single value stops being a plausible header.
    expect(detectHeaderRow(rows(1, 4, 4, 4))).toBe(2);
  });

  it("keeps a header with one unnamed column over fuller data rows", () => {
    expect(detectHeaderRow(rows(1, 0, 3, 4, 4))).toBe(3);
  });

  it("picks the first qualifying row, never a later fuller one", () => {
    expect(detectHeaderRow(rows(0, 5, 6, 6))).toBe(2);
  });

  it("returns the caller's own row labels", () => {
    expect(
      detectHeaderRow([
        { row: 7, values: 1 },
        { row: 9, values: 5 },
        { row: 12, values: 5 },
      ]),
    ).toBe(9);
  });

  it("answers the first row holding a value whenever no row holds more than one", () => {
    expect(detectHeaderRow(rows(0, 0, 1, 1, 1))).toBe(3);
  });

  it("measures a row against the rows that follow it, not the whole sheet", () => {
    // A three-column table of twenty rows, then a blank row, then a
    // six-column block. The table's header is the header; the wider block
    // further down does not turn the table into title rows.
    const table = Array.from({ length: 20 }, () => 3);
    expect(detectHeaderRow(rows(...table, 0, 6, 6, 6))).toBe(1);
  });

  it("looks a bounded number of populated rows ahead", () => {
    // Exactly HEADER_LOOKAHEAD_ROWS one-value lines above a wide header are
    // still skipped, because the header is within reach of the first line...
    const inReach = Array.from({ length: HEADER_LOOKAHEAD_ROWS }, () => 1);
    expect(detectHeaderRow(rows(...inReach, 6, 6))).toBe(
      HEADER_LOOKAHEAD_ROWS + 1,
    );
    // ...and one more line is one too many: the first line then sees nothing
    // fuller than itself and is the header, as it always was.
    const outOfReach = Array.from(
      { length: HEADER_LOOKAHEAD_ROWS + 1 },
      () => 1,
    );
    expect(detectHeaderRow(rows(...outOfReach, 6, 6))).toBe(1);
  });

  it("skips a title above year columns without needing to know they are numbers", () => {
    // A title and a subtitle above `Region | 2023 | 2024 | 2025`.
    expect(detectHeaderRow(rows(1, 1, 0, 4, 4))).toBe(4);
    expect(detectHeaderRow(rows(1, 1, 4, 4))).toBe(3);
  });

  it("never skips rows that themselves form a table", () => {
    // A three-column table of a header and two rows, then an eight-column
    // block within reach: each small-table row names more than a third of
    // the wide header, and two adjacent rows of three or more are a table.
    expect(detectHeaderRow(rows(3, 3, 3, 8, 8, 8))).toBe(1);
    expect(detectHeaderRow(rows(3, 3, 0, 8, 8, 8))).toBe(1);
    // Even a three-column table above a block wide enough that a third of it
    // is more than three: adjacent rows of three or more are a table.
    expect(detectHeaderRow(rows(3, 3, 0, 12, 12, 12))).toBe(1);
    // A three-value line set apart from the wide block by a blank row is a
    // title line and is skipped; the same line directly above the block may
    // be a three-name header over its first record, and reads as one.
    expect(detectHeaderRow(rows(3, 0, 12, 12, 12))).toBe(3);
    expect(detectHeaderRow(rows(3, 12, 12, 12))).toBe(1);
    // The property: whenever the first two populated rows are adjacent and
    // each holds at least three values, the first is the header, whatever
    // follows.
    for (const first of [3, 4, 5]) {
      for (const second of [3, 4, 5]) {
        for (const later of [6, 8, 12, 20]) {
          expect(
            detectHeaderRow(rows(0, first, second, later, later, later)),
          ).toBe(2);
        }
      }
    }
  });

  it("always answers a populated row, and the first one without evidence", () => {
    // Every sequence of up to five rows holding zero to six values each: the
    // answer is always a row holding a value, and whenever no blank row
    // separates the rows and every row names at least a third of every row
    // below it, the answer is the first populated row.
    const range = [0, 1, 2, 3, 4, 5, 6];
    for (const a of range) {
      for (const b of range) {
        for (const c of range) {
          for (const d of range) {
            for (const e of range) {
              const sequence = rows(a, b, c, d, e);
              const populated = sequence.filter((count) => count.values > 0);
              const header = detectHeaderRow(sequence);
              if (populated.length === 0) {
                expect(header).toBeUndefined();
                continue;
              }
              expect(populated.some((count) => count.row === header)).toBe(
                true,
              );
              const contiguous = populated.every(
                (count, index) =>
                  index === 0 || count.row === populated[index - 1]!.row + 1,
              );
              const noSparseLine = populated.every((count, index) =>
                populated
                  .slice(index + 1)
                  .every((below) => count.values * 3 >= below.values),
              );
              if (contiguous && noSparseLine) {
                expect(header).toBe(populated[0]!.row);
              }
            }
          }
        }
      }
    }
  });
});

describe("countTitleRows", () => {
  it("counts only the rows above the header that hold a value", () => {
    const sequence = rows(1, 0, 2, 6, 6);
    expect(countTitleRows(sequence, 4)).toBe(2);
    expect(countTitleRows(sequence, 1)).toBe(0);
    expect(countTitleRows(sequence, 5)).toBe(3);
  });
});

describe("isBlankValue", () => {
  it("treats nothing, null, and the empty string as blank and everything else as a value", () => {
    expect(isBlankValue(undefined)).toBe(true);
    expect(isBlankValue(null)).toBe(true);
    expect(isBlankValue("")).toBe(true);
    expect(isBlankValue(" ")).toBe(false);
    expect(isBlankValue(0)).toBe(false);
    expect(isBlankValue(false)).toBe(false);
    expect(isBlankValue("#REF!")).toBe(false);
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
            // counts, and the rule reads the gap in row numbers as the blank
            // row it is.
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
      { row: 1, values: 1 },
      { row: 3, values: 3 },
      { row: 4, values: 3 },
      { row: 5, values: 2 },
    ]);
    expect([...profile.lastValueRow.entries()]).toEqual([
      [0, 5],
      [1, 4],
      [2, 5],
    ]);
  });

  it("counts an uncalculated formula as nothing and an error as a value", () => {
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
      { row: 1, values: 1 },
      { row: 2, values: 2 },
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
