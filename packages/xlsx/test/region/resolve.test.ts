import { describe, expect, it } from "vitest";

import { RangeBinding } from "../../src/region/range-binding.js";
import { resolveRegions } from "../../src/region/resolve.js";
import { TableBinding } from "../../src/region/table-binding.js";
import { FakeWorkbookModel, fakeTable } from "./fakes.js";

const FULLWIDTH_REGION = "\uFF32\uFF45\uFF47\uFF49\uFF4F\uFF4E"; // fullwidth Region

function salesWorkbook(): FakeWorkbookModel {
  return new FakeWorkbookModel({
    sheets: [
      {
        grid: [
          ["Region", "Amount"],
          ["North", 10],
          ["South", 20],
        ],
        name: "Data",
      },
      {
        grid: [["Total"], [30]],
        name: "Summary",
      },
      {
        grid: [
          ["Region", "Amount"],
          ["East", 5],
        ],
        name: "More Data",
      },
    ],
  });
}

describe("resolveRegions: { find }", () => {
  it("returns one region from the first sheet, in workbook order, that matches", async () => {
    const workbook = salesWorkbook();
    const [region, ...rest] = await resolveRegions(
      workbook,
      { find: "Region" },
      "Region",
    );

    expect(rest).toHaveLength(0);
    expect(region).toBeInstanceOf(RangeBinding);
    expect(region?.sheetName).toBe("Data");
    expect(region?.origin).toEqual({ kind: "detected-header" });
    expect(region?.headerRow).toBe(1);
    expect(region?.body).toEqual({
      end: { column: 1, row: 3 },
      start: { column: 0, row: 2 },
    });
    expect(region?.columns).toEqual([
      { index: 0, name: "Region" },
      { index: 1, name: "Amount" },
    ]);
  });

  it("matches header text after NFKC, trimming, and case folding", async () => {
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["  REGION  ", "Amount"],
            ["North", 10],
          ],
          name: "Data",
        },
      ],
    });

    const [region] = await resolveRegions(
      workbook,
      { find: FULLWIDTH_REGION },
      FULLWIDTH_REGION,
    );

    expect(region?.headerRow).toBe(1);
  });

  it("prefers the topmost matching row, then the leftmost matching column", async () => {
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Notes", "Ignore"],
            ["Region", "Amount", "Region"],
            ["North", 10, "North"],
            ["Region", 0, "Region"],
          ],
          name: "Data",
        },
      ],
    });

    const [region] = await resolveRegions(
      workbook,
      { find: "Region" },
      "Region",
    );

    expect(region?.headerRow).toBe(2);
    expect(region?.body.start.row).toBe(3);
  });

  it("uses find as the anchor and column as the key column when they differ", async () => {
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [["Amount"], ["Region", "Amount"], ["North", 10]],
          name: "Data",
        },
      ],
    });

    // "Amount" alone would anchor on the stray label in row 1; anchoring on
    // "Region" pins the real header row and the key column is found there.
    const [region] = await resolveRegions(
      workbook,
      { find: "Region" },
      "Amount",
    );

    expect(region?.headerRow).toBe(2);
    expect(region?.body.start.row).toBe(3);
  });

  it("skips a sheet whose anchor row does not also carry the key column", async () => {
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Region", "Notes"],
            ["North", "n/a"],
          ],
          name: "Partial",
        },
        {
          grid: [
            ["Region", "Amount"],
            ["South", 20],
          ],
          name: "Complete",
        },
      ],
    });

    const [region] = await resolveRegions(
      workbook,
      { find: "Region" },
      "Amount",
    );

    expect(region?.sheetName).toBe("Complete");
  });

  it("refuses when no sheet carries the header", async () => {
    await expect(
      resolveRegions(salesWorkbook(), { find: "Territory" }, "Territory"),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_COLUMN_NOT_FOUND" });
  });

  it("refuses when no column is supplied", async () => {
    await expect(
      resolveRegions(salesWorkbook(), { find: "Region" }),
    ).rejects.toMatchObject({ code: "XLSX_NO_COLUMNS" });
  });
});

describe('resolveRegions: "all-worksheets"', () => {
  it("returns every matching sheet in workbook order and skips the rest", async () => {
    const regions = await resolveRegions(
      salesWorkbook(),
      "all-worksheets",
      "Region",
    );

    expect(regions.map((region) => region.sheetName)).toEqual([
      "Data",
      "More Data",
    ]);
  });

  it("refuses when nothing matches", async () => {
    await expect(
      resolveRegions(salesWorkbook(), "all-worksheets", "Territory"),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_COLUMN_NOT_FOUND" });
  });

  it("refuses when no column is supplied", async () => {
    await expect(
      resolveRegions(salesWorkbook(), "all-worksheets"),
    ).rejects.toMatchObject({ code: "XLSX_NO_COLUMNS" });
  });
});

describe("resolveRegions: table association", () => {
  function tableWorkbook(headerRow: number): FakeWorkbookModel {
    return new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Region", "Amount"],
            ["North", 10],
            ["South", 20],
            ["Total", 30],
          ],
          name: "Data",
        },
      ],
      tables: [
        fakeTable({
          columnNames: ["Region", "Amount"],
          headerRow,
          name: "Sales",
          ref: `A${String(headerRow)}:B4`,
          sheetName: "Data",
          totalsRow: true,
        }),
      ],
    });
  }

  it("binds a detected header that sits on a table's header row", async () => {
    const [region] = await resolveRegions(
      tableWorkbook(1),
      { find: "Region" },
      "Region",
    );

    expect(region).toBeInstanceOf(TableBinding);
    expect(region?.origin).toEqual({ kind: "table", tableName: "Sales" });
    // The totals row in row 4 is excluded from the body.
    expect(region?.body).toEqual({
      end: { column: 1, row: 3 },
      start: { column: 0, row: 2 },
    });
  });

  it("does not associate when the table starts on another row", async () => {
    const [region] = await resolveRegions(
      tableWorkbook(2),
      { find: "Region" },
      "Region",
    );

    expect(region).toBeInstanceOf(RangeBinding);
  });

  it("does not associate when the table names a different column at that offset", async () => {
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Region", "Amount"],
            ["North", 10],
          ],
          name: "Data",
        },
      ],
      tables: [
        fakeTable({
          columnNames: ["Territory", "Amount"],
          name: "Sales",
          ref: "A1:B2",
          sheetName: "Data",
        }),
      ],
    });

    const [region] = await resolveRegions(
      workbook,
      { find: "Region" },
      "Region",
    );

    expect(region).toBeInstanceOf(RangeBinding);
  });

  it("does not associate a table that has no header row", async () => {
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Region", "Amount"],
            ["North", 10],
          ],
          name: "Data",
        },
      ],
      tables: [
        fakeTable({
          columnNames: ["Region", "Amount"],
          headerRow: 0,
          name: "Sales",
          ref: "A1:B2",
          sheetName: "Data",
        }),
      ],
    });

    const [region] = await resolveRegions(
      workbook,
      { find: "Region" },
      "Region",
    );

    expect(region).toBeInstanceOf(RangeBinding);
  });
});

describe("resolveRegions: { sheet }", () => {
  function titledWorkbook(): FakeWorkbookModel {
    return new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Region"],
            [null],
            ["Region", "Amount"],
            ["North", 10],
            ["South", 20],
          ],
          name: "Data",
        },
      ],
    });
  }

  it("honors a headerRow override instead of detecting a stray match", async () => {
    const [region] = await resolveRegions(
      titledWorkbook(),
      { headerRow: 3, sheet: "Data" },
      "Region",
    );

    expect(region?.origin).toEqual({ kind: "declared-header" });
    expect(region?.headerRow).toBe(3);
    expect(region?.body).toEqual({
      end: { column: 1, row: 5 },
      start: { column: 0, row: 4 },
    });
  });

  it("detects the header when no override is given", async () => {
    const [region] = await resolveRegions(
      titledWorkbook(),
      { sheet: "Data" },
      "Region",
    );

    expect(region?.origin).toEqual({ kind: "detected-header" });
    expect(region?.headerRow).toBe(1);
  });

  it("refuses when the override row does not carry the column", async () => {
    await expect(
      resolveRegions(
        titledWorkbook(),
        { headerRow: 4, sheet: "Data" },
        "Region",
      ),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_COLUMN_NOT_FOUND" });
  });

  it("refuses a header row below 1", async () => {
    await expect(
      resolveRegions(
        titledWorkbook(),
        { headerRow: 0, sheet: "Data" },
        "Region",
      ),
    ).rejects.toMatchObject({ code: "XLSX_INVALID_HEADER_ROW" });
  });

  it("refuses an unknown worksheet", async () => {
    await expect(
      resolveRegions(titledWorkbook(), { sheet: "Missing" }, "Region"),
    ).rejects.toMatchObject({ code: "XLSX_WORKSHEET_NOT_FOUND" });
  });

  it("uses the declared row without a column to search for", async () => {
    const [region] = await resolveRegions(titledWorkbook(), {
      headerRow: 3,
      sheet: "Data",
    });

    expect(region?.headerRow).toBe(3);
    expect(region?.columns).toEqual([
      { index: 0, name: "Region" },
      { index: 1, name: "Amount" },
    ]);
  });

  it("detects the header from the first used row without a column or an override", async () => {
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          firstRow: 2,
          grid: [
            ["Region", "Amount"],
            ["North", 10],
          ],
          name: "Data",
        },
      ],
    });

    const [region] = await resolveRegions(workbook, { sheet: "Data" });

    expect(region?.headerRow).toBe(2);
    expect(region?.body).toEqual({
      end: { column: 1, row: 3 },
      start: { column: 0, row: 3 },
    });
  });

  it("skips a title block above the header without a column or an override", async () => {
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Quarterly review log", null, null, null],
            [null, null, null, null],
            ["Prepared by", "Reviewer 1", null, null],
            [null, null, null, null],
            ["Case_ID", "Region", "Failed Checks", "Owner"],
            ["R-1", "north", 5, "Reviewer 2"],
          ],
          name: "Data",
        },
      ],
    });

    const [region] = await resolveRegions(workbook, { sheet: "Data" });

    expect(region?.origin).toEqual({ kind: "detected-header" });
    expect(region?.headerRow).toBe(5);
    expect(region?.body).toEqual({
      end: { column: 3, row: 6 },
      start: { column: 0, row: 6 },
    });
    expect(region?.columns.map((column) => column.name)).toEqual([
      "Case_ID",
      "Region",
      "Failed Checks",
      "Owner",
    ]);
  });

  it("keeps the first used row and every used column when no row holds a value", async () => {
    // A sheet of uncalculated formulas is populated, so an inspection still
    // describes it from its first row rather than reporting no header at all,
    // and with its columns: no value anywhere means no spacer can be told from
    // a column, so none is left out.
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          firstRow: 2,
          grid: [
            [
              { formula: { kind: "normal", text: "1+1" } },
              { formula: { kind: "normal", text: "2+2" } },
            ],
            [
              { formula: { kind: "normal", text: "3+3" } },
              { formula: { kind: "normal", text: "4+4" } },
            ],
          ],
          name: "Formulas",
        },
      ],
    });

    const [region] = await resolveRegions(workbook, { sheet: "Formulas" });

    expect(region?.headerRow).toBe(2);
    expect(region?.columns).toEqual([
      { index: 0, name: "" },
      { index: 1, name: "" },
    ]);
  });

  it("keeps a sparse header over rows holding numbers rather than skipping it", async () => {
    // Two of four columns named, over data with a number in it. By count the
    // first data row would be the header; the guard keeps row 1, so no data
    // row is lost.
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Case_ID", "Region", null, null],
            ["R-1", "north", "late", 5],
            ["R-2", "south", "early", 7],
          ],
          name: "Data",
        },
      ],
    });

    const [region] = await resolveRegions(workbook, { sheet: "Data" });

    expect(region?.headerRow).toBe(1);
    expect(region?.columns.map((column) => column.name)).toEqual([
      "Case_ID",
      "Region",
      "",
      "",
    ]);
  });

  it("leaves spacer columns out of the region's columns", async () => {
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Region", null, "Amount", null, null],
            ["North", null, 10, null, "side note"],
            ["South", null, 20, null, null],
          ],
          name: "Data",
        },
      ],
    });

    const [detected] = await resolveRegions(workbook, { sheet: "Data" });
    // The body still spans the used range - a spacer was examined and found
    // empty - but the columns are only the ones holding anything, and a blank
    // header over values stays as an unnamed column.
    expect(detected?.body).toEqual({
      end: { column: 4, row: 3 },
      start: { column: 0, row: 2 },
    });
    expect(detected?.columns).toEqual([
      { index: 0, name: "Region" },
      { index: 2, name: "Amount" },
      { index: 4, name: "" },
    ]);

    // A region found by its header text answers the same, so the columns a
    // region carries never depend on which selector found it.
    const [found] = await resolveRegions(workbook, { sheet: "Data" }, "Region");
    expect(found?.columns).toEqual(detected?.columns);
    const [searched] = await resolveRegions(
      workbook,
      { find: "Amount" },
      "Amount",
    );
    expect(searched?.columns).toEqual(detected?.columns);
  });

  it("declares a header row without dropping the columns under it", async () => {
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Report title", null],
            ["Region", "Amount"],
            ["North", 10],
          ],
          name: "Data",
        },
      ],
    });

    // Declaring the title row as the header reads the title as a header cell
    // and the real header as data; the second column holds values under it,
    // so it stays as an unnamed column.
    const [region] = await resolveRegions(workbook, {
      headerRow: 1,
      sheet: "Data",
    });
    expect(region?.origin).toEqual({ kind: "declared-header" });
    expect(region?.columns).toEqual([
      { index: 0, name: "Report title" },
      { index: 1, name: "" },
    ]);
  });

  it("looks for the column on the detected header row before anywhere else", async () => {
    // A title block that repeats the column's name above the real header:
    // the topmost match would key on the title line. The detected header row
    // is looked at first, so the split keys on the row the inspection
    // reports, with every selector.
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Region", "North", null, null],
            [null, null, null, null],
            ["Case_ID", "Region", "Owner", "Status"],
            ["R-1", "north", "Reviewer 1", "open"],
          ],
          name: "Data",
        },
      ],
    });

    const [bySheet] = await resolveRegions(
      workbook,
      { sheet: "Data" },
      "Region",
    );
    expect(bySheet?.headerRow).toBe(3);
    expect(bySheet?.columns.map((column) => column.name)).toEqual([
      "Case_ID",
      "Region",
      "Owner",
      "Status",
    ]);
    const [byFind] = await resolveRegions(
      workbook,
      { find: "Region" },
      "Region",
    );
    expect(byFind?.headerRow).toBe(3);
    const [everywhere] = await resolveRegions(
      workbook,
      "all-worksheets",
      "Region",
    );
    expect(everywhere?.headerRow).toBe(3);
    // A declared header row is never second-guessed.
    const [declared] = await resolveRegions(
      workbook,
      { headerRow: 1, sheet: "Data" },
      "Region",
    );
    expect(declared?.headerRow).toBe(1);
  });

  it("falls back to the topmost match when the detected header row lacks the column", async () => {
    // The rule places the header on row 1 here (a one-value line directly
    // above a three-column header is not evidence of a title), so the column
    // is not on the detected row; the whole-sheet search still finds it on
    // row 2, as it always did.
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Quarterly review log", null, null],
            ["Case_ID", "Region", "Status"],
            ["R-1", "north", "open"],
          ],
          name: "Data",
        },
      ],
    });

    const [region] = await resolveRegions(
      workbook,
      { sheet: "Data" },
      "Region",
    );
    expect(region?.headerRow).toBe(2);
  });

  it("skips a banner merged across the columns directly above the header", async () => {
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Quarterly review log", null, null, null],
            ["Case_ID", "Region", "Owner", "Status"],
            ["R-1", "north", "Reviewer 1", "open"],
          ],
          merges: ["A1:D1"],
          name: "Data",
        },
      ],
    });

    const [region] = await resolveRegions(workbook, { sheet: "Data" });
    expect(region?.headerRow).toBe(2);
    expect(region?.columns.map((column) => column.name)).toEqual([
      "Case_ID",
      "Region",
      "Owner",
      "Status",
    ]);

    // The same title typed into one unmerged cell is no evidence: it could
    // be a header naming one column, so it reads as the header, as before.
    const unmerged = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Quarterly review log", null, null, null],
            ["Case_ID", "Region", "Owner", "Status"],
            ["R-1", "north", "Reviewer 1", "open"],
          ],
          name: "Data",
        },
      ],
    });
    const [kept] = await resolveRegions(unmerged, { sheet: "Data" });
    expect(kept?.headerRow).toBe(1);
  });

  it("finds a column by the name the inspection gives it, whatever the cell's type", async () => {
    // A boolean header cell is stored as 1 and named `true`; a numeric one
    // stored as 1E3 is named 1000. The search matches the name, so a column
    // chosen from an inspection is found by every selector.
    const workbook = new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            [{ text: "1", type: "b" }, { text: "1E3" }, "Region"],
            ["yes", 5, "north"],
          ],
          name: "Data",
        },
      ],
    });

    const [byName] = await resolveRegions(workbook, { sheet: "Data" }, "true");
    expect(byName?.columns.map((column) => column.name)).toEqual([
      "true",
      "1000",
      "Region",
    ]);
    const [numeric] = await resolveRegions(workbook, { find: "1000" }, "1000");
    expect(numeric?.headerRow).toBe(1);
    await expect(
      resolveRegions(workbook, { sheet: "Data" }, "1"),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_COLUMN_NOT_FOUND" });
  });
});

describe("resolveRegions: { range }", () => {
  function namedWorkbook(): FakeWorkbookModel {
    return new FakeWorkbookModel({
      definedNames: [
        { name: "SalesData", reference: "Data!$A$1:$B$3" },
        { name: "_xlnm.Print_Area", reference: "Data!$A$1:$B$9" },
        { name: "Broken", reference: "Data!not-a-range" },
      ],
      sheets: [
        {
          grid: [
            ["Region", "Amount"],
            ["North", 10],
            ["South", 20],
            ["Excluded", 99],
          ],
          name: "Data",
        },
        {
          grid: [["Region"], ["East"]],
          name: "My Sheet",
        },
      ],
    });
  }

  it("resolves a defined name case-insensitively", async () => {
    const [region] = await resolveRegions(namedWorkbook(), {
      range: "salesdata",
    });

    expect(region).toBeInstanceOf(RangeBinding);
    expect(region?.origin).toEqual({
      kind: "named-range",
      rangeName: "SalesData",
    });
    expect(region?.headerRow).toBe(1);
    // Row 4 is outside the named range and stays out of the body.
    expect(region?.body).toEqual({
      end: { column: 1, row: 3 },
      start: { column: 0, row: 2 },
    });
  });

  it("resolves an explicit sheet-qualified range", async () => {
    const [region] = await resolveRegions(namedWorkbook(), {
      range: "Data!A1:B3",
    });

    expect(region?.origin).toEqual({
      kind: "explicit-range",
      reference: "Data!A1:B3",
    });
    expect(region?.body.end.row).toBe(3);
  });

  it("resolves a quoted sheet name", async () => {
    const [region] = await resolveRegions(namedWorkbook(), {
      range: "'My Sheet'!A1:A2",
    });

    expect(region?.sheetName).toBe("My Sheet");
  });

  it("ignores built-in defined names", async () => {
    await expect(
      resolveRegions(namedWorkbook(), { range: "_xlnm.Print_Area" }),
    ).rejects.toMatchObject({ code: "XLSX_INVALID_NAMED_RANGE" });
  });

  it("refuses a defined name that is not a single worksheet range", async () => {
    await expect(
      resolveRegions(namedWorkbook(), { range: "Broken" }),
    ).rejects.toMatchObject({ code: "XLSX_INVALID_NAMED_RANGE" });
  });

  it("refuses text that is neither a defined name nor a reference", async () => {
    await expect(
      resolveRegions(namedWorkbook(), { range: "A1:B3" }),
    ).rejects.toMatchObject({ code: "XLSX_INVALID_NAMED_RANGE" });
  });

  it("refuses a range on a sheet that does not exist", async () => {
    await expect(
      resolveRegions(namedWorkbook(), { range: "Missing!A1:B3" }),
    ).rejects.toMatchObject({ code: "XLSX_WORKSHEET_NOT_FOUND" });
  });
});

describe("resolveRegions: { table }", () => {
  function tableWorkbook(): FakeWorkbookModel {
    return new FakeWorkbookModel({
      sheets: [
        {
          grid: [
            ["Region", "Amount"],
            ["North", 10],
            ["South", 20],
          ],
          name: "Data",
        },
      ],
      tables: [
        fakeTable({
          columnNames: ["Region", "Amount"],
          name: "Sales",
          ref: "A1:B3",
          sheetName: "Data",
        }),
      ],
    });
  }

  it("resolves a table by name, case-insensitively", async () => {
    const [region] = await resolveRegions(tableWorkbook(), { table: "sales" });

    expect(region).toBeInstanceOf(TableBinding);
    expect(region?.origin).toEqual({ kind: "table", tableName: "Sales" });
    expect(region?.body).toEqual({
      end: { column: 1, row: 3 },
      start: { column: 0, row: 2 },
    });
  });

  it("refuses an unknown table", async () => {
    await expect(
      resolveRegions(tableWorkbook(), { table: "Missing" }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_NO_TABLE" });
  });
});
