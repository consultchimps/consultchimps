import { describe, expect, it } from "vitest";

import { TableBinding } from "../../src/region/table-binding.js";
import { FakeWorksheetModel, fakeTable, formulaCell } from "./fakes.js";

const TOLERANT = { strict: false };
const STRICT = { strict: true };

function salesSheet(): FakeWorksheetModel {
  return new FakeWorksheetModel({
    grid: [
      ["Region", "Amount"],
      ["North", 10],
      ["South", 20],
      ["North", 30],
      [formulaCell("SUBTOTAL(109,Sales[Amount])"), 60],
    ],
    name: "Data",
  });
}

function salesTable(): TableBinding {
  return new TableBinding(
    salesSheet(),
    fakeTable({
      columnNames: ["Region", "Amount"],
      name: "Sales",
      ref: "A1:B5",
      sheetName: "Data",
      totalsRow: true,
    }),
  );
}

describe("TableBinding", () => {
  it("describes itself from the table definition", () => {
    const binding = salesTable();

    expect(binding.sheetName).toBe("Data");
    expect(binding.origin).toEqual({ kind: "table", tableName: "Sales" });
    expect(binding.headerRow).toBe(1);
    expect(binding.totalsRow).toBe(true);
    expect(binding.columns).toEqual([
      { index: 0, name: "Region" },
      { index: 1, name: "Amount" },
    ]);
  });

  it("excludes the totals row from the body", () => {
    expect(salesTable().body).toEqual({
      end: { column: 1, row: 4 },
      start: { column: 0, row: 2 },
    });
  });

  it("includes the last row when there is no totals row", () => {
    const binding = new TableBinding(
      salesSheet(),
      fakeTable({
        columnNames: ["Region", "Amount"],
        name: "Sales",
        ref: "A1:B5",
        sheetName: "Data",
      }),
    );

    expect(binding.body.end.row).toBe(5);
  });

  it("offsets column indexes by the table's first column", () => {
    const binding = new TableBinding(
      new FakeWorksheetModel({
        firstColumn: 2,
        grid: [
          ["Region", "Amount"],
          ["North", 10],
        ],
        name: "Data",
      }),
      fakeTable({
        columnNames: ["Region", "Amount"],
        name: "Sales",
        ref: "C1:D2",
        sheetName: "Data",
      }),
    );

    expect(binding.columns).toEqual([
      { index: 2, name: "Region" },
      { index: 3, name: "Amount" },
    ]);
  });

  it("keys every body row, and only body rows", () => {
    const binding = salesTable();
    const keys = binding.rowKeys(binding.columns[0]!, TOLERANT);

    expect([...keys]).toEqual([
      [2, "string:north"],
      [3, "string:south"],
      [4, "string:north"],
    ]);
  });

  it("reports blank keys as undefined without dropping the row", () => {
    const worksheet = new FakeWorksheetModel({
      grid: [
        ["Region", "Amount"],
        ["North", 10],
        ["   ", 20],
        [null, 30],
      ],
      name: "Data",
    });
    const binding = new TableBinding(
      worksheet,
      fakeTable({
        columnNames: ["Region", "Amount"],
        name: "Sales",
        ref: "A1:B4",
        sheetName: "Data",
      }),
    );

    expect([...binding.rowKeys(binding.columns[0]!, TOLERANT)]).toEqual([
      [2, "string:north"],
      [3, undefined],
      [4, undefined],
    ]);
  });

  it("separates values under strict matching that tolerant matching folds", () => {
    const worksheet = new FakeWorksheetModel({
      grid: [
        ["Region", "Code"],
        [" North ", "100"],
        ["north", 100],
      ],
      name: "Data",
    });
    const binding = new TableBinding(
      worksheet,
      fakeTable({
        columnNames: ["Region", "Code"],
        name: "Sales",
        ref: "A1:B3",
        sheetName: "Data",
      }),
    );
    const region = binding.columns[0]!;
    const code = binding.columns[1]!;

    expect([...binding.rowKeys(region, TOLERANT).values()]).toEqual([
      "string:north",
      "string:north",
    ]);
    expect([...binding.rowKeys(region, STRICT).values()]).toEqual([
      "string: North ",
      "string:north",
    ]);
    // Numeric text and a real number group together only when tolerant.
    expect([...binding.rowKeys(code, TOLERANT).values()]).toEqual([
      "number:100",
      "number:100",
    ]);
    expect([...binding.rowKeys(code, STRICT).values()]).toEqual([
      "string:100",
      "number:100",
    ]);
  });

  it("deletes exactly the unwanted body rows, with renumbering, and never the totals row", () => {
    const worksheet = salesSheet();
    const binding = new TableBinding(
      worksheet,
      fakeTable({
        columnNames: ["Region", "Amount"],
        name: "Sales",
        ref: "A1:B5",
        sheetName: "Data",
        totalsRow: true,
      }),
    );
    const keys = binding.rowKeys(binding.columns[0]!, TOLERANT);

    const report = binding.filterRows(
      (row) => keys.get(row) === "string:north",
    );

    expect(report).toEqual({ deletedRows: 1, retainedRows: 2 });
    expect(worksheet.deleteCalls).toEqual([{ renumber: true, rows: [3] }]);
  });

  it("deletes every body row when nothing is kept", () => {
    const worksheet = salesSheet();
    const binding = new TableBinding(
      worksheet,
      fakeTable({
        columnNames: ["Region", "Amount"],
        name: "Sales",
        ref: "A1:B5",
        sheetName: "Data",
        totalsRow: true,
      }),
    );

    const report = binding.filterRows(() => false);

    expect(report).toEqual({ deletedRows: 3, retainedRows: 0 });
    expect(worksheet.deleteCalls).toEqual([
      { renumber: true, rows: [2, 3, 4] },
    ]);
    // Header row 1 and totals row 5 survive.
    expect(worksheet.row(1)?.cells[0]?.value).toBe("Region");
    expect(worksheet.cellText({ column: 1, row: 2 })).toBe("60");
  });

  it("touches the model at all only when rows are removed", () => {
    const worksheet = salesSheet();
    const binding = new TableBinding(
      worksheet,
      fakeTable({
        columnNames: ["Region", "Amount"],
        name: "Sales",
        ref: "A1:B5",
        sheetName: "Data",
        totalsRow: true,
      }),
    );

    const report = binding.filterRows(() => true);

    expect(report).toEqual({ deletedRows: 0, retainedRows: 3 });
    expect(worksheet.deleteCalls).toEqual([]);
  });

  it("handles a table with a header and a totals row but no body", () => {
    const worksheet = new FakeWorksheetModel({
      grid: [
        ["Region", "Amount"],
        ["Total", 0],
      ],
      name: "Data",
    });
    const binding = new TableBinding(
      worksheet,
      fakeTable({
        columnNames: ["Region", "Amount"],
        name: "Sales",
        ref: "A1:B2",
        sheetName: "Data",
        totalsRow: true,
      }),
    );

    expect(binding.body.end.row).toBe(1);
    expect([...binding.rowKeys(binding.columns[0]!, TOLERANT)]).toEqual([]);
    expect(binding.filterRows(() => false)).toEqual({
      deletedRows: 0,
      retainedRows: 0,
    });
    expect(worksheet.deleteCalls).toEqual([]);
  });

  it("renumbers even when the sheet holds formulas, because structured references move with the table", () => {
    const worksheet = new FakeWorksheetModel({
      grid: [
        ["Region", "Amount"],
        ["North", 10],
        ["South", 20],
        [formulaCell("SUM(Sales[Amount])"), formulaCell("SUM(B2:B3)", "array")],
      ],
      name: "Data",
    });
    const binding = new TableBinding(
      worksheet,
      fakeTable({
        columnNames: ["Region", "Amount"],
        name: "Sales",
        ref: "A1:B4",
        sheetName: "Data",
        totalsRow: true,
      }),
    );

    binding.filterRows(() => false);

    expect(worksheet.deleteCalls[0]?.renumber).toBe(true);
  });
});
