import { describe, expect, it } from "vitest";

import type { RegionEditReport } from "../../src/region/types.js";
import {
  RangeBinding,
  evaluateFormulaGuard,
  isRangeEditReport,
  readHeaderColumns,
} from "../../src/region/range-binding.js";
import { FakeWorksheetModel, formulaCell } from "./fakes.js";

const TOLERANT = { strict: false };
const STRICT = { strict: true };

function sheetWithFormula(
  formula: ReturnType<typeof formulaCell> | undefined,
): FakeWorksheetModel {
  return new FakeWorksheetModel({
    grid: [
      ["Region", "Amount"],
      ["North", 10],
      ["South", 20],
      [null, formula ?? null],
    ],
    name: "Data",
  });
}

function bindingFor(worksheet: FakeWorksheetModel): RangeBinding {
  return new RangeBinding({
    body: {
      end: { column: 1, row: 3 },
      start: { column: 0, row: 2 },
    },
    headerRow: 1,
    origin: { kind: "detected-header" },
    worksheet,
  });
}

describe("readHeaderColumns", () => {
  it("reads names straight from the header cells, keeping blanks in place", () => {
    const worksheet = new FakeWorksheetModel({
      grid: [
        ["Region", "", "Amount"],
        ["North", "x", 10],
      ],
      name: "Data",
    });

    expect(readHeaderColumns(worksheet, 1, 0, 3)).toEqual([
      { index: 0, name: "Region" },
      { index: 1, name: "" },
      { index: 2, name: "Amount" },
      { index: 3, name: "" },
    ]);
  });
});

describe("RangeBinding", () => {
  it("describes itself from its boundaries and header cells", () => {
    const binding = bindingFor(sheetWithFormula(undefined));

    expect(binding.sheetName).toBe("Data");
    expect(binding.origin).toEqual({ kind: "detected-header" });
    expect(binding.headerRow).toBe(1);
    expect(binding.columns).toEqual([
      { index: 0, name: "Region" },
      { index: 1, name: "Amount" },
    ]);
  });

  it("accepts explicit columns instead of reading the header", () => {
    const binding = new RangeBinding({
      body: { end: { column: 1, row: 3 }, start: { column: 0, row: 2 } },
      columns: [{ index: 1, name: "Amount" }],
      headerRow: 1,
      origin: { kind: "explicit-range", reference: "Data!A1:B3" },
      worksheet: sheetWithFormula(undefined),
    });

    expect(binding.columns).toEqual([{ index: 1, name: "Amount" }]);
  });

  it("keys body rows, blank and strict semantics included", () => {
    const worksheet = new FakeWorksheetModel({
      grid: [["Region"], [" North "], ["north"], ["  "]],
      name: "Data",
    });
    const binding = new RangeBinding({
      body: { end: { column: 0, row: 4 }, start: { column: 0, row: 2 } },
      headerRow: 1,
      origin: { kind: "detected-header" },
      worksheet,
    });
    const column = binding.columns[0]!;

    expect([...binding.rowKeys(column, TOLERANT).values()]).toEqual([
      "string:north",
      "string:north",
      undefined,
    ]);
    expect([...binding.rowKeys(column, STRICT).values()]).toEqual([
      "string: North ",
      "string:north",
      undefined,
    ]);
  });
});

describe("RangeBinding formula guard", () => {
  it("allows renumbering when the sheet holds no formulas", () => {
    expect(evaluateFormulaGuard(sheetWithFormula(undefined))).toEqual({
      canRenumber: true,
    });
  });

  it("allows renumbering for structured references only", () => {
    expect(
      evaluateFormulaGuard(sheetWithFormula(formulaCell("SUM(Sales[Amount])"))),
    ).toEqual({ canRenumber: true });
  });

  it("allows renumbering when the only A1-looking text is inside a string literal", () => {
    expect(
      evaluateFormulaGuard(
        sheetWithFormula(formulaCell('CONCATENATE("A1",Sales[Region])')),
      ),
    ).toEqual({ canRenumber: true });
  });

  it("does not mistake a function name ending in digits for a reference", () => {
    expect(
      evaluateFormulaGuard(sheetWithFormula(formulaCell("LOG10(Sales[Rate])"))),
    ).toEqual({ canRenumber: true });
  });

  it("refuses renumbering for an A1 reference, naming the cell", () => {
    expect(
      evaluateFormulaGuard(sheetWithFormula(formulaCell("SUM(B2:B3)"))),
    ).toEqual({
      canRenumber: false,
      cell: "B4",
      reason: "a1-reference",
    });
  });

  it("refuses renumbering for an absolute A1 reference", () => {
    expect(
      evaluateFormulaGuard(sheetWithFormula(formulaCell("$B$2*2"))),
    ).toMatchObject({ canRenumber: false, reason: "a1-reference" });
  });

  it("refuses renumbering for a shared formula", () => {
    expect(
      evaluateFormulaGuard(
        sheetWithFormula(formulaCell("Sales[Amount]*2", "shared")),
      ),
    ).toEqual({
      canRenumber: false,
      cell: "B4",
      reason: "shared-formula",
    });
  });

  it("refuses renumbering for an array formula", () => {
    expect(
      evaluateFormulaGuard(
        sheetWithFormula(formulaCell("Sales[Amount]*2", "array")),
      ),
    ).toMatchObject({ canRenumber: false, reason: "array-formula" });
  });

  it("scans the whole sheet, not just the region body", () => {
    const worksheet = new FakeWorksheetModel({
      grid: [
        ["Region", "Amount"],
        ["North", 10],
        ["South", 20],
        [formulaCell("SUM(B2:B3)"), null],
      ],
      name: "Data",
    });

    expect(bindingFor(worksheet).formulaGuard().canRenumber).toBe(false);
  });
});

describe("RangeBinding.filterRows", () => {
  it("deletes with renumbering when no formula ties rows to positions", () => {
    const worksheet = sheetWithFormula(formulaCell("SUM(Sales[Amount])"));
    const binding = bindingFor(worksheet);

    const report = binding.filterRows((row) => row === 2);

    expect(report).toMatchObject({
      deletedRows: 1,
      renumbered: true,
      retainedRows: 1,
    });
    expect(report.formulaGuard).toEqual({ canRenumber: true });
    expect(worksheet.deleteCalls).toEqual([{ renumber: true, rows: [3] }]);
  });

  it("renumbers even when the guard finds a position-dependent formula", () => {
    const worksheet = sheetWithFormula(formulaCell("SUM(B2:B3)"));
    const binding = bindingFor(worksheet);

    const report = binding.filterRows((row) => row === 2);

    // The guard is a signal now, not a veto: L1 rewrites the reference as part
    // of the same edit, so leaving a gap would only make the output worse.
    expect(report).toMatchObject({ deletedRows: 1, renumbered: true });
    expect(report.formulaGuard).toMatchObject({ reason: "a1-reference" });
    expect(worksheet.deleteCalls).toEqual([{ renumber: true, rows: [3] }]);
  });

  it("leaves the model alone when every row is kept", () => {
    const worksheet = sheetWithFormula(undefined);

    const report = bindingFor(worksheet).filterRows(() => true);

    expect(report).toMatchObject({ deletedRows: 0, retainedRows: 2 });
    expect(worksheet.deleteCalls).toEqual([]);
  });

  it("keeps the guard decision reachable from a plain RegionEditReport", () => {
    const worksheet = sheetWithFormula(formulaCell("SUM(B2:B3)"));
    const report: RegionEditReport = bindingFor(worksheet).filterRows(
      () => false,
    );

    expect(isRangeEditReport(report)).toBe(true);
    expect(isRangeEditReport({ deletedRows: 0, retainedRows: 0 })).toBe(false);
  });
});
