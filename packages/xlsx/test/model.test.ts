/**
 * L1 unit tests: the document model's row-edit invariant pass.
 *
 * Every assertion here is one of the invariants `deleteRows` promises. They
 * are written against the model directly rather than through an operation, so
 * a failure names the layer that broke.
 */
import { describe, expect, it } from "vitest";

import { WorkbookModel, RowRelocation } from "../src/model/index.js";
import { WorkbookPackage } from "../src/package/index.js";
import {
  buildCorpusWorkbook,
  calcChainReferences,
  conditionalFormattingSqref,
  CORPUS_PARTS,
  CORPUS_SHEET,
  dataValidationSqref,
  hyperlinkReferences,
  mergedCellReferences,
  readPackagePart,
  worksheetCellFormula,
  worksheetCellValue,
  worksheetRowNumbers,
} from "./corpus/fixtures.js";

/**
 * The corpus places Alpha's records on rows 4, 6 and 9. Removing everything
 * else from the data area is the movement every dependent-reference assertion
 * in this file is built on: 4 stays, 6 lands on 5, 9 lands on 6.
 */
const REMOVED_DATA_ROWS = [5, 7, 8, 10, 11, 12];

async function deleteAlphaRows(options: {
  sharedFormula?: boolean;
}): Promise<{ dataXml: string; calcChain: string; bytes: Uint8Array }> {
  const source = await buildCorpusWorkbook({
    shape: "range",
    sharedFormula: options.sharedFormula ?? false,
  });
  const model = await WorkbookModel.load(source);
  model.deleteRows(CORPUS_SHEET, REMOVED_DATA_ROWS, { renumber: true });
  const bytes = await model.save();

  return {
    bytes,
    calcChain: await readPackagePart(bytes, CORPUS_PARTS.calcChain),
    dataXml: await readPackagePart(bytes, CORPUS_PARTS.dataSheet),
  };
}

describe("model: deleteRows maintains row-dependent invariants", () => {
  it("renumbers surviving rows and the cell references inside them", async () => {
    const { dataXml } = await deleteAlphaRows({});

    expect(worksheetRowNumbers(dataXml)).toEqual([1, 3, 4, 5, 6]);
    expect(worksheetCellValue(dataXml, "D4")).toBe("10");
    expect(worksheetCellValue(dataXml, "D5")).toBe("30");
    expect(worksheetCellValue(dataXml, "D6")).toBe("60");
    // The records keep their identity while their rows compact.
    expect(worksheetCellValue(dataXml, "A4")).toBe("1");
    expect(worksheetCellValue(dataXml, "A5")).toBe("3");
    expect(worksheetCellValue(dataXml, "A6")).toBe("6");
  });

  it("rewrites A1 formula text so each formula still reads its own row", async () => {
    const { dataXml } = await deleteAlphaRows({});

    expect(worksheetCellFormula(dataXml, "E4")).toBe("D4*2");
    expect(worksheetCellFormula(dataXml, "E5")).toBe("D5*2");
    expect(worksheetCellFormula(dataXml, "E6")).toBe("D6*2");
  });

  it("shrinks a shared formula's span with the rows it covers", async () => {
    const { dataXml } = await deleteAlphaRows({ sharedFormula: true });

    expect(dataXml).toContain('ref="F4:F6"');
    expect(dataXml).not.toContain('ref="F4:F9"');
  });

  it("moves merged ranges with the rows they cover", async () => {
    const { dataXml } = await deleteAlphaRows({});

    expect(mergedCellReferences(dataXml)).toEqual(["A1:F1", "H5:I5"]);
    // The container's count stays honest.
    expect(dataXml).toContain('<mergeCells count="2">');
  });

  it("shrinks conditional-formatting and data-validation sqrefs", async () => {
    const { dataXml } = await deleteAlphaRows({});

    expect(conditionalFormattingSqref(dataXml)).toBe("D4:D6");
    expect(dataValidationSqref(dataXml)).toBe("A4:A6");
  });

  it("moves hyperlinks with the row they decorate", async () => {
    const { dataXml } = await deleteAlphaRows({});

    expect(hyperlinkReferences(dataXml)).toEqual(["A5"]);
  });

  it("drops calculation-chain entries for deleted cells and renumbers the rest", async () => {
    const { calcChain } = await deleteAlphaRows({});
    const references = calcChainReferences(calcChain);

    expect(references.filter((reference) => reference.startsWith("E"))).toEqual(
      ["E4", "E5", "E6"],
    );
    // Row 12 left, so the footer's chain entry left with it.
    expect(references).not.toContain("B12");
    // Another worksheet's entries are not this edit's business.
    expect(references).toContain("B2");
    expect(references).toContain("B3");
  });

  it("moves comment anchors and their VML shapes with the row", async () => {
    const { bytes } = await deleteAlphaRows({});

    // The note sits on B6, an Alpha record that lands on row 5.
    expect(await readPackagePart(bytes, CORPUS_PARTS.comments)).toContain(
      '<comment ref="B5"',
    );
    // The legacy drawing anchors the same note by a zero-based row.
    expect(
      await readPackagePart(bytes, "xl/drawings/vmlDrawing1.vml"),
    ).toContain("<x:Row>4</x:Row>");
  });

  it("removes a comment whose row was deleted", async () => {
    const source = await buildCorpusWorkbook({ shape: "range" });
    const model = await WorkbookModel.load(source);
    // Row 6 carries the only comment; deleting it must take the note too.
    model.deleteRows(CORPUS_SHEET, [6], { renumber: true });
    const bytes = await model.save();

    expect(await readPackagePart(bytes, CORPUS_PARTS.comments)).not.toContain(
      "<comment ",
    );
    expect(
      await readPackagePart(bytes, "xl/drawings/vmlDrawing1.vml"),
    ).not.toContain("<x:Row>");
  });

  it("leaves cross-sheet formulas alone", async () => {
    const { bytes } = await deleteAlphaRows({});
    const summary = await readPackagePart(bytes, CORPUS_PARTS.summarySheet);

    // Summary!B2 describes the source workbook's geometry. Rewriting it here
    // would silently restate an aggregate the model cannot recompute.
    expect(worksheetCellFormula(summary, "B2")).toBe(
      `SUM(${CORPUS_SHEET}!D4:D9)`,
    );
  });

  it("leaves parts it did not edit byte-identical", async () => {
    const source = await buildCorpusWorkbook({ shape: "range", pivot: true });
    const model = await WorkbookModel.load(source);
    model.deleteRows(CORPUS_SHEET, REMOVED_DATA_ROWS, { renumber: true });
    const bytes = await model.save();

    // Comments and their VML drawing are deliberately absent: they anchor on
    // rows this edit moved, so the invariant pass maintains them.
    for (const part of [
      CORPUS_PARTS.summarySheet,
      CORPUS_PARTS.veryHiddenSheet,
      CORPUS_PARTS.sharedStrings,
      CORPUS_PARTS.pivotCacheRecords,
      "xl/styles.xml",
    ]) {
      expect(await readPackagePart(bytes, part)).toBe(
        await readPackagePart(source, part),
      );
    }
  });

  it("deletes without renumbering when asked, leaving survivors in place", async () => {
    const source = await buildCorpusWorkbook({ shape: "range" });
    const model = await WorkbookModel.load(source);
    model.deleteRows(CORPUS_SHEET, [5, 7, 8], { renumber: false });
    const dataXml = await readPackagePart(
      await model.save(),
      CORPUS_PARTS.dataSheet,
    );

    expect(worksheetRowNumbers(dataXml)).toEqual([1, 3, 4, 6, 9, 12]);
    expect(worksheetCellFormula(dataXml, "E6")).toBe("D6*2");
    // The conditional-formatting span still ends on the last surviving row.
    expect(conditionalFormattingSqref(dataXml)).toBe("D4:D9");
  });
});

describe("model: relocating an Excel Table's rows", () => {
  it("compacts the data rows, follows the totals row and resizes the table", async () => {
    const source = await buildCorpusWorkbook({
      shape: "table",
      formulas: "structured",
    });
    const model = await WorkbookModel.load(source);
    const table = model.tableDefinitions[0]!;

    // Alpha keeps rows 4, 6 and 9; the totals row on 10 follows them up.
    model.relocateRows(
      CORPUS_SHEET,
      RowRelocation.explicit([
        [4, 4],
        [5, null],
        [6, 5],
        [7, null],
        [8, null],
        [9, 6],
        [10, 7],
      ]),
      [table],
    );
    const bytes = await model.save();
    const dataXml = await readPackagePart(bytes, CORPUS_PARTS.dataSheet);
    const tableXml = await readPackagePart(bytes, CORPUS_PARTS.table);

    // Rows below the table stay exactly where they were.
    expect(worksheetRowNumbers(dataXml)).toEqual([1, 3, 4, 5, 6, 7, 12]);
    expect(tableXml).toContain('ref="A3:F7"');
    expect(tableXml).toContain('<autoFilter ref="A3:F6"/>');
    // The dependent structures follow the same plan the rows did.
    expect(mergedCellReferences(dataXml)).toEqual(["A1:F1", "H5:I5"]);
    expect(conditionalFormattingSqref(dataXml)).toBe("D4:D6");
    expect(dataValidationSqref(dataXml)).toBe("A4:A6");
    expect(hyperlinkReferences(dataXml)).toEqual(["A5"]);
  });

  it("resizes the worksheet's tables by default", async () => {
    const source = await buildCorpusWorkbook({
      shape: "table",
      formulas: "structured",
    });
    const model = await WorkbookModel.load(source);
    const report = model.deleteRows(CORPUS_SHEET, [5, 7, 8], {
      renumber: true,
    });
    const tableXml = await readPackagePart(
      await model.save(),
      CORPUS_PARTS.table,
    );

    expect(tableXml).toContain('ref="A3:F7"');
    expect(tableXml).toContain('<autoFilter ref="A3:F6"/>');
    expect(report.adjusted.tableRefs).toBe(1);
  });

  it("leaves the table claiming its original range when scoped out", async () => {
    const source = await buildCorpusWorkbook({
      shape: "table",
      formulas: "structured",
    });
    const model = await WorkbookModel.load(source);
    // A binding that deliberately does not move a table passes no tables.
    const report = model.deleteRows(CORPUS_SHEET, [5, 7, 8], {
      renumber: true,
      tables: [],
    });
    const tableXml = await readPackagePart(
      await model.save(),
      CORPUS_PARTS.table,
    );

    expect(tableXml).toContain('ref="A3:F10"');
    expect(report.adjusted.tableRefs).toBe(0);
  });

  it("refuses a relocation that would reorder rows", async () => {
    const source = await buildCorpusWorkbook({ shape: "range" });
    const model = await WorkbookModel.load(source);

    expect(() =>
      model.relocateRows(
        CORPUS_SHEET,
        RowRelocation.explicit([
          [4, 9],
          [9, 4],
        ]),
      ),
    ).toThrowError(/reorder worksheet rows/u);
  });
});

describe("model: reading structure", () => {
  it("exposes worksheets, their part paths and their visibility", async () => {
    const model = await WorkbookModel.load(
      await buildCorpusWorkbook({ shape: "range" }),
    );

    expect(model.sheets.map((sheet) => sheet.name)).toEqual([
      "Data",
      "Summary",
      "Hidden",
      "VeryHidden",
    ]);
    expect(model.sheets.map((sheet) => sheet.visibility)).toEqual([
      "visible",
      "visible",
      "hidden",
      "veryHidden",
    ]);
    expect(model.sheets[0]?.partPath).toBe(CORPUS_PARTS.dataSheet);
    // The workbook's own sheet id, which calculation-chain entries index by.
    expect(model.sheetEntries[0]?.sheetId).toBe(1);
  });

  it("exposes defined names, tables and shared strings", async () => {
    const model = await WorkbookModel.load(
      await buildCorpusWorkbook({ shape: "table", formulas: "structured" }),
    );
    const definedNames = await model.definedNames();
    const tables = await model.tables();

    expect(definedNames.map((name) => name.name)).toContain("CorpusRange");
    expect(
      definedNames.find((name) => name.name === "LocalNote")?.localSheetId,
    ).toBe(0);
    expect(tables.map((table) => table.name)).toEqual(["DataTable"]);
    expect(tables[0]).toMatchObject({
      sheetName: CORPUS_SHEET,
      partPath: CORPUS_PARTS.table,
      headerRow: 3,
      totalsRow: true,
    });
    expect(tables[0]?.range).toEqual({
      start: { row: 3, column: 0 },
      end: { row: 10, column: 5 },
    });
    expect(model.sharedStrings()).toContain("Alpha side note");
  });

  it("exposes rows, cells and their formulas as parsed structure", async () => {
    const model = await WorkbookModel.load(
      await buildCorpusWorkbook({ shape: "range", sharedFormula: true }),
    );
    const worksheet = model.worksheet(CORPUS_SHEET)!;

    expect(worksheet.rows().map((row) => row.number)).toEqual([
      1, 3, 4, 5, 6, 7, 8, 9, 12,
    ]);
    expect(worksheet.info.name).toBe(CORPUS_SHEET);
    expect(worksheet.usedRange).toEqual({
      start: { row: 1, column: 0 },
      end: { row: 12, column: 8 },
    });

    const amount = worksheet
      .row(4)!
      .cells.find((cell) => cell.ref.column === 3)!;
    expect(amount).toMatchObject({ ref: { row: 4, column: 3 }, value: "10" });
    expect(worksheet.cellText({ row: 4, column: 3 })).toBe("10");
    expect(worksheet.cellText({ row: 99, column: 0 })).toBeUndefined();

    const doubled = worksheet
      .row(4)!
      .cells.find((cell) => cell.ref.column === 4)!;
    expect(doubled.formula).toMatchObject({ kind: "normal", text: "D4*2" });

    const ratio = worksheet
      .row(4)!
      .cells.find((cell) => cell.ref.column === 5)!;
    expect(ratio.formula).toMatchObject({ kind: "shared", sharedIndex: 0 });
    expect(ratio.formula?.range).toEqual({
      start: { row: 4, column: 5 },
      end: { row: 9, column: 5 },
    });

    // A styled header cell keeps its type and style index.
    expect(worksheet.row(3)!.cells[0]).toMatchObject({
      type: "s",
      styleIndex: 1,
    });
    expect(worksheet.row(99)).toBeUndefined();
  });

  it("reports which structures a row edit adjusted", async () => {
    const model = await WorkbookModel.load(
      await buildCorpusWorkbook({ shape: "range", sharedFormula: true }),
    );
    const report = model.deleteRows(CORPUS_SHEET, REMOVED_DATA_ROWS, {
      renumber: true,
    });

    expect(report).toMatchObject({ deletedRows: 4, retainedRows: 5 });
    expect(report.adjusted.mergedRanges).toBe(1);
    expect(report.adjusted.conditionalFormatting).toBe(1);
    expect(report.adjusted.dataValidations).toBe(1);
    expect(report.adjusted.hyperlinks).toBe(1);
    expect(report.adjusted.formulaReferences).toBeGreaterThan(0);
    expect(report.adjusted.calcChainEntries).toBeGreaterThan(0);
    expect(report.adjusted.tableRefs).toBe(0);
  });

  it("reports an unknown worksheet rather than guessing", async () => {
    const model = await WorkbookModel.load(
      await buildCorpusWorkbook({ shape: "range" }),
    );

    expect(model.worksheet("Missing")).toBeUndefined();
    expect(() =>
      model.deleteRows("Missing", [1], { renumber: true }),
    ).toThrowError(/not in this workbook/u);
  });

  it("reads a value view lazily when none was supplied", async () => {
    const model = await WorkbookModel.load(
      await buildCorpusWorkbook({ shape: "range" }),
    );

    expect(model.values().SheetNames).toContain(CORPUS_SHEET);
  });

  it("has no value view when it was built from a package alone", async () => {
    const workbookPackage = await WorkbookPackage.load(
      await buildCorpusWorkbook({ shape: "range" }),
    );

    expect(() =>
      WorkbookModel.fromPackage(workbookPackage).values(),
    ).toThrowError(/no value view/u);
  });
});

describe("model: deterministic output", () => {
  it("produces identical bytes for identical edits", async () => {
    const source = await buildCorpusWorkbook({ shape: "range" });

    const runs: Uint8Array[] = [];
    for (let run = 0; run < 2; run += 1) {
      const model = await WorkbookModel.load(source);
      model.deleteRows(CORPUS_SHEET, REMOVED_DATA_ROWS, { renumber: true });
      runs.push(await model.save());
    }

    expect([...runs[0]!]).toEqual([...runs[1]!]);
  });
});
