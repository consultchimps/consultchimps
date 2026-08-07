/**
 * L1 unit tests: reading cell values, and the formula-reference adjustment the
 * whole row-edit design leans on.
 *
 * The reference tests exercise `relocateFormulaRows` directly, because every
 * higher layer's correctness reduces to it: a range that shrinks wrongly, a
 * sheet-qualified reference that is rewritten, or a function name mistaken for
 * a cell would each corrupt a workbook silently.
 */
import { describe, expect, it } from "vitest";

import {
  excelSerialToDate,
  isDateFormatCode,
  relocateFormulaRows,
  relocateReference,
  relocateSqref,
  RowRelocation,
  StyleTable,
  WorkbookModel,
} from "../src/model/index.js";
import { WorkbookPackage } from "../src/package/index.js";
import {
  buildCorpusWorkbook,
  CORPUS_PARTS,
  CORPUS_SHEET,
} from "./corpus/fixtures.js";

/** Alpha's movement: 4 stays, 6 lands on 5, 9 lands on 6; 5, 7 and 8 leave. */
const ALPHA = RowRelocation.explicit([
  [4, 4],
  [5, null],
  [6, 5],
  [7, null],
  [8, null],
  [9, 6],
]);

describe("references: formula text", () => {
  it("rewrites a plain reference to the row it now lives on", () => {
    expect(relocateFormulaRows("D6*2", ALPHA)).toBe("D5*2");
    expect(relocateFormulaRows("D9+D4", ALPHA)).toBe("D6+D4");
  });

  it("keeps absolute markers while moving the row", () => {
    expect(relocateFormulaRows("$D$9", ALPHA)).toBe("$D$6");
    expect(relocateFormulaRows("D$9", ALPHA)).toBe("D$6");
    expect(relocateFormulaRows("$D9", ALPHA)).toBe("$D6");
  });

  it("shrinks a range to its surviving extent rather than moving endpoints", () => {
    // Rows 4 to 9 held six records; three survive, on rows 4 to 6.
    expect(relocateFormulaRows("SUM(D4:D9)", ALPHA)).toBe("SUM(D4:D6)");
    // Rows 5 to 8 held one surviving record, which moved to row 5, so the
    // range collapses onto it rather than keeping either stale endpoint.
    expect(relocateFormulaRows("SUM(D5:D8)", ALPHA)).toBe("SUM(D5:D5)");
  });

  it("collapses a reference to rows that all left into #REF!", () => {
    expect(relocateFormulaRows("D7*2", ALPHA)).toBe("#REF!*2");
    expect(relocateFormulaRows("SUM(D7:D8)", ALPHA)).toBe("SUM(#REF!)");
  });

  it("leaves sheet-qualified references alone", () => {
    // Another worksheet describes the source workbook's geometry; rewriting
    // it here would silently restate a cross-sheet aggregate.
    expect(relocateFormulaRows("SUM(Data!D4:D9)", ALPHA)).toBe(
      "SUM(Data!D4:D9)",
    );
    expect(relocateFormulaRows("'My Sheet'!D9", ALPHA)).toBe("'My Sheet'!D9");
    expect(relocateFormulaRows("[1]Other!D9", ALPHA)).toBe("[1]Other!D9");
  });

  it("does not mistake function names or numbers for references", () => {
    expect(relocateFormulaRows("LOG10(D9)", ALPHA)).toBe("LOG10(D6)");
    expect(relocateFormulaRows("1E9+D9", ALPHA)).toBe("1E9+D6");
    expect(relocateFormulaRows("SUM(D9)*100", ALPHA)).toBe("SUM(D6)*100");
  });

  it("leaves string literals and structured references untouched", () => {
    expect(relocateFormulaRows('IF(D9=1,"D9","D7")', ALPHA)).toBe(
      'IF(D6=1,"D9","D7")',
    );
    expect(
      relocateFormulaRows("DataTable[[#This Row],[Amount]]*2", ALPHA),
    ).toBe("DataTable[[#This Row],[Amount]]*2");
    expect(relocateFormulaRows("SUBTOTAL(109,DataTable[Amount])", ALPHA)).toBe(
      "SUBTOTAL(109,DataTable[Amount])",
    );
  });

  it("is a no-op when the plan moves nothing", () => {
    const identity = RowRelocation.compacting(new Set(), 20, true);
    expect(relocateFormulaRows("SUM(D4:D9)", identity)).toBe("SUM(D4:D9)");
  });
});

describe("references: attributes and sqrefs", () => {
  it("relocates a shared or array formula span", () => {
    expect(relocateReference("F4:F9", ALPHA)).toBe("F4:F6");
    expect(relocateReference("G4:G4", ALPHA)).toBe("G4:G4");
  });

  it("relocates each range of an sqref list and drops the dead ones", () => {
    expect(relocateSqref("D4:D9 A7:A8 B4", ALPHA)).toBe("D4:D6 B4");
    expect(relocateSqref("A7:A8", ALPHA)).toBeUndefined();
  });

  it("handles a whole-column range without walking every row", () => {
    const compacting = RowRelocation.compacting(new Set([5, 7, 8]), 12, true);
    // Rows past the sheet's last row shift up by the number deleted.
    expect(relocateReference("D1:D1048576", compacting)).toBe("D1:D1048573");
  });
});

describe("styles: date detection", () => {
  it("recognizes the built-in date and time formats", () => {
    const styles = StyleTable.parse(
      '<styleSheet><cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="14"/>' +
        '<xf numFmtId="2"/></cellXfs></styleSheet>',
    );

    expect(styles.isDateStyle(1)).toBe(true);
    expect(styles.isDateStyle(0)).toBe(false);
    expect(styles.isDateStyle(2)).toBe(false);
    expect(styles.isDateStyle(undefined)).toBe(false);
    expect(styles.isDateStyle(99)).toBe(false);
  });

  it("recognizes a custom format whose code describes a date", () => {
    const styles = StyleTable.parse(
      '<styleSheet><numFmts count="2">' +
        '<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>' +
        '<numFmt numFmtId="165" formatCode="0.00&quot; months&quot;"/>' +
        '</numFmts><cellXfs count="2"><xf numFmtId="164"/><xf numFmtId="165"/>' +
        "</cellXfs></styleSheet>",
    );

    expect(styles.isDateStyle(0)).toBe(true);
    expect(styles.isDateStyle(1)).toBe(false);
  });

  it("reads date tokens only outside literals and bracketed sections", () => {
    expect(isDateFormatCode("yyyy-mm-dd")).toBe(true);
    expect(isDateFormatCode("[$-409]h:mm AM/PM")).toBe(true);
    expect(isDateFormatCode('#,##0.00" days"')).toBe(false);
    expect(isDateFormatCode("General")).toBe(false);
    expect(isDateFormatCode("0.00%")).toBe(false);
  });

  it("converts serials in both of Excel's date systems", () => {
    // 1 January 2024 is serial 45292 in the 1900 system.
    expect(excelSerialToDate(45292, false).getFullYear()).toBe(2024);
    expect(excelSerialToDate(45292, false).getMonth()).toBe(0);
    expect(excelSerialToDate(45292, false).getDate()).toBe(1);
    // The same day is 1462 days earlier in the 1904 system.
    expect(excelSerialToDate(45292 - 1462, true).getDate()).toBe(1);
    expect(excelSerialToDate(45292 - 1462, true).getFullYear()).toBe(2024);
    // Times survive as local-time components.
    expect(excelSerialToDate(45292.5, false).getHours()).toBe(12);
  });
});

describe("model: cell values", () => {
  it("resolves shared strings rather than handing back their index", async () => {
    const model = await WorkbookModel.load(
      await buildCorpusWorkbook({ shape: "range" }),
    );
    const worksheet = model.worksheet(CORPUS_SHEET)!;

    // C3 is the "Group" header, stored as a shared-string index.
    expect(worksheet.cellText({ row: 3, column: 2 })).toBe("Group");
    expect(worksheet.cellValue({ row: 3, column: 2 })).toBe("Group");
    expect(worksheet.cellValue({ row: 4, column: 2 })).toBe("Alpha");
  });

  it("types numbers, absent cells and blanks", async () => {
    const model = await WorkbookModel.load(
      await buildCorpusWorkbook({ shape: "range" }),
    );
    const worksheet = model.worksheet(CORPUS_SHEET)!;

    expect(worksheet.cellValue({ row: 4, column: 3 })).toBe(10);
    expect(worksheet.cellValue({ row: 99, column: 0 })).toBeUndefined();
    // I6 exists as a styled shell with no value at all.
    expect(worksheet.cellValue({ row: 6, column: 8 })).toBeUndefined();
  });

  it("reads a date-formatted number as a Date", async () => {
    // Give the corpus a date-formatted style and point one cell at it,
    // through L0, so the value read exercises the real styles.xml path.
    const workbookPackage = await WorkbookPackage.load(
      await buildCorpusWorkbook({ shape: "range" }),
    );
    // Appended, so the existing style indexes keep pointing where they did.
    workbookPackage.writeText(
      "xl/styles.xml",
      workbookPackage
        .requireText("xl/styles.xml")
        .replace(
          "</cellXfs>",
          '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>',
        )
        .replace('<cellXfs count="2">', '<cellXfs count="3">'),
    );
    workbookPackage.writeText(
      CORPUS_PARTS.dataSheet,
      workbookPackage
        .requireText(CORPUS_PARTS.dataSheet)
        .replace('<c r="D4"><v>10</v></c>', '<c r="D4" s="2"><v>45292</v></c>'),
    );

    const model = WorkbookModel.fromPackage(workbookPackage);
    const value = model.worksheet(CORPUS_SHEET)!.cellValue({
      row: 4,
      column: 3,
    });

    expect(value).toBeInstanceOf(Date);
    expect((value as Date).getFullYear()).toBe(2024);
    // A plain number in the same column keeps its number type.
    expect(
      model.worksheet(CORPUS_SHEET)!.cellValue({ row: 5, column: 3 }),
    ).toBe(20);
  });

  it("reports a headerless table's header row as 0", async () => {
    const workbookPackage = await WorkbookPackage.load(
      await buildCorpusWorkbook({ shape: "table", formulas: "structured" }),
    );
    workbookPackage.writeText(
      CORPUS_PARTS.table,
      workbookPackage
        .requireText(CORPUS_PARTS.table)
        .replace('headerRowCount="1"', 'headerRowCount="0"'),
    );
    const tables = await WorkbookModel.fromPackage(workbookPackage).tables();

    // The seam types headerRow as a row number, so "no header" is 0 and
    // callers read anything below 1 as headerless.
    expect(tables[0]?.headerRow).toBe(0);
  });

  it("reports a table with a header row by its first row", async () => {
    const model = await WorkbookModel.load(
      await buildCorpusWorkbook({ shape: "table", formulas: "structured" }),
    );

    expect((await model.tables())[0]?.headerRow).toBe(3);
  });
});
