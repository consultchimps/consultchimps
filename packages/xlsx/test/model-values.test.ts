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
import { calendarIsoText, utcCalendarParts } from "../src/model/calendar.js";
import { WorkbookPackage } from "../src/package/index.js";
import { normalizeSplitValue } from "../src/region/values.js";
import {
  buildCorpusWorkbook,
  CORPUS_PARTS,
  CORPUS_SHEET,
} from "./corpus/fixtures.js";
import { inZone, ZONES } from "./zones.js";

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
    // A serial names a calendar moment and carries no zone, so the UTC face is
    // the face the workbook wrote. Reading the local face would report a
    // different day to everybody who is not in UTC.
    expect(excelSerialToDate(45292, false).toISOString()).toBe(
      "2024-01-01T00:00:00.000Z",
    );
    // The same day is 1462 days earlier in the 1904 system.
    expect(excelSerialToDate(45292 - 1462, true).toISOString()).toBe(
      "2024-01-01T00:00:00.000Z",
    );
    // A fractional serial is the time of day, in that same face.
    expect(excelSerialToDate(45292.5, false).toISOString()).toBe(
      "2024-01-01T12:00:00.000Z",
    );
  });

  it("reads the same serial the same way in every time zone", () => {
    // The defect this pins down: the components used to be re-assembled in
    // local time, so `toISOString()` subtracted the host offset and a split
    // named its outputs after a different calendar day in every zone - the day
    // before the one in the cell, for everybody east of UTC.
    for (const [serial, date1904, expected] of [
      [45292, false, "2024-01-01T00:00:00.000Z"],
      [45292.75, false, "2024-01-01T18:00:00.000Z"],
      [45292 - 1462, true, "2024-01-01T00:00:00.000Z"],
    ] as const) {
      for (const zone of ZONES) {
        expect(
          inZone(zone, () => excelSerialToDate(serial, date1904).toISOString()),
        ).toBe(expected);
      }
    }
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
    // The UTC face, which is the face the serial named. Reading the local one
    // would report 2023 to anybody west of UTC.
    expect((value as Date).toISOString()).toBe("2024-01-01T00:00:00.000Z");
    // A plain number in the same column keeps its number type.
    expect(
      model.worksheet(CORPUS_SHEET)!.cellValue({ row: 5, column: 3 }),
    ).toBe(20);
  });

  /** The corpus with one cell replaced by a `t="d"` cell holding this text. */
  async function datedModel(text: string): Promise<WorkbookModel> {
    const workbookPackage = await WorkbookPackage.load(
      await buildCorpusWorkbook({ shape: "range" }),
    );
    workbookPackage.writeText(
      CORPUS_PARTS.dataSheet,
      workbookPackage
        .requireText(CORPUS_PARTS.dataSheet)
        .replace(
          '<c r="D4"><v>10</v></c>',
          `<c r="D4" t="d"><v>${text}</v></c>`,
        ),
    );
    return WorkbookModel.fromPackage(workbookPackage);
  }

  /** That cell's value, read in each zone and spelled the one way. */
  async function dated(text: string): Promise<string[]> {
    const model = await datedModel(text);
    return ZONES.map((zone) =>
      inZone(zone, () => {
        const value = model
          .worksheet(CORPUS_SHEET)!
          .cellValue({ row: 4, column: 3 });
        return value instanceof Date
          ? calendarIsoText(utcCalendarParts(value))
          : String(value);
      }),
    );
  }

  /** The same, read once, in UTC. */
  async function datedInUtc(text: string): Promise<string> {
    return (await dated(text))[0] as string;
  }

  it("reads a date cell the same way in every time zone", async () => {
    // A `t="d"` cell writes ISO 8601 with no zone on it, so `new Date(text)`
    // read it as a moment in the host's zone: the same cell was 18:00 in UTC,
    // 14:00 in UTC+4, and the next calendar day west of UTC.
    // A date and a time, a date on its own, and text that already carries a
    // zone and is therefore a moment rather than a calendar face.
    expect(await dated("2024-01-01T18:00:00")).toEqual(
      ZONES.map(() => "2024-01-01T18:00:00.000Z"),
    );
    expect(await dated("2024-01-01")).toEqual(
      ZONES.map(() => "2024-01-01T00:00:00.000Z"),
    );
    expect(await dated("2024-01-01T18:00:00Z")).toEqual(
      ZONES.map(() => "2024-01-01T18:00:00.000Z"),
    );
    // A written offset says how far ahead of UTC the clock that wrote it was.
    expect(await dated("2024-01-01T18:00:00+04:00")).toEqual(
      ZONES.map(() => "2024-01-01T14:00:00.000Z"),
    );
    // Text that is not a date at all is still handed back as text.
    expect(await dated("not a date")).toEqual(ZONES.map(() => "not a date"));
  });

  it("reads an early year as the year that was written", async () => {
    // The defect this pins down: a date constructor remaps a year from 0 to 99
    // into the twentieth century, so 0099 came back as 1999 and nothing said
    // so. The components are converted by arithmetic instead.
    expect(new Date(Date.UTC(99, 0, 1)).getUTCFullYear()).toBe(1999);

    expect(await datedInUtc("0099-01-01")).toBe("0099-01-01T00:00:00.000Z");
    expect(await datedInUtc("0001-01-01")).toBe("0001-01-01T00:00:00.000Z");
  });

  it("keeps text that names no moment as the text it is", async () => {
    // A date constructor normalises an out-of-range field rather than refusing
    // it: month 13 becomes January of the next year, and 30 February becomes
    // the first days of March. Neither is what the cell says. A value that is
    // not a date is carried as the text the file holds, the same way every
    // other unconvertible cell is, so nothing is invented and nothing is lost.
    expect(new Date(Date.UTC(2024, 12, 1)).toISOString()).toBe(
      "2025-01-01T00:00:00.000Z",
    );

    for (const text of [
      "2024-13-01",
      "2024-00-01",
      "2024-02-30",
      "2023-02-29",
      "2024-01-01T24:00:00",
      "2024-01-01T23:59:60",
      "2024-01-01T18:00:00+24:00",
    ]) {
      expect(await datedInUtc(text)).toBe(text);
    }
    // The day the calendar does have, in the year that has it.
    expect(await datedInUtc("2024-02-29")).toBe("2024-02-29T00:00:00.000Z");
  });

  it("refuses a moment whose offset moves it past the years it can write", async () => {
    // Both are good timestamps. Their UTC faces are the year 10000 and the
    // year before 0000, and ISO 8601 writes neither without an expanded year
    // that nothing downstream reads: not the `date` column in
    // `@consultchimps/db`, whose grammar spells a year with four digits, and
    // not a split's output filename. Judging only the components as written
    // would have handed the formatter a value it has no spelling for.
    expect(await datedInUtc("9999-12-31T23:30:00-01:00")).toBe(
      "9999-12-31T23:30:00-01:00",
    );
    expect(await datedInUtc("0000-01-01T00:30:00+01:00")).toBe(
      "0000-01-01T00:30:00+01:00",
    );
    // The same boundary days, with no offset to move them, still read as
    // dates, and so does an offset that stays inside the range.
    expect(await datedInUtc("9999-12-31T23:30:00")).toBe(
      "9999-12-31T23:30:00.000Z",
    );
    expect(await datedInUtc("0000-01-01T00:30:00")).toBe(
      "0000-01-01T00:30:00.000Z",
    );
    expect(await datedInUtc("0001-01-01T00:30:00+01:00")).toBe(
      "0000-12-31T23:30:00.000Z",
    );
  });

  it("keys a cell it could not read by the text, not by a broken date", async () => {
    // The key becomes an output workbook's filename, so a value that has no
    // spelling must not reach it as one.
    const model = await datedModel("9999-12-31T23:30:00-01:00");
    const key = inZone("UTC", () =>
      normalizeSplitValue(
        model.worksheet(CORPUS_SHEET)!.cellValue({ row: 4, column: 3 }),
        true,
      ),
    );

    expect(key?.display).toBe("9999-12-31T23:30:00-01:00");
    expect(key?.key).toBe("string:9999-12-31T23:30:00-01:00");
  });

  it("keys a date cell for a split the way the cell reads", async () => {
    // The group key becomes an output workbook's name, so it has to be the
    // same characters the cell reads as, in every zone.
    const model = await datedModel("2024-01-01T18:00:00");
    const keys = ZONES.map((zone) =>
      inZone(
        zone,
        () =>
          normalizeSplitValue(
            model.worksheet(CORPUS_SHEET)!.cellValue({ row: 4, column: 3 }),
            true,
          )?.key,
      ),
    );

    expect(keys).toEqual(ZONES.map(() => "date:2024-01-01T18:00:00.000Z"));
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
