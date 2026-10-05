import JSZip from "jszip";
import { afterEach, describe, expect, it } from "vitest";

import type { Table } from "@consultchimps/tabular";

import { readWorkbookTablesBytes } from "../src/bytes.js";
import { buildTableWorkbookBytes } from "../src/shared.js";
import {
  escapeCellText,
  TableWorkbookWriter,
} from "../src/package/table-writer.js";
import { sheetNamesOf, sheetRows } from "./support/read-workbook.js";

const originalTimeZone = process.env.TZ;
afterEach(() => {
  if (originalTimeZone === undefined) delete process.env.TZ;
  else process.env.TZ = originalTimeZone;
});

// One value of each kind the writer has to get right, including text that
// XML cannot carry and text that looks like OOXML's own escape.
const TRICKY: Table = {
  columns: ["Text", "Number", "Flag"],
  rows: [
    { Text: 'A & B <c> "d"', Number: 1.5, Flag: true },
    { Text: "line\u0001break\u001fend", Number: 1e21, Flag: false },
    { Text: "_x0041_ stays literal", Number: -0.000001, Flag: null },
    { Text: "_x005F_ also literal", Number: 0, Flag: true },
    { Text: "  padded  ", Number: -5, Flag: null },
    { Text: "emoji \u{1F600} and é", Number: null, Flag: false },
    { Text: "lone \uD800 surrogate", Number: 42, Flag: true },
    { Text: null, Number: 7, Flag: null },
  ],
};

async function part(bytes: Uint8Array, name: string): Promise<string | null> {
  const zip = await JSZip.loadAsync(bytes);
  return (await zip.file(name)?.async("string")) ?? null;
}

describe("table workbook writer", () => {
  it("round-trips every value through the cell reader and the table reader", async () => {
    const bytes = buildTableWorkbookBytes(TRICKY, "Data");
    const expected = TRICKY.rows.map((row) =>
      TRICKY.columns.map((column) => row[column] ?? null),
    );

    expect(await sheetRows(bytes, "Data")).toEqual([
      TRICKY.columns,
      ...expected,
    ]);

    const [table] = await readWorkbookTablesBytes({
      name: "out.xlsx",
      bytes,
    });
    expect(table!.columns).toEqual(TRICKY.columns);
    expect(
      table!.rows.map((row) =>
        TRICKY.columns.map((column) => row[column] ?? null),
      ),
    ).toEqual(expected);
  });

  it("stores text inline and keeps the dimension, filter, and widths", async () => {
    const bytes = buildTableWorkbookBytes(TRICKY, "Data");
    expect(await part(bytes, "xl/sharedStrings.xml")).toBeNull();
    const sheet = (await part(bytes, "xl/worksheets/sheet1.xml"))!;
    expect(sheet).toContain('<dimension ref="A1:C9"/>');
    expect(sheet).toContain('<autoFilter ref="A1:C9"/>');
    expect(sheet).toContain('t="inlineStr"');
    // The longest text, "_x0041_ stays literal", is 21 characters, so its
    // column is 23 wide; the rest keep the 10-character minimum or more.
    expect(sheet).toContain('<col min="1" max="1" width="23.83203125"');
    expect(await part(bytes, "xl/workbook.xml")).toContain("'Data'!$A$1:$C$9");
  });

  it("gives identical bytes on every run and in every time zone", () => {
    const first = buildTableWorkbookBytes(TRICKY, "Data");
    expect(buildTableWorkbookBytes(TRICKY, "Data")).toEqual(first);
    for (const zone of [
      "Asia/Dubai",
      "America/Los_Angeles",
      "Pacific/Kiritimati",
    ]) {
      process.env.TZ = zone;
      expect(buildTableWorkbookBytes(TRICKY, "Data"), zone).toEqual(first);
    }
  });

  it("names a sheet with spaces and apostrophes, and refuses names Excel rejects", async () => {
    const bytes = buildTableWorkbookBytes(TRICKY, "Q3 client's data");
    expect(await sheetNamesOf(bytes)).toEqual(["Q3 client's data"]);
    expect(await part(bytes, "xl/workbook.xml")).toContain(
      "'Q3 client''s data'!$A$1:$C$9",
    );

    for (const name of [
      "",
      "x".repeat(32),
      "a/b",
      "a[1]",
      "Q1\nData",
      "Q1\tData",
      "'quoted'",
      "History",
    ]) {
      expect(() => buildTableWorkbookBytes(TRICKY, name), name).toThrow(
        expect.objectContaining({ code: "XLSX_INVALID_SHEET_NAME" }),
      );
    }
  });

  it("writes a header-only table", async () => {
    const bytes = buildTableWorkbookBytes(
      { columns: ["Only"], rows: [] },
      "Data",
    );
    expect(await sheetRows(bytes, "Data")).toEqual([["Only"]]);
  });

  it("refuses a row count that does not match the rows written", () => {
    const writer = new TableWorkbookWriter({
      sheetName: "Data",
      columns: ["A"],
      widths: [10],
      rowCount: 2,
      onChunk: () => {},
    });
    writer.writeRow(["one"]);
    expect(() => writer.finish()).toThrow(/1 rows were written but 2/u);
    expect(() => writer.writeRow(["two"])).not.toThrow();
    expect(() => writer.writeRow(["three"])).toThrow(/More rows/u);
  });

  it("keeps carriage returns, which XML parsers would turn into line feeds", async () => {
    const table: Table = {
      columns: ["Note"],
      rows: [{ Note: "windows\r\nline" }, { Note: "lone\rreturn" }],
    };
    const bytes = buildTableWorkbookBytes(table, "Data");
    // The file carries every carriage return escaped, as Excel writes them.
    const sheet = (await part(bytes, "xl/worksheets/sheet1.xml"))!;
    expect(sheet).toContain("windows_x000D_\nline");
    expect(sheet).toContain("lone_x000D_return");
    expect(escapeCellText("a\r\nb")).toBe("a_x000D_\nb");
    // Both read back: a lone return, and one before a line feed.
    expect((await sheetRows(bytes, "Data")).slice(1)).toEqual([
      ["windows\r\nline"],
      ["lone\rreturn"],
    ]);
  });

  it("escapes OOXML's own escape sequence so it reads back literally", () => {
    expect(escapeCellText("_x0041_")).toBe("_x005F_x0041_");
    expect(escapeCellText("tab\tand newline\n stay")).toBe(
      "tab\tand newline\n stay",
    );
    expect(escapeCellText("plain")).toBe("plain");
  });
});
