/**
 * The streaming reader consolidation reads through (ADR 0006): what it makes
 * of each cell, and a cell-by-cell comparison with the SheetJS-backed reader it
 * replaces, whose answers it must give except where it deliberately differs.
 */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import {
  StreamedWorkbook,
  type StreamedCell,
  type StreamedValue,
} from "../src/operations/consolidate/reader.js";
import { CellError } from "../src/package/cell-error.js";
import { bytesSource } from "../src/package/index.js";
import { TableWorkbookWriter } from "../src/package/table-writer.js";
import {
  cellToPrimitive,
  parseWorkbookBytes,
  readWorkbookDates,
} from "../src/shared.js";
import { buildCorpusWorkbook } from "./corpus/fixtures.js";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";

interface HandSheet {
  name: string;
  data: string;
  before?: string;
  after?: string;
  state?: string;
}

interface HandWorkbook {
  sheets: HandSheet[];
  strings?: string;
  styles?: string;
  workbookProperties?: string;
}

/** A package written by hand, for markup no spreadsheet library writes. */
async function handWorkbook(workbook: HandWorkbook): Promise<Uint8Array> {
  const zip = new JSZip();
  const overrides = [
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`,
    ...workbook.sheets.map(
      (_, index) =>
        `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    ),
    workbook.strings === undefined
      ? ""
      : `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`,
    workbook.styles === undefined
      ? ""
      : `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`,
  ];
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.join("")}</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0"?><Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0"?><workbook xmlns="${MAIN}" xmlns:r="${REL}">${workbook.workbookProperties ?? ""}<sheets>${workbook.sheets
      .map(
        (sheet, index) =>
          `<sheet name="${sheet.name}" sheetId="${index + 1}" r:id="rId${index + 1}"${sheet.state ? ` state="${sheet.state}"` : ""}/>`,
      )
      .join("")}</sheets></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0"?><Relationships xmlns="${PKG_REL}">${workbook.sheets
      .map(
        (_, index) =>
          `<Relationship Id="rId${index + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`,
      )
      .join(
        "",
      )}${workbook.strings === undefined ? "" : `<Relationship Id="rS" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`}${workbook.styles === undefined ? "" : `<Relationship Id="rT" Type="${REL}/styles" Target="styles.xml"/>`}</Relationships>`,
  );
  workbook.sheets.forEach((sheet, index) => {
    zip.file(
      `xl/worksheets/sheet${index + 1}.xml`,
      `<?xml version="1.0"?><worksheet xmlns="${MAIN}">${sheet.before ?? ""}<sheetData>${sheet.data}</sheetData>${sheet.after ?? ""}</worksheet>`,
    );
  });
  if (workbook.strings !== undefined) {
    zip.file(
      "xl/sharedStrings.xml",
      `<?xml version="1.0"?><sst xmlns="${MAIN}">${workbook.strings}</sst>`,
    );
  }
  if (workbook.styles !== undefined) {
    zip.file(
      "xl/styles.xml",
      `<?xml version="1.0"?><styleSheet xmlns="${MAIN}">${workbook.styles}</styleSheet>`,
    );
  }
  return zip.generateAsync({ type: "uint8array" });
}

/** Styles where style 1 is a date format and style 2 a custom date. */
const DATE_STYLES = `<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/></numFmts><cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="164"/></cellXfs>`;

interface SheetGrid {
  name: string;
  range: string | undefined;
  cells: Map<string, StreamedValue>;
}

function rangeText(range: {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}): string {
  return XLSX.utils.encode_range({
    s: { r: range.startRow, c: range.startColumn },
    e: { r: range.endRow, c: range.endColumn },
  });
}

async function streamedGrids(bytes: Uint8Array): Promise<SheetGrid[]> {
  const workbook = await StreamedWorkbook.open(
    bytesSource("book.xlsx", bytes),
    {
      source: "book.xlsx",
      file: "book.xlsx",
      details: { source: "book.xlsx" },
    },
  );
  const grids: SheetGrid[] = [];
  for (const sheet of workbook.sheets) {
    const cells = new Map<string, StreamedValue>();
    const read = await workbook.readWorksheet(sheet, {
      begin: () => {
        cells.clear();
      },
      row: (row, rowCells: readonly StreamedCell[]) => {
        for (const cell of rowCells) {
          cells.set(`${row},${cell.column}`, cell.value);
        }
      },
    });
    grids.push({
      name: sheet.name,
      range: read.range === undefined ? undefined : rangeText(read.range),
      cells,
    });
  }
  return grids;
}

const ERROR_TEXT: Record<number, string> = {
  0: "#NULL!",
  7: "#DIV/0!",
  15: "#VALUE!",
  23: "#REF!",
  29: "#NAME?",
  36: "#NUM!",
  42: "#N/A",
  43: "#GETTING_DATA",
  255: "#WTF?",
};

/** The SheetJS-backed reader's answer, with error cells spelled as text. */
async function engineGrids(bytes: Uint8Array): Promise<SheetGrid[]> {
  const workbook = parseWorkbookBytes(bytes, "book.xlsx");
  const dates = await readWorkbookDates(bytes, "book.xlsx", {});
  return workbook.SheetNames.map((name) => {
    const sheet = workbook.Sheets[name]!;
    const reference = sheet["!ref"];
    const cells = new Map<string, StreamedValue>();
    if (reference !== undefined) {
      const range = XLSX.utils.decode_range(reference);
      const sheetDates = dates.forSheet(name);
      for (let row = range.s.r; row <= range.e.r; row += 1) {
        for (let column = range.s.c; column <= range.e.c; column += 1) {
          const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })] as
            XLSX.CellObject | undefined;
          const value = cellToPrimitive(cell, sheetDates(row, column));
          if (value === null) continue;
          cells.set(
            `${row},${column}`,
            cell?.t === "e" && typeof value === "number"
              ? new CellError(ERROR_TEXT[value] ?? String(value))
              : value,
          );
        }
      }
    }
    return { name, range: reference, cells };
  });
}

function plain(grids: SheetGrid[]): unknown {
  return grids.map((grid) => ({
    name: grid.name,
    range: grid.range,
    cells: [...grid.cells].map(([key, value]) => [
      key,
      value instanceof CellError ? { error: value.text } : value,
    ]),
  }));
}

async function expectSameAsEngine(bytes: Uint8Array): Promise<void> {
  expect(plain(await streamedGrids(bytes))).toEqual(
    plain(await engineGrids(bytes)),
  );
}

function sheetJsWorkbook(
  rows: unknown[][],
  options: { bookSST: boolean; cellDates: boolean },
): Uint8Array {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet(rows, { cellDates: options.cellDates }),
    "Data",
  );
  return new Uint8Array(
    XLSX.write(workbook, {
      type: "array",
      bookType: "xlsx",
      bookSST: options.bookSST,
      cellDates: options.cellDates,
    }) as ArrayBuffer,
  );
}

describe("streamed reader: values match the engine's", () => {
  const mixedRows = [
    ["Name", "Amount", "Flag", "When", "Note"],
    ["Ada", 1.5, true, new Date(Date.UTC(2024, 0, 31)), "a & b < c"],
    ["Bo", -0.000001, false, null, "  padded  "],
    [
      null,
      1e21,
      null,
      new Date(Date.UTC(1999, 11, 31, 12, 30)),
      "emoji \u{1F600}",
    ],
    ["", 0, true, null, "x"],
  ];

  it.each([
    ["inline strings", { bookSST: false, cellDates: false }],
    ["shared strings", { bookSST: true, cellDates: false }],
    ["dates as dates", { bookSST: true, cellDates: true }],
  ])("reads a SheetJS workbook with %s", async (_, options) => {
    await expectSameAsEngine(sheetJsWorkbook(mixedRows, options));
  });

  it.each([
    ["a plain range", { shape: "range" as const }],
    [
      "a table with formulas, totals and dependents",
      {
        shape: "table" as const,
        formulas: "a1" as const,
        totalsRow: true,
        dependents: true,
        summarySheet: true,
        footerBlock: true,
        implicitRow: true,
        hiddenSheets: true,
        uncachedFormula: true,
        sharedFormula: true,
        arrayFormula: true,
      },
    ],
  ])("reads the corpus workbook as %s", async (_, options) => {
    await expectSameAsEngine(await buildCorpusWorkbook(options));
  });

  it("reads hand-written markup the way the engine does", async () => {
    await expectSameAsEngine(
      await handWorkbook({
        styles: DATE_STYLES,
        strings: `<si><t>plain</t></si><si><r><t>ri</t></r><r><rPr><b/></rPr><t>ch</t></r><rPh><t>skip</t></rPh></si><si><t xml:space="preserve"> sp </t></si><si><t>a</t><rPh><t>b</t></rPh></si><si></si><si> </si>`,
        sheets: [
          {
            name: "NoDimension",
            data:
              `<row r="2"><c r="B2" t="s"><v>0</v></c><c r="C2" t="s"><v>1</v></c><c r="D2" t="s"><v>2</v></c><c r="E2" t="s"><v>3</v></c><c r="F2" t="s"><v>4</v></c></row>` +
              `<row r="3"><c r="B3" s="1"><v>45322.5</v></c><c r="C3" s="2"><v>60</v></c><c r="D3" t="d"><v>2024-02-29T10:00:00Z</v></c><c r="E3" t="d"><v>not a date</v></c><c r="F3" t="d"><v>  </v></c></row>` +
              `<row r="4"><c r="B4" t="b"/><c r="C4" t="b"><v>true</v></c><c r="D4" t="e"><v>#N/A</v></c><c r="E4" t="e"><v>#BOGUS</v></c><c r="F4" t="str"/></row>` +
              `<row r="5"><c r="B5"><f>1+1</f></c><c r="C5" t="n"/><c r="D5" s="1"/><c r="H5" t="s"/><c r="E5" t="inlineStr"><is><r><t>in</t></r><r><t>line</t></r></is></c><c r="F5" t="inlineStr"/></row>` +
              `<row r="6"/><row r="7"><c><v>1</v></c><c><v>2</v></c><c r="F7"><v>12abc</v></c><c><v>3</v></c></row>` +
              `<row r="8" ht="20"></row><row r="12"/><row r="14"/>`,
          },
          {
            name: "Formulas",
            data: `<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>1+1</f></c></row><row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">B1</f><v>2</v></c><c r="D2"><f t="shared" si="0"/></c><c r="C2"><f t="array" ref="C2:E3">X</f><v>3</v></c></row><row r="3"><c r="E3" s="0"/><c r="F3"><f t="shared" si="9"/></c></row>`,
          },
          {
            name: "EmptyItems",
            data: `<row r="1"><c r="A1" t="s"><v>5</v></c><c r="B1" t="s"><v>6</v></c><c r="C1" t="inlineStr"><is></is></c><c r="D1" t="inlineStr"><is> </is></c><c r="E1"><v>1</v></c></row>`,
          },
          {
            name: "Dimension",
            before: `<dimension ref="B2:C3"/>`,
            data: `<row r="1"><c r="A1"><v>9</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="D2"><v>3</v></c></row><row r="3"><c r="C3"><v>4</v></c></row><row r="4"><c r="C4"><v>5</v></c></row>`,
          },
          {
            name: "Shuffled",
            data: `<row r="3"><c r="A3"><v>3</v></c></row><row r="1"><c r="B1"><v>1</v></c><c r="A1"><v>0</v></c></row><row r="2"><c r="A5"><v>5</v></c></row><row r="2"><c r="B2"><v>2</v></c></row>`,
          },
          {
            name: "SingleCellDimension",
            before: `<dimension ref="A1"/>`,
            data: `<row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>3</v></c></row>`,
          },
          {
            name: "Merged",
            data: `<row r="1"><c r="A1" t="str"><v>Title</v></c></row><row r="3"><c r="A3"><v>1</v></c></row>`,
            after: `<mergeCells count="2"><mergeCell ref="A1:C1"/><mergeCell ref="A3:A4"/></mergeCells>`,
          },
          {
            name: "Hidden",
            state: "hidden",
            data: `<row r="1"><c r="A1"><v>1</v></c></row>`,
          },
        ],
      }),
    );
  });

  it("reads the 1904 date system by the workbook's own count", async () => {
    await expectSameAsEngine(
      await handWorkbook({
        styles: DATE_STYLES,
        workbookProperties: `<workbookPr date1904="1"/>`,
        sheets: [
          {
            name: "Dates",
            data: `<row r="1"><c r="A1" s="1"><v>0</v></c><c r="B1" s="1"><v>1462.25</v></c><c r="C1" s="2"><v>-1</v></c></row>`,
          },
        ],
      }),
    );
  });
});

describe("streamed reader: the package's own streaming writer", () => {
  it("reads back what the table writer wrote, data descriptors and all", async () => {
    const chunks: Uint8Array[] = [];
    const writer = new TableWorkbookWriter({
      sheetName: "Out",
      columns: ["Name", "Amount"],
      widths: [10, 10],
      rowCount: 2,
      onChunk: (chunk) => chunks.push(chunk),
    });
    writer.writeRow(["a\r\nb", 1.5]);
    writer.writeRow([null, true]);
    writer.finish();
    const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
    let offset = 0;
    for (const chunk of chunks) {
      bytes.set(chunk, offset);
      offset += chunk.length;
    }
    const [grid] = await streamedGrids(bytes);
    expect(grid!.range).toBe("A1:B3");
    expect([...grid!.cells]).toEqual([
      ["0,0", "Name"],
      ["0,1", "Amount"],
      ["1,0", "a\r\nb"],
      ["1,1", 1.5],
      ["2,1", true],
    ]);
  });
});

describe("streamed reader: where it deliberately differs", () => {
  it("decodes escapes once and keeps a carriage return before a line feed", async () => {
    const grids = await streamedGrids(
      await handWorkbook({
        strings: `<si><t>one_x000D_\ntwo</t></si><si><t>_x005F_x0041_ and _x0041_</t></si>`,
        sheets: [
          {
            name: "Text",
            data: `<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="str"><v>&amp;lt;b&amp;gt;</v></c><c r="D1" t="inlineStr"><is><t>a_x000D_\nb</t></is></c></row>`,
          },
        ],
      }),
    );
    expect([...grids[0]!.cells.values()]).toEqual([
      "one\r\ntwo",
      "_x0041_ and A",
      "&lt;b&gt;",
      "a\r\nb",
    ]);
  });

  it("reads an error cell as its text, not the engine's code", async () => {
    const grids = await streamedGrids(
      await handWorkbook({
        sheets: [
          {
            name: "Errors",
            data: `<row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c><c r="B1" t="str"><v>#DIV/0!</v></c></row>`,
          },
        ],
      }),
    );
    const [error, text] = [...grids[0]!.cells.values()];
    expect(error).toBeInstanceOf(CellError);
    expect(String(error)).toBe("#DIV/0!");
    expect(text).toBe("#DIV/0!");
  });
});

describe("streamed reader: what it refuses", () => {
  const open = (bytes: Uint8Array) =>
    StreamedWorkbook.open(bytesSource("book.xlsx", bytes), {
      source: "book.xlsx",
      file: "book.xlsx",
      details: { source: "book.xlsx" },
    });

  it("refuses bytes that are not a workbook package", async () => {
    await expect(
      open(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 9, 9, 9, 9])),
    ).rejects.toMatchObject({
      code: "XLSX_READ_FAILED",
      message: "Could not read workbook: book.xlsx",
      details: { source: "book.xlsx" },
    });
  });

  it("refuses a password-protected workbook, which is no zip package", async () => {
    const container = XLSX.CFB.utils.cfb_new();
    XLSX.CFB.utils.cfb_add(container, "/EncryptionInfo", new Uint8Array(8));
    XLSX.CFB.utils.cfb_add(container, "/EncryptedPackage", new Uint8Array(64));
    await expect(
      open(
        new Uint8Array(
          XLSX.CFB.write(container, { type: "array" }) as number[],
        ),
      ),
    ).rejects.toMatchObject({ code: "XLSX_READ_FAILED" });
  });

  it("refuses a package with no workbook part", async () => {
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
    );
    await expect(
      open(await zip.generateAsync({ type: "uint8array" })),
    ).rejects.toMatchObject({ code: "XLSX_READ_FAILED" });
  });

  it("refuses a worksheet whose shared string the table does not hold", async () => {
    const bytes = await handWorkbook({
      strings: `<si><t>only</t></si>`,
      sheets: [
        {
          name: "Broken",
          data: `<row r="1"><c r="A1" t="s"><v>7</v></c></row>`,
        },
      ],
    });
    const workbook = await open(bytes);
    await expect(
      workbook.readWorksheet(workbook.sheets[0]!, {
        begin() {},
        row() {},
      }),
    ).rejects.toMatchObject({
      code: "XLSX_READ_FAILED",
      details: { source: "book.xlsx", worksheet: "Broken" },
    });
  });

  it("refuses a worksheet whose markup is not well formed, naming it", async () => {
    const bytes = await handWorkbook({
      sheets: [{ name: "Bad", data: `<row r="1"><c r="A1"><v>1</c></row>` }],
    });
    const workbook = await open(bytes);
    await expect(
      workbook.readWorksheet(workbook.sheets[0]!, {
        begin() {},
        row() {},
      }),
    ).rejects.toMatchObject({
      code: "XLSX_READ_FAILED",
      message: 'Could not read worksheet "Bad" in workbook: book.xlsx',
    });
  });
});
