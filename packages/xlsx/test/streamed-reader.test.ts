/**
 * The streaming reader consolidation reads through (ADR 0006): what it makes
 * of each cell. The expected grids were recorded from the SheetJS-backed reader
 * it replaced, whose answers it gives except where it deliberately differs.
 */
import { readFile } from "node:fs/promises";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { encodeRange } from "../src/model/references.js";
import {
  StreamedWorkbook,
  type StreamedCell,
  type StreamedValue,
} from "../src/operations/consolidate/reader.js";
import { CellError } from "../src/package/cell-error.js";
import { bytesSource } from "../src/package/index.js";
import { TableWorkbookWriter } from "../src/package/table-writer.js";
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
      range: read.range === undefined ? undefined : encodeRange(read.range),
      cells,
    });
  }
  return grids;
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

type Expected = {
  name: string;
  range: string | undefined;
  cells: [string, string | number | boolean | { error: string }][];
}[];

async function expectGrids(
  bytes: Uint8Array,
  expected: Expected,
): Promise<void> {
  expect(plain(await streamedGrids(bytes))).toEqual(expected);
}

const MIXED_GRID: Expected = [
  {
    name: "Data",
    range: "A1:E5",
    cells: [
      ["0,0", "Name"],
      ["0,1", "Amount"],
      ["0,2", "Flag"],
      ["0,3", "When"],
      ["0,4", "Note"],
      ["1,0", "Ada"],
      ["1,1", 1.5],
      ["1,2", true],
      ["1,3", "2024-01-31T00:00:00.000Z"],
      ["1,4", "a & b < c"],
      ["2,0", "Bo"],
      ["2,1", -0.000001],
      ["2,2", false],
      ["2,4", "  padded  "],
      ["3,1", 1e21],
      ["3,3", "1999-12-31T12:30:00.000Z"],
      ["3,4", "emoji 😀"],
      ["4,0", ""],
      ["4,1", 0],
      ["4,2", true],
      ["4,4", "x"],
    ],
  },
];

/** A date cell: its serial, and the ISO text a declared date cell holds. */
interface MixedDate {
  serial: number;
  iso: string;
}

const MIXED_ROWS: (string | number | boolean | MixedDate | null)[][] = [
  ["Name", "Amount", "Flag", "When", "Note"],
  [
    "Ada",
    1.5,
    true,
    { serial: 45322, iso: "2024-01-31T00:00:00.000Z" },
    "a & b < c",
  ],
  ["Bo", -0.000001, false, null, "  padded  "],
  [
    null,
    1e21,
    null,
    { serial: 36525.520833333336, iso: "1999-12-31T12:30:00.000Z" },
    "emoji 😀",
  ],
  ["", 0, true, null, "x"],
];

function escapeText(text: string): string {
  return text.replace(/&/gu, "&amp;").replace(/</gu, "&lt;");
}

/**
 * The mixed rows as a spreadsheet library writes them: text as formula
 * strings or shared strings, dates as styled serials or declared ISO text.
 */
function mixedWorkbook(options: {
  sharedStrings: boolean;
  declaredDates: boolean;
}): Promise<Uint8Array> {
  const strings: string[] = [];
  const data = MIXED_ROWS.map((row, rowIndex) => {
    const cells = row.map((value, column) => {
      const ref = `${"ABCDE"[column]}${rowIndex + 1}`;
      if (value === null) return "";
      if (typeof value === "string") {
        if (options.sharedStrings) {
          strings.push(value);
          return `<c r="${ref}" t="s"><v>${strings.length - 1}</v></c>`;
        }
        const space = value.trim() === value ? "" : ' xml:space="preserve"';
        return `<c r="${ref}" t="str"><v${space}>${escapeText(value)}</v></c>`;
      }
      if (typeof value === "boolean") {
        return `<c r="${ref}" t="b"><v>${value ? 1 : 0}</v></c>`;
      }
      if (typeof value === "number") {
        return `<c r="${ref}"><v>${value}</v></c>`;
      }
      return options.declaredDates
        ? `<c r="${ref}" s="1" t="d"><v>${value.iso}</v></c>`
        : `<c r="${ref}" s="1"><v>${value.serial}</v></c>`;
    });
    return `<row r="${rowIndex + 1}">${cells.join("")}</row>`;
  });
  return handWorkbook({
    styles: `<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs>`,
    ...(options.sharedStrings
      ? {
          strings: strings
            .map((text) =>
              text.trim() === text
                ? `<si><t>${escapeText(text)}</t></si>`
                : `<si><t xml:space="preserve">${escapeText(text)}</t></si>`,
            )
            .join(""),
        }
      : {}),
    sheets: [
      { name: "Data", data: data.join(""), before: `<dimension ref="A1:E5"/>` },
    ],
  });
}

const PLAIN_RANGE_GRID: Expected = [
  {
    name: "Data",
    range: "A1:I12",
    cells: [
      ["0,0", "Corpus allocation report"],
      ["2,0", "Record"],
      ["2,1", "Client"],
      ["2,2", "Group"],
      ["2,3", "Amount"],
      ["2,4", "Doubled"],
      ["2,5", "Ratio"],
      ["3,0", 1],
      ["3,1", "Client A"],
      ["3,2", "Alpha"],
      ["3,3", 10],
      ["3,4", 20],
      ["3,5", 5],
      ["4,0", 2],
      ["4,1", "Client B"],
      ["4,2", "Beta"],
      ["4,3", 20],
      ["4,4", 40],
      ["4,5", 10],
      ["5,0", 3],
      ["5,1", "Client C"],
      ["5,2", "Alpha"],
      ["5,3", 30],
      ["5,4", 60],
      ["5,5", 15],
      ["5,7", "Alpha side note"],
      ["6,0", 4],
      ["6,1", "Client D"],
      ["6,2", "Beta"],
      ["6,3", 40],
      ["6,4", 80],
      ["6,5", 20],
      ["7,0", 5],
      ["7,1", "Client E"],
      ["7,2", "Gamma"],
      ["7,3", 50],
      ["7,4", 100],
      ["7,5", 25],
      ["8,0", 6],
      ["8,1", "Client F"],
      ["8,2", "Alpha"],
      ["8,3", 60],
      ["8,4", 120],
      ["8,5", 30],
      ["11,0", "Footer note"],
      ["11,1", 210],
    ],
  },
  {
    name: "Summary",
    range: "A1:B3",
    cells: [
      ["0,0", "Portfolio summary"],
      ["1,0", "Total across every record"],
      ["1,1", 210],
      ["2,0", "First amount"],
      ["2,1", 10],
    ],
  },
  {
    name: "Hidden",
    range: "A1:B4",
    cells: [
      ["0,0", "Record"],
      ["0,1", "Group"],
      ["1,0", 7],
      ["1,1", "Alpha"],
      ["2,0", 8],
      ["2,1", "Beta"],
      ["3,0", 9],
      ["3,1", "Alpha"],
    ],
  },
  {
    name: "VeryHidden",
    range: "A1:A2",
    cells: [
      ["0,0", "Archive note"],
      ["1,0", "Retained without filtering"],
    ],
  },
];

const TABLE_GRID: Expected = [
  {
    name: "Data",
    range: "A1:I12",
    cells: [
      ["0,0", "Corpus allocation report"],
      ["2,0", "Record"],
      ["2,1", "Client"],
      ["2,2", "Group"],
      ["2,3", "Amount"],
      ["2,4", "Doubled"],
      ["2,5", "Ratio"],
      ["3,0", 1],
      ["3,1", "Client A"],
      ["3,2", "Alpha"],
      ["3,3", 10],
      ["3,4", 20],
      ["3,5", 5],
      ["3,6", 210],
      ["4,0", 2],
      ["4,1", "Client B"],
      ["4,2", "Beta"],
      ["4,3", 20],
      ["4,4", 40],
      ["4,5", 10],
      ["5,0", 3],
      ["5,1", "Client C"],
      ["5,2", "Alpha"],
      ["5,3", 30],
      ["5,4", 60],
      ["5,5", 15],
      ["5,7", "Alpha side note"],
      ["6,0", 4],
      ["6,1", "Client D"],
      ["6,2", "Beta"],
      ["6,3", 40],
      ["6,4", 80],
      ["6,5", 20],
      ["7,0", 5],
      ["7,1", "Client E"],
      ["7,2", "Gamma"],
      ["7,3", 50],
      ["7,4", 100],
      ["7,5", 25],
      ["8,0", 6],
      ["8,1", "Client F"],
      ["8,2", "Alpha"],
      ["8,3", 60],
      ["8,4", 120],
      ["8,5", 30],
      ["9,0", "Total"],
      ["9,3", 210],
      ["10,0", "Implicit note"],
      ["10,1", 1],
      ["11,0", "Footer note"],
      ["11,1", 210],
    ],
  },
  {
    name: "Summary",
    range: "A1:B4",
    cells: [
      ["0,0", "Portfolio summary"],
      ["1,0", "Total across every record"],
      ["1,1", 210],
      ["2,0", "First amount"],
      ["2,1", 10],
      ["3,0", "Never recalculated"],
    ],
  },
  {
    name: "Hidden",
    range: "A1:B4",
    cells: [
      ["0,0", "Record"],
      ["0,1", "Group"],
      ["1,0", 7],
      ["1,1", "Alpha"],
      ["2,0", 8],
      ["2,1", "Beta"],
      ["3,0", 9],
      ["3,1", "Alpha"],
    ],
  },
  {
    name: "VeryHidden",
    range: "A1:A2",
    cells: [
      ["0,0", "Archive note"],
      ["1,0", "Retained without filtering"],
    ],
  },
];

describe("streamed reader: values match the engine's", () => {
  it.each([
    ["inline strings", { sharedStrings: false, declaredDates: false }],
    ["shared strings", { sharedStrings: true, declaredDates: false }],
    ["dates as dates", { sharedStrings: true, declaredDates: true }],
  ])("reads a workbook with %s", async (_, options) => {
    await expectGrids(await mixedWorkbook(options), MIXED_GRID);
  });

  it.each([
    ["a plain range", { shape: "range" as const }, PLAIN_RANGE_GRID],
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
      TABLE_GRID,
    ],
  ])("reads the corpus workbook as %s", async (_, options, expected) => {
    await expectGrids(await buildCorpusWorkbook(options), expected);
  });

  it("reads hand-written markup the way the engine does", async () => {
    await expectGrids(
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
      [
        {
          name: "NoDimension",
          range: "A2:H14",
          cells: [
            ["1,1", "plain"],
            ["1,2", "rich"],
            ["1,3", " sp "],
            ["1,4", "a"],
            ["1,5", ""],
            ["2,1", "2024-01-31T12:00:00.000Z"],
            ["2,2", 60],
            ["2,3", "2024-02-29T10:00:00.000Z"],
            ["2,4", "not a date"],
            ["3,1", false],
            ["3,2", true],
            ["3,3", { error: "#N/A" }],
            ["3,5", ""],
            ["4,4", "inline"],
            ["4,5", ""],
            ["6,0", 1],
            ["6,1", 2],
            ["6,5", 12],
            ["6,6", 3],
          ],
        },
        {
          name: "Formulas",
          range: "A1:E3",
          cells: [
            ["0,0", 1],
            ["1,0", 2],
            ["1,2", 3],
          ],
        },
        {
          name: "EmptyItems",
          range: "A1:E1",
          cells: [
            ["0,0", ""],
            ["0,1", ""],
            ["0,2", ""],
            ["0,4", 1],
          ],
        },
        {
          name: "Dimension",
          range: "B2:C3",
          cells: [
            ["1,1", 2],
            ["2,2", 4],
          ],
        },
        {
          name: "Shuffled",
          range: "A1:B3",
          cells: [
            ["0,0", 0],
            ["0,1", 1],
            ["1,1", 2],
            ["2,0", 3],
          ],
        },
        {
          name: "SingleCellDimension",
          range: "A1:C1",
          cells: [
            ["0,0", 1],
            ["0,2", 3],
          ],
        },
        {
          name: "Merged",
          range: "A1:A3",
          cells: [
            ["0,0", "Title"],
            ["2,0", 1],
          ],
        },
        { name: "Hidden", range: "A1", cells: [["0,0", 1]] },
      ],
    );
  });

  it("reads the 1904 date system by the workbook's own count", async () => {
    await expectGrids(
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
      [
        {
          name: "Dates",
          range: "A1:C1",
          cells: [
            ["0,0", "1904-01-01T00:00:00.000Z"],
            ["0,1", "1908-01-02T06:00:00.000Z"],
            ["0,2", -1],
          ],
        },
      ],
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

/** A password-protected placeholder: a compound file, no zip package. */
async function encryptedPlaceholder(): Promise<Uint8Array> {
  return readFile(
    new URL("./fixtures/encrypted-placeholder.xlsx", import.meta.url),
  );
}

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
    await expect(open(await encryptedPlaceholder())).rejects.toMatchObject({
      code: "XLSX_READ_FAILED",
    });
  });

  it("refuses an entry whose bytes fail its CRC, though its length holds", async () => {
    const bytes = await handWorkbook({
      sheets: [
        { name: "Data", data: `<row r="1"><c r="A1"><v>1234</v></c></row>` },
      ],
    });
    const at = Buffer.from(bytes).indexOf("<v>1234</v>");
    bytes[at + 3] = "9".charCodeAt(0);
    const workbook = await open(bytes);
    await expect(
      workbook.readWorksheet(workbook.sheets[0]!, { begin() {}, row() {} }),
    ).rejects.toMatchObject({
      code: "XLSX_READ_FAILED",
      message: 'Could not read worksheet "Data" in workbook: book.xlsx',
    });
  });

  it("refuses a structural part that declares an implausible size before allocating it", async () => {
    const bytes = await handWorkbook({
      sheets: [{ name: "Data", data: "" }],
    });
    const name = Buffer.from("[Content_Types].xml");
    const header = Buffer.from(bytes).lastIndexOf(name) - 46;
    new DataView(bytes.buffer, bytes.byteOffset).setUint32(
      header + 24,
      0x7fffffff,
      true,
    );
    await expect(open(bytes)).rejects.toMatchObject({
      code: "XLSX_READ_FAILED",
      message: "Could not read workbook: book.xlsx",
    });
  });

  it("finds the directory record behind a zip comment that imitates one", async () => {
    const bytes = await handWorkbook({
      sheets: [
        { name: "Data", data: `<row r="1"><c r="A1"><v>1</v></c></row>` },
      ],
    });
    const comment = new Uint8Array(30);
    comment.set([0x50, 0x4b, 0x05, 0x06], 2);
    const withComment = new Uint8Array(bytes.length + comment.length);
    withComment.set(bytes);
    withComment.set(comment, bytes.length);
    new DataView(withComment.buffer).setUint16(
      bytes.length - 2,
      comment.length,
      true,
    );
    const grids = await streamedGrids(withComment);
    expect([...grids[0]!.cells.values()]).toEqual([1]);
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
