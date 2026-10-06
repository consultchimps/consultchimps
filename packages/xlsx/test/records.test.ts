import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { readWorksheetRecordsBytes } from "../src/bytes.js";
import { readWorksheetRecords } from "../src/index.js";
import { buildSheetFixture } from "./support/workbook-fixture.js";

describe("readWorksheetRecords", () => {
  it("returns deterministic displayed text and skips only completely empty rows", async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "consultchimps-xlsx-records-"),
    );

    try {
      const workbookPath = path.join(directory, "records.xlsx");
      await writeFile(
        workbookPath,
        await buildSheetFixture("Companies", [
          ["name", "percentage", "date", "active", "empty", "formula"],
          [
            "Company A",
            { value: 0.125, format: "0.0%" },
            // 2024-01-02 as a serial.
            { value: 45293, format: "yyyy-mm-dd" },
            true,
            null,
            { formula: "B2*2", value: 0.25, format: "0.0%" },
          ],
          [null, null, null, null, null, null],
          [
            "Company B",
            { value: -0.021, format: "0.0%" },
            null,
            false,
            null,
            null,
          ],
        ]),
      );

      await expect(
        readWorksheetRecords(workbookPath, {
          headerRow: 1,
          worksheet: "companies",
        }),
      ).resolves.toEqual({
        columns: ["name", "percentage", "date", "active", "empty", "formula"],
        rows: [
          {
            active: "TRUE",
            date: "2024-01-02",
            empty: "",
            formula: "25.0%",
            name: "Company A",
            percentage: "12.5%",
          },
          {
            active: "FALSE",
            date: "",
            empty: "",
            formula: "",
            name: "Company B",
            percentage: "-2.1%",
          },
        ],
        skippedEmptyRows: 1,
        sourceRows: [2, 4],
        uncachedFormulas: [],
        worksheet: "Companies",
      });

      await expect(
        readWorksheetRecords(workbookPath, { headerRow: 1 }),
      ).resolves.toMatchObject({ worksheet: "Companies" });
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });
});

describe("readWorksheetRecordsBytes display text", () => {
  const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const REL =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  // Each style applies one number format, in cellXfs order from s="1".
  const FORMATS = [
    "m/d/yy",
    "0.0%",
    '"$"#,##0.00',
    "[h]:mm",
    '"Ref "@',
    "[$€-407]#,##0.00",
    "#,##0",
  ];

  async function formatted(
    date1904: boolean,
    shuffled = false,
    stylesPart = "styles.xml",
  ): Promise<Uint8Array> {
    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/${stylesPart}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    );
    zip.file(
      "_rels/.rels",
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    );
    zip.file(
      "xl/workbook.xml",
      `<?xml version="1.0"?><workbook xmlns="${MAIN}" xmlns:r="${REL}">${date1904 ? `<workbookPr date1904="1"/>` : ""}<sheets><sheet name="Values" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    );
    zip.file(
      "xl/_rels/workbook.xml.rels",
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="${stylesPart}"/></Relationships>`,
    );
    const custom = FORMATS.map(
      (code, index) =>
        `<numFmt numFmtId="${164 + index}" formatCode="${code.replaceAll("&", "&amp;").replaceAll('"', "&quot;")}"/>`,
    ).join("");
    const styles = FORMATS.map(
      (_, index) => `<xf numFmtId="${164 + index}" applyNumberFormat="1"/>`,
    ).join("");
    zip.file(
      `xl/${stylesPart}`,
      `<?xml version="1.0"?><styleSheet xmlns="${MAIN}"><numFmts count="${FORMATS.length}">${custom}</numFmts><cellXfs count="${FORMATS.length + 1}"><xf numFmtId="0"/>${styles}</cellXfs></styleSheet>`,
    );
    const headers = [
      "Date",
      "Share",
      "Amount",
      "Elapsed",
      "Reference",
      "Euro",
      "Loss",
      "Flag",
      "Error",
      "Plain",
    ];
    const column = (index: number): string => String.fromCharCode(65 + index);
    const header = headers
      .map(
        (name, index) =>
          `<c r="${column(index)}1" t="inlineStr"><is><t>${name}</t></is></c>`,
      )
      .join("");
    const serial = date1904 ? 45292 - 1462 : 45292;
    const values =
      `<c r="A2" s="1"><v>${serial}</v></c><c r="B2" s="2"><v>0.125</v></c>` +
      `<c r="C2" s="3"><v>1234.5</v></c><c r="D2" s="4"><v>1.5</v></c>` +
      `<c r="E2" s="5" t="inlineStr"><is><t>A-7</t></is></c>` +
      `<c r="F2" s="6"><v>1234.5</v></c><c r="G2" s="7"><v>-1234.4</v></c>` +
      `<c r="H2" t="b"><v>1</v></c><c r="I2" t="e"><v>#N/A</v></c>` +
      `<c r="J2"><v>0.30000000000000004</v></c>`;
    zip.file(
      "xl/worksheets/sheet1.xml",
      `<?xml version="1.0"?><worksheet xmlns="${MAIN}"><sheetData>${shuffled ? `<row r="2">${values}</row><row r="1">${header}</row>` : `<row r="1">${header}</row><row r="2">${values}</row>`}</sheetData></worksheet>`,
    );
    return new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
  }

  it.each([
    [false, false, "styles.xml"],
    [true, false, "styles.xml"],
    // Rows out of order are gathered before they are delivered.
    [false, true, "styles.xml"],
    // The styles part lives wherever the relationship says.
    [false, false, "theme/formats.xml"],
  ])(
    "applies each cell's number format (1904: %s, shuffled rows: %s, styles: %s)",
    async (date1904, shuffled, stylesPart) => {
      const records = await readWorksheetRecordsBytes({
        name: "values.xlsx",
        bytes: await formatted(date1904, shuffled, stylesPart),
      });

      expect(records.rows).toEqual([
        {
          Date: "1/1/24",
          Share: "12.5%",
          Amount: "$1,234.50",
          // A duration is the same in both date systems.
          Elapsed: "36:00",
          Reference: "Ref A-7",
          // The locale tag keeps its symbol and the reader's separators.
          Euro: "€1,234.50",
          Loss: "-1,234",
          Flag: "TRUE",
          Error: "#N/A",
          Plain: "0.3",
        },
      ]);
    },
  );
});
