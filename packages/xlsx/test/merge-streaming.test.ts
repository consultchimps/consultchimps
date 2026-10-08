/**
 * Merging reads each input in pieces and writes the merged workbook as it is
 * produced (ADR 0006), and gives the bytes the byte surface gives.
 */
import type { RandomAccessSource } from "@consultchimps/core";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import {
  blobSource,
  mergeWorkbookSources,
  mergeWorkbooksBytes,
} from "../src/bytes.js";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MIB = 1024 * 1024;

function cell(ref: string, text: string, style = 0): string {
  return `<c r="${ref}" s="${style}" t="inlineStr"><is><t>${text}</t></is></c>`;
}

/**
 * A stored workbook over 1 MiB: rows using a few cell styles, a shared string,
 * a formula, a conditional format, and a binary part no step reads.
 */
async function workbook(label: string, rows: number): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId3" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
  );
  zip.file(
    "xl/styles.xml",
    `<styleSheet xmlns="${MAIN}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>`,
  );
  zip.file(
    "xl/sharedStrings.xml",
    `<sst xmlns="${MAIN}" count="1" uniqueCount="1"><si><t>${label} shared</t></si></sst>`,
  );
  const body = [
    `<row r="1" s="1" customFormat="1">${cell("A1", "Case", 1)}${cell("B1", "Note", 1)}<c r="C1" t="s"><v>0</v></c></row>`,
  ];
  for (let row = 2; row <= rows; row += 1) {
    body.push(
      `<row r="${row}">${cell(`A${row}`, `${label}-${row}`, row % 3)}${cell(`B${row}`, `A note for row ${row}`)}<c r="C${row}" s="2"><f>ROW()</f><v>${row}</v></c></row>`,
    );
  }
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<worksheet xmlns="${MAIN}" xmlns:r="${REL}"><dimension ref="A1:C${rows}"/><cols><col min="1" max="1" width="12" style="1"/></cols><sheetData>${body.join("")}</sheetData><conditionalFormatting sqref="A2:A${rows}"><cfRule type="containsBlanks" dxfId="0" priority="1"/></conditionalFormatting></worksheet>`,
  );
  zip.file(
    "xl/embeddings/blob.bin",
    new Uint8Array(512 * 1024).map((_, index) => (index * 31) % 251),
  );
  return zip.generateAsync({ type: "uint8array", compression: "STORE" });
}

/** A source that refuses any read over 1 MiB, and so any whole read. */
function guardedSource(name: string, bytes: Uint8Array): RandomAccessSource {
  const inner = blobSource(name, new Blob([bytes.slice()]));
  return {
    name,
    size: inner.size,
    readAt: (offset, length, signal) => {
      if (length > MIB) {
        throw new Error(`Read ${length} bytes of ${inner.size} at once`);
      }
      return inner.readAt(offset, length, signal);
    },
  };
}

describe("merging workbooks read in pieces", () => {
  it.each([
    ["with the sheet index", {}],
    ["without the sheet index", { includeSheetIndex: false }],
    ["values only", { values: true }],
  ])(
    "merges %s without a whole read, as the byte surface does",
    async (_, options) => {
      const north = await workbook("North", 9_000);
      const south = await workbook("South", 9_000);
      expect(north.length).toBeGreaterThan(MIB);
      const reference = await mergeWorkbooksBytes({
        ...options,
        inputs: [
          { name: "north.xlsx", bytes: north },
          { name: "south.xlsx", bytes: south },
        ],
      });

      const chunks: Uint8Array[] = [];
      const outcome = await mergeWorkbookSources({
        ...options,
        inputs: [
          guardedSource("north.xlsx", north),
          guardedSource("south.xlsx", south),
        ],
        output: {
          write: (chunk) => {
            chunks.push(chunk.slice());
          },
          flush: () => Promise.resolve(),
          abort: () => Promise.resolve(),
        },
      });

      expect(outcome.result).toEqual(reference.result);
      expect(outcome.outputName).toBe(reference.outputs[0]!.name);
      expect(chunks.length).toBeGreaterThan(1);
      expect(
        Buffer.concat(chunks).equals(Buffer.from(reference.outputs[0]!.bytes)),
      ).toBe(true);
      expect(outcome.result.metrics.outputSheets).toBe(2);
    },
    120_000,
  );

  it("discards what it wrote when an input cannot be read", async () => {
    let aborted = 0;
    await expect(
      mergeWorkbookSources({
        inputs: [
          guardedSource("north.xlsx", await workbook("North", 20)),
          {
            name: "broken.xlsx",
            size: 100,
            readAt: () => Promise.resolve(new Uint8Array(100)),
          },
        ],
        output: {
          write: () => undefined,
          flush: () => Promise.resolve(),
          abort: () => {
            aborted += 1;
            return Promise.resolve();
          },
        },
      }),
    ).rejects.toMatchObject({ code: "XLSX_READ_FAILED" });
    expect(aborted).toBe(1);
  });
});
