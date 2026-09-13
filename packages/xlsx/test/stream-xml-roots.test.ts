import type { RandomAccessFile } from "@consultchimps/core";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import {
  inspectWorkbookStream,
  openWorkbookRegionStream,
  type ScratchFactory,
  type StreamRow,
} from "../src/stream.js";

class MemoryFile implements RandomAccessFile {
  #bytes = new Uint8Array();

  constructor(readonly name: string) {}

  get size(): number {
    return this.#bytes.byteLength;
  }

  async readAt(offset: number, length: number): Promise<Uint8Array> {
    return this.#bytes.slice(offset, offset + length);
  }

  async writeAt(offset: number, bytes: Uint8Array): Promise<void> {
    const required = offset + bytes.byteLength;
    if (required > this.#bytes.byteLength) {
      const expanded = new Uint8Array(required);
      expanded.set(this.#bytes);
      this.#bytes = expanded;
    }
    this.#bytes.set(bytes, offset);
  }

  async truncate(size: number): Promise<void> {
    this.#bytes = this.#bytes.slice(0, size);
  }

  async close(): Promise<void> {}
}

class MemoryScratch implements ScratchFactory {
  #nextId = 0;

  async create(): Promise<MemoryFile> {
    return new MemoryFile(`scratch-${this.#nextId++}`);
  }
}

const variants = [
  {
    label: "Transitional",
    spreadsheet: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    relationships:
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  },
  {
    label: "Strict",
    spreadsheet: "http://purl.oclc.org/ooxml/spreadsheetml/main",
    relationships: "http://purl.oclc.org/ooxml/officeDocument/relationships",
  },
] as const;

const packageRelationships =
  "http://schemas.openxmlformats.org/package/2006/relationships";

function source(bytes: Uint8Array) {
  return {
    name: "xml-roots.xlsx",
    size: bytes.byteLength,
    async readAt(offset: number, length: number) {
      return bytes.slice(offset, offset + length);
    },
  };
}

function parts(variant: (typeof variants)[number]): Record<string, string> {
  const { spreadsheet, relationships } = variant;
  return {
    "[Content_Types].xml": `<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>`,
    "_rels/.rels": `<Relationships xmlns='${packageRelationships}'><Relationship Id='office' Type='${relationships}/officeDocument' Target='xl/workbook.xml'/></Relationships>`,
    "xl/workbook.xml": `<x:workbook xmlns:x='${spreadsheet}' xmlns:r='${relationships}'><x:sheets><x:sheet name='Data' r:id='sheet'/></x:sheets></x:workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships xmlns='${packageRelationships}'><Relationship Id='sheet' Type='${relationships}/worksheet' Target='worksheets/sheet1.xml'/><Relationship Id='strings' Type='${relationships}/sharedStrings' Target='sharedStrings.xml'/><Relationship Id='styles' Type='${relationships}/styles' Target='styles.xml'/></Relationships>`,
    "xl/worksheets/sheet1.xml": `<x:worksheet xmlns:x='${spreadsheet}' xmlns:r='${relationships}'><x:sheetData><x:row r='1'><x:c r='A1' t='s'><x:v>0</x:v></x:c><x:c r='B1' t='s'><x:v>1</x:v></x:c></x:row><x:row r='2'><x:c r='A2' t='s'><x:v>2</x:v></x:c><x:c r='B2'><x:v>7</x:v></x:c></x:row></x:sheetData><x:tableParts count='1'><x:tablePart r:id='table'/></x:tableParts></x:worksheet>`,
    "xl/worksheets/_rels/sheet1.xml.rels": `<Relationships xmlns='${packageRelationships}'><Relationship Id='table' Type='${relationships}/table' Target='../tables/table1.xml'/></Relationships>`,
    "xl/tables/table1.xml": `<x:table xmlns:x='${spreadsheet}' name='Inventory' displayName='Inventory' ref='A1:B2'><x:tableColumns count='2'><x:tableColumn id='1' name='Name'/><x:tableColumn id='2' name='Count'/></x:tableColumns></x:table>`,
    "xl/sharedStrings.xml": `<x:sst xmlns:x='${spreadsheet}'><x:si><x:t>Name</x:t></x:si><x:si><x:t>Count</x:t></x:si><x:si><x:t>North</x:t></x:si></x:sst>`,
    "xl/styles.xml": `<x:styleSheet xmlns:x='${spreadsheet}'><x:cellXfs count='1'><x:xf numFmtId='0'/></x:cellXfs></x:styleSheet>`,
  };
}

async function fixture(
  variant: (typeof variants)[number],
  replace: Readonly<Record<string, string>> = {},
): Promise<Uint8Array> {
  const zip = new JSZip();
  const contents = { ...parts(variant), ...replace };
  for (const [name, value] of Object.entries(contents)) zip.file(name, value);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

async function rows(bytes: Uint8Array): Promise<readonly StreamRow[]> {
  const reader = await openWorkbookRegionStream(
    source(bytes),
    { table: "Inventory" },
    { scratch: new MemoryScratch() },
  );
  const result: StreamRow[] = [];
  for await (const batch of reader.batches({ batchSize: 1 })) {
    result.push(...batch);
  }
  await reader.close();
  return result;
}

describe("streamed XML document roots", () => {
  it.each(variants)("reads a supported $label workbook", async (variant) => {
    const bytes = await fixture(variant);

    expect(await rows(bytes)).toEqual([
      {
        sourceRow: 2,
        cells: {
          Name: { kind: "string", value: "North" },
          Count: { kind: "number", raw: "7" },
        },
      },
    ]);
  });

  it.each([
    [
      "wrong local name",
      "<notWorkbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'/>",
    ],
    [
      "wrong namespace",
      "<workbook xmlns='urn:synthetic-unsupported'><sheets/></workbook>",
    ],
    ["missing namespace", "<workbook><sheets/></workbook>"],
  ])("rejects a workbook with a %s", async (_case, workbook) => {
    const bytes = await fixture(variants[0], {
      "xl/workbook.xml": workbook,
    });

    await expect(
      inspectWorkbookStream(source(bytes), { scratch: new MemoryScratch() }),
    ).rejects.toMatchObject({
      code: "XLSX_READ_FAILED",
      cause: expect.objectContaining({
        message: expect.stringMatching(/must have a workbook root/iu),
      }),
    });
  });

  it.each([
    ["package relationships", "_rels/.rels", "notRelationships"],
    [
      "workbook relationships",
      "xl/_rels/workbook.xml.rels",
      "notRelationships",
    ],
    ["worksheet", "xl/worksheets/sheet1.xml", "notWorksheet"],
    [
      "worksheet relationships",
      "xl/worksheets/_rels/sheet1.xml.rels",
      "notRelationships",
    ],
    ["table", "xl/tables/table1.xml", "notTable"],
  ])("rejects a wrong %s root during inspection", async (_case, part, root) => {
    const bytes = await fixture(variants[0], {
      [part]: `<${root} xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'/>`,
    });

    await expect(
      inspectWorkbookStream(source(bytes), { scratch: new MemoryScratch() }),
    ).rejects.toMatchObject({
      code: "XLSX_READ_FAILED",
      cause: expect.objectContaining({
        message: expect.stringMatching(/must have a .* root/iu),
      }),
    });
  });

  it.each([
    ["shared-string", "xl/sharedStrings.xml", "notSst"],
    ["style", "xl/styles.xml", "notStyleSheet"],
  ])(
    "rejects a wrong %s root before reading rows",
    async (_case, part, root) => {
      const bytes = await fixture(variants[0], {
        [part]: `<${root} xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'/>`,
      });

      await expect(rows(bytes)).rejects.toMatchObject({
        code: "XLSX_READ_FAILED",
        cause: expect.objectContaining({
          message: expect.stringMatching(/must have a .* root/iu),
        }),
      });
    },
  );
});
