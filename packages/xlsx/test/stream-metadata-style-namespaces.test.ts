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
    packageRelationships:
      "http://schemas.openxmlformats.org/package/2006/relationships",
  },
  {
    label: "Strict",
    spreadsheet: "http://purl.oclc.org/ooxml/spreadsheetml/main",
    relationships: "http://purl.oclc.org/ooxml/officeDocument/relationships",
    packageRelationships:
      "http://schemas.openxmlformats.org/package/2006/relationships",
  },
] as const;

function source(bytes: Uint8Array) {
  return {
    name: "namespace-metadata.xlsx",
    size: bytes.byteLength,
    async readAt(offset: number, length: number) {
      return bytes.slice(offset, offset + length);
    },
  };
}

async function fixture(options: {
  readonly spreadsheet: string;
  readonly relationships: string;
  readonly packageRelationships: string;
  readonly hostileMetadata?: boolean | undefined;
  readonly hostileStyles?: boolean | undefined;
}): Promise<Uint8Array> {
  const zip = new JSZip();
  const foreign = "urn:synthetic-extension";
  const fakePackageRelationships = options.hostileMetadata
    ? `<ext:Relationship Id='office' Type='${options.relationships}/officeDocument' Target='xl/wrong.xml'/><Container><Relationship Id='office' Type='${options.relationships}/officeDocument' Target='xl/wrong.xml'/></Container>`
    : "";
  const fakeWorkbookRelationships = options.hostileMetadata
    ? `<ext:Relationship Id='sheet' Type='${options.relationships}/worksheet' Target='worksheets/wrong.xml'/><Container><Relationship Id='sheet' Type='${options.relationships}/worksheet' Target='worksheets/wrong.xml'/></Container>`
    : "";
  const fakeTableRelationships = options.hostileMetadata
    ? `<ext:Relationship Id='table' Type='${options.relationships}/table' Target='../tables/wrong.xml'/><Container><Relationship Id='table' Type='${options.relationships}/table' Target='../tables/wrong.xml'/></Container>`
    : "";
  const styles = options.hostileStyles
    ? `<x:styleSheet xmlns:x='${options.spreadsheet}' xmlns:ext='${foreign}'>` +
      "<ext:numFmt numFmtId='164' formatCode='General'/><x:numFmt numFmtId='164' formatCode='General'/>" +
      "<x:numFmts count='1'><ext:numFmt numFmtId='164' formatCode='General'/><x:numFmt ext:numFmtId='0' numFmtId='164' formatCode='yyyy-mm-dd'/></x:numFmts>" +
      "<x:cellXfs count='1'><ext:xf numFmtId='0'/><ext:wrapper><x:xf numFmtId='0'/></ext:wrapper><x:xf ext:numFmtId='0' numFmtId='164'/></x:cellXfs>" +
      "<x:xf numFmtId='0'/></x:styleSheet>"
    : `<x:styleSheet xmlns:x='${options.spreadsheet}'><x:cellXfs count='1'><x:xf numFmtId='14'/></x:cellXfs></x:styleSheet>`;
  const table = options.hostileMetadata
    ? `<x:table xmlns:x='${options.spreadsheet}' xmlns:ext='${foreign}' ext:name='Wrong' name='Inventory' displayName='Inventory' ref='A1:A2'>` +
      "<ext:tableColumns><x:tableColumn name='Nested'/></ext:tableColumns><x:tableColumn name='Misplaced'/>" +
      "<x:tableColumns count='1'><ext:tableColumn name='Foreign'/><x:tableColumn ext:name='Wrong' id='1' name='Date'/></x:tableColumns></x:table>"
    : `<x:table xmlns:x='${options.spreadsheet}' name='Inventory' displayName='Inventory' ref='A1:A2'><x:tableColumns count='1'><x:tableColumn id='1' name='Date'/></x:tableColumns></x:table>`;

  zip.file(
    "[Content_Types].xml",
    "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>",
  );
  zip.file(
    "_rels/.rels",
    `<Relationships xmlns='${options.packageRelationships}' xmlns:ext='${foreign}'>${fakePackageRelationships}<Relationship Id='office' Type='${options.relationships}/officeDocument' Target='xl/workbook.xml'/></Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `<x:workbook xmlns:x='${options.spreadsheet}' xmlns:r='${options.relationships}' xmlns:ext='${foreign}'><x:sheets><x:sheet ext:name='Wrong' name='Data' ext:id='wrong' r:id='sheet'/></x:sheets></x:workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<Relationships xmlns='${options.packageRelationships}' xmlns:ext='${foreign}'>${fakeWorkbookRelationships}<Relationship Id='sheet' Type='${options.relationships}/worksheet' Target='worksheets/sheet1.xml'/><Relationship Id='styles' Type='${options.relationships}/styles' Target='styles.xml'/></Relationships>`,
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<x:worksheet xmlns:x='${options.spreadsheet}' xmlns:r='${options.relationships}'><x:sheetData><x:row r='1'><x:c r='A1' t='inlineStr'><x:is><x:t>Date</x:t></x:is></x:c></x:row><x:row r='2'><x:c r='A2' s='0'><x:v>1</x:v></x:c></x:row></x:sheetData><x:tableParts count='1'><x:tablePart r:id='table'/></x:tableParts></x:worksheet>`,
  );
  zip.file(
    "xl/worksheets/_rels/sheet1.xml.rels",
    `<Relationships xmlns='${options.packageRelationships}' xmlns:ext='${foreign}'>${fakeTableRelationships}<Relationship Id='table' Type='${options.relationships}/table' Target='../tables/table1.xml'/></Relationships>`,
  );
  zip.file("xl/tables/table1.xml", table);
  zip.file("xl/styles.xml", styles);
  zip.file("xl/wrong.xml", `<x:workbook xmlns:x='${options.spreadsheet}'/>`);
  zip.file(
    "xl/worksheets/wrong.xml",
    `<x:worksheet xmlns:x='${options.spreadsheet}'/>`,
  );
  zip.file("xl/tables/wrong.xml", table);
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

describe.each(variants)("$label metadata and style namespaces", (variant) => {
  it("ignores foreign and misplaced relationships and table columns", async () => {
    const bytes = await fixture({ ...variant, hostileMetadata: true });

    await expect(
      inspectWorkbookStream(source(bytes), { scratch: new MemoryScratch() }),
    ).resolves.toMatchObject({
      sheets: [{ name: "Data", visibility: "visible" }],
      tables: [
        {
          name: "Inventory",
          sheet: "Data",
          reference: "A1:A2",
          columns: ["Date"],
        },
      ],
    });
  });

  it("uses only direct SpreadsheetML number formats and cell styles", async () => {
    const bytes = await fixture({ ...variant, hostileStyles: true });

    expect(await rows(bytes)).toEqual([
      {
        sourceRow: 2,
        cells: { Date: { kind: "date", raw: "1", iso: "1900-01-01" } },
      },
    ]);
  });
});
