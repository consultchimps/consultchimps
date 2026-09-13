import type { RandomAccessFile } from "@consultchimps/core";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import {
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

function source(bytes: Uint8Array) {
  return {
    name: "namespace-values.xlsx",
    size: bytes.byteLength,
    async readAt(offset: number, length: number) {
      return bytes.slice(offset, offset + length);
    },
  };
}

async function fixture(options: {
  readonly spreadsheet: string;
  readonly relationships: string;
  readonly worksheet: string;
  readonly sharedStrings?: string | undefined;
}): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>",
  );
  zip.file(
    "_rels/.rels",
    `<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='office' Type='${options.relationships}/officeDocument' Target='xl/workbook.xml'/></Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `<x:workbook xmlns:x='${options.spreadsheet}' xmlns:r='${options.relationships}'><x:sheets><x:sheet name='Data' r:id='sheet'/></x:sheets></x:workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='sheet' Type='${options.relationships}/worksheet' Target='worksheets/sheet1.xml'/>${options.sharedStrings === undefined ? "" : `<Relationship Id='strings' Type='${options.relationships}/sharedStrings' Target='sharedStrings.xml'/>`}</Relationships>`,
  );
  zip.file("xl/worksheets/sheet1.xml", options.worksheet);
  if (options.sharedStrings !== undefined) {
    zip.file("xl/sharedStrings.xml", options.sharedStrings);
  }
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

async function readRows(bytes: Uint8Array): Promise<readonly StreamRow[]> {
  const reader = await openWorkbookRegionStream(
    source(bytes),
    { sheet: "Data", headerRow: 1 },
    { scratch: new MemoryScratch() },
  );
  const rows: StreamRow[] = [];
  for await (const batch of reader.batches({ batchSize: 1 })) {
    rows.push(...batch);
  }
  await reader.close();
  return rows;
}

describe.each(variants)("$label worksheet value namespaces", (variant) => {
  it("ignores foreign cells and values below extension elements", async () => {
    const worksheet =
      `<x:worksheet xmlns:x='${variant.spreadsheet}' xmlns:ext='urn:synthetic-extension'><x:sheetData>` +
      "<ext:row r='1'><ext:c r='A1' t='inlineStr'><ext:is><ext:t>Foreign header</ext:t></ext:is></ext:c></ext:row>" +
      "<ext:container><x:row r='1'><x:c r='A1' t='inlineStr'><x:is><x:t>Nested header</x:t></x:is></x:c></x:row></ext:container>" +
      "<x:row r='1'><x:c r='A1' t='inlineStr'><x:is><x:t>Name</x:t></x:is></x:c><x:c r='B1' t='inlineStr'><x:is><x:t>Count</x:t></x:is></x:c><x:c r='C1' t='inlineStr'><x:is><x:t>Plain</x:t></x:is></x:c></x:row>" +
      "<x:row r='2'><x:c r='A2' t='inlineStr'><x:is><x:t>North</x:t><ext:t> foreign</ext:t><ext:wrapper><x:t> nested</x:t></ext:wrapper></x:is></x:c>" +
      "<x:c r='B2'><x:v>7</x:v><ext:v>9</ext:v><ext:wrapper><x:v>8</x:v></ext:wrapper></x:c>" +
      "<x:c r='C2'><ext:f>1+1</ext:f><x:v>3</x:v></x:c>" +
      "<ext:c r='B2'><ext:v>99</ext:v></ext:c><ext:wrapper><x:c r='B2'><x:v>88</x:v></x:c></ext:wrapper></x:row>" +
      "</x:sheetData><ext:row r='3'><ext:c r='A3'><ext:v>4</ext:v></ext:c></ext:row></x:worksheet>";
    const bytes = await fixture({ ...variant, worksheet });

    expect(await readRows(bytes)).toEqual([
      {
        sourceRow: 2,
        cells: {
          Name: { kind: "string", value: "North" },
          Count: { kind: "number", raw: "7" },
          Plain: { kind: "number", raw: "3" },
        },
      },
    ]);
  });

  it("ignores foreign and nested shared-string items and text", async () => {
    const worksheet =
      `<x:worksheet xmlns:x='${variant.spreadsheet}'><x:sheetData>` +
      "<x:row r='1'><x:c r='A1' t='s'><x:v>0</x:v></x:c></x:row>" +
      "<x:row r='2'><x:c r='A2' t='s'><x:v>1</x:v></x:c></x:row>" +
      "</x:sheetData></x:worksheet>";
    const sharedStrings =
      `<x:sst xmlns:x='${variant.spreadsheet}' xmlns:ext='urn:synthetic-extension'>` +
      "<ext:si><ext:t>Foreign index</ext:t></ext:si>" +
      "<ext:container><x:si><x:t>Nested index</x:t></x:si></ext:container>" +
      "<x:si><x:t>Name</x:t><ext:t> foreign</ext:t><ext:wrapper><x:t> nested</x:t></ext:wrapper></x:si>" +
      "<x:si><x:r><x:t>Nor</x:t></x:r><x:r><x:t>th</x:t></x:r><x:rPh><x:t>phonetic</x:t></x:rPh></x:si>" +
      "</x:sst>";
    const bytes = await fixture({ ...variant, worksheet, sharedStrings });

    expect(await readRows(bytes)).toEqual([
      {
        sourceRow: 2,
        cells: { Name: { kind: "string", value: "North" } },
      },
    ]);
  });
});
