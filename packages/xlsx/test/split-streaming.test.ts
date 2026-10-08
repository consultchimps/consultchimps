/**
 * Splitting reads its input in pieces and writes each output as it is
 * produced (ADR 0006), in every mode, and gives the bytes the byte surface
 * gives. The row relocation it relies on handles a worksheet of any size: a
 * spread of every row into one call's arguments overflowed the stack.
 */
import type { RandomAccessSource } from "@consultchimps/core";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import {
  blobSource,
  planSplitWorkbookBytes,
  planSplitWorkbookSource,
  splitWorkbookBytes,
  splitWorkbookSource,
  type SplitWorkbookBytesOptions,
} from "../src/bytes.js";
import { RowRelocation } from "../src/model/references.js";
import { WorksheetModel } from "../src/model/worksheet-model.js";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MIB = 1024 * 1024;
const ROWS = 9_000;
const REGIONS = ["North", "South", "East"];

function cell(ref: string, text: string): string {
  return `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;
}

/**
 * A stored workbook over 1 MiB: a header and rows on "Data", an Excel Table
 * and a named range over them, and a second sheet without the column.
 */
async function workbook(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Notes" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="Area">Data!$A$1:$C$${ROWS + 1}</definedName></definedNames></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
  );
  const rows = [
    `<row r="1">${cell("A1", "Region")}${cell("B1", "Case")}${cell("C1", "Note")}</row>`,
  ];
  for (let row = 2; row <= ROWS + 1; row += 1) {
    rows.push(
      `<row r="${row}">${cell(`A${row}`, REGIONS[row % 3]!)}${cell(`B${row}`, `C-${row}`)}${cell(`C${row}`, `A note for case ${row}`)}</row>`,
    );
  }
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<worksheet xmlns="${MAIN}" xmlns:r="${REL}"><dimension ref="A1:C${ROWS + 1}"/><sheetData>${rows.join("")}</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`,
  );
  zip.file(
    "xl/worksheets/_rels/sheet1.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/table" Target="../tables/table1.xml"/></Relationships>`,
  );
  zip.file(
    "xl/tables/table1.xml",
    `<table xmlns="${MAIN}" id="1" name="Cases" displayName="Cases" ref="A1:C${ROWS + 1}"><autoFilter ref="A1:C${ROWS + 1}"/><tableColumns count="3"><tableColumn id="1" name="Region"/><tableColumn id="2" name="Case"/><tableColumn id="3" name="Note"/></tableColumns></table>`,
  );
  zip.file(
    "xl/worksheets/sheet2.xml",
    `<worksheet xmlns="${MAIN}"><sheetData><row r="1">${cell("A1", "Topic")}</row><row r="2">${cell("A2", "Kept")}</row></sheetData></worksheet>`,
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

describe("splitting a workbook read in pieces", () => {
  const modes: Array<[string, Omit<SplitWorkbookBytesOptions, "input">]> = [
    ["every worksheet", { column: "Region" }],
    ["every worksheet, values only", { column: "Region", values: true }],
    [
      "a compact rebuild leaving blanks out",
      { column: "Region", sheet: "Data", includeBlank: false },
    ],
    ["one worksheet", { column: "Region", sheet: "Data" }],
    [
      "an Excel Table, keeping the workbook",
      { column: "Region", table: "Cases" },
    ],
    [
      "an Excel Table, values only",
      { column: "Region", table: "Cases", values: true },
    ],
    [
      "an Excel Table, compact",
      { column: "Region", table: "Cases", preserveWorkbook: false },
    ],
    ["a named range", { column: "Region", range: "Area" }],
  ];

  it.each(modes)(
    "splits %s without a whole read, as the byte surface does",
    async (_, options) => {
      const bytes = await workbook();
      expect(bytes.length).toBeGreaterThan(MIB);
      const reference = await splitWorkbookBytes({
        ...options,
        input: { name: "cases.xlsx", bytes },
      });

      const written = new Map<string, Uint8Array[]>();
      const outcome = await splitWorkbookSource({
        ...options,
        input: guardedSource("cases.xlsx", bytes),
        output: (name) => {
          const chunks: Uint8Array[] = [];
          written.set(name, chunks);
          return {
            write: (chunk) => {
              chunks.push(chunk.slice());
            },
            flush: () => Promise.resolve(),
            abort: () => Promise.resolve(),
          };
        },
      });

      expect(outcome.result).toEqual(reference.result);
      expect([...written.keys()]).toEqual(
        reference.outputs.map((output) => output.name),
      );
      for (const output of reference.outputs) {
        const chunks = written.get(output.name)!;
        expect(chunks.length).toBeGreaterThan(1);
        expect(Buffer.concat(chunks)).toEqual(Buffer.from(output.bytes));
      }
      expect(reference.outputs).toHaveLength(3);

      const plan = await planSplitWorkbookSource({
        ...options,
        input: guardedSource("cases.xlsx", bytes),
      });
      expect(plan).toEqual(
        await planSplitWorkbookBytes({
          ...options,
          input: { name: "cases.xlsx", bytes },
        }),
      );
    },
    120_000,
  );
});

describe("row relocation at a size a spread overflows", () => {
  const LARGE = 200_000;

  it("builds an explicit relocation over every row", () => {
    const entries: Array<[number, number | null]> = [];
    for (let row = 1; row <= LARGE; row += 1) entries.push([row, null]);
    const relocation = RowRelocation.explicit(entries);
    expect(relocation.target(LARGE)).toBeNull();
    expect(relocation.target(LARGE + 5)).toBe(LARGE + 5);
  });

  it("deletes a set of rows larger than any argument list", () => {
    const worksheet = WorksheetModel.parse(
      `<worksheet xmlns="${MAIN}"><sheetData><row r="1">${cell("A1", "Kept")}</row></sheetData></worksheet>`,
      {
        name: "Data",
        partPath: "xl/worksheets/sheet1.xml",
        visibility: "visible",
      },
      {
        markWorksheetChanged: () => undefined,
        relocateTables: () => undefined,
        relocateCalcChain: () => undefined,
        relocateComments: () => undefined,
      } as never,
    );
    const doomed = new Set<number>();
    for (let row = 2; row <= LARGE; row += 1) doomed.add(row);
    expect(() =>
      worksheet.deleteRows(doomed, { renumber: true }),
    ).not.toThrow();
    expect(worksheet.lastRow).toBe(1);
  });

  it("measures the used range of a worksheet with no dimension and many cells", () => {
    const rows: string[] = [];
    for (let row = 1; row <= LARGE; row += 1) {
      rows.push(`<row r="${row}"><c r="B${row}"><v>${row}</v></c></row>`);
    }
    const worksheet = WorksheetModel.parse(
      `<worksheet xmlns="${MAIN}"><sheetData>${rows.join("")}</sheetData></worksheet>`,
      {
        name: "Data",
        partPath: "xl/worksheets/sheet1.xml",
        visibility: "visible",
      },
      {} as never,
    );
    expect(worksheet.usedRange).toEqual({
      start: { row: 1, column: 1 },
      end: { row: LARGE, column: 1 },
    });
  }, 60_000);
});
