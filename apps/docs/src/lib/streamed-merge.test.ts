import { mergeWorkbooksBytes } from "@consultchimps/xlsx/bytes";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import type { NamedFile } from "./operation-tasks";
import { mergeFiles } from "./streamed-merge";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** A one-sheet workbook of `rows` rows, stored uncompressed so it is large. */
async function workbook(prefix: string, rows: number): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Log" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  );
  const cell = (ref: string, text: string) =>
    `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;
  const body = [`<row r="1">${cell("A1", "Case")}${cell("B1", "Note")}</row>`];
  for (let row = 2; row <= rows; row += 1) {
    body.push(
      `<row r="${row}">${cell(`A${row}`, `${prefix}-${row}`)}${cell(`B${row}`, `A longer note for row ${row}`)}</row>`,
    );
  }
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<worksheet xmlns="${MAIN}"><sheetData>${body.join("")}</sheetData></worksheet>`,
  );
  return zip.generateAsync({ type: "uint8array", compression: "STORE" });
}

/** A file whose reads fail when one asks for more than 1 MiB at once. */
function guardedFile(name: string, bytes: Uint8Array): NamedFile {
  const blob = new Blob([bytes.slice()]);
  const guarded = {
    size: blob.size,
    slice: (start: number, end: number) => {
      if (end - start > 1024 * 1024) {
        throw new Error(`Read ${end - start} of ${blob.size} bytes at once`);
      }
      return blob.slice(start, end);
    },
    arrayBuffer: () => Promise.reject(new Error("Read the whole file")),
  };
  return { name, file: guarded as unknown as Blob };
}

const controls = () => ({
  signal: new AbortController().signal,
  onProgress: () => undefined,
});

/** No origin private file system: the output is gathered in memory. */
const inMemory = { storage: undefined, locks: undefined };

describe("mergeFiles", () => {
  it("reads the files in pieces and gives the byte surface's workbook", async () => {
    const north = await workbook("N", 12_000);
    const south = await workbook("S", 12_000);
    expect(north.length).toBeGreaterThan(1024 * 1024);
    const reference = await mergeWorkbooksBytes({
      inputs: [
        { name: "north.xlsx", bytes: north },
        { name: "south.xlsx", bytes: south },
      ],
    });

    const { result, outputs } = await mergeFiles(
      [guardedFile("north.xlsx", north), guardedFile("south.xlsx", south)],
      {},
      controls(),
      inMemory,
      new Set(),
    );

    expect(result).toEqual(reference.result);
    expect(outputs.map((output) => output.name)).toEqual(["merged.xlsx"]);
    expect(
      Buffer.from(await outputs[0]!.blob.arrayBuffer()).equals(
        Buffer.from(reference.outputs[0]!.bytes),
      ),
    ).toBe(true);
  }, 60_000);

  it("merges a CSV file in pieces as the byte surface does", async () => {
    const lines = ["Region;Amount"];
    for (let row = 1; row <= 90_000; row += 1)
      lines.push(`R-${row % 7};${row},5`);
    const csv = new TextEncoder().encode(lines.join("\r\n"));
    expect(csv.length).toBeGreaterThan(1024 * 1024);
    const north = await workbook("N", 20);
    const options = { csv: { numbers: true, decimalSeparator: "," } } as const;
    const reference = await mergeWorkbooksBytes({
      inputs: [
        { name: "north.xlsx", bytes: north },
        { name: "amounts.csv", bytes: csv },
      ],
      ...options,
    });

    const { result, outputs } = await mergeFiles(
      [guardedFile("north.xlsx", north), guardedFile("amounts.csv", csv)],
      options,
      controls(),
      inMemory,
      new Set(),
    );

    expect(result).toEqual(reference.result);
    expect(result.metrics.csvInputFiles).toBe(1);
    expect(
      Buffer.from(await outputs[0]!.blob.arrayBuffer()).equals(
        Buffer.from(reference.outputs[0]!.bytes),
      ),
    ).toBe(true);
  }, 60_000);

  it("fails as an unreadable file, not a damaged workbook", async () => {
    const bytes = await workbook("N", 20);
    const blob = new Blob([bytes.slice()]);
    const gone = {
      size: blob.size,
      slice: () => ({
        arrayBuffer: () => Promise.reject(new Error("NotReadableError")),
      }),
    } as unknown as Blob;
    await expect(
      mergeFiles(
        [{ name: "north.xlsx", file: gone }],
        {},
        controls(),
        inMemory,
        new Set(),
      ),
    ).rejects.toMatchObject({ code: "FILE_UNREADABLE" });
  });
});
