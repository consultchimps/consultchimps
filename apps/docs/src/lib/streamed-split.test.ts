import { splitWorkbookBytes } from "@consultchimps/xlsx/bytes";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import type { NamedFile } from "./operation-tasks";
import { planSplitFile, splitFile } from "./streamed-split";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** A one-sheet workbook of `rows` rows, stored uncompressed so it is large. */
async function workbook(rows: number): Promise<Uint8Array> {
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
  const body = [
    `<row r="1">${cell("A1", "Region")}${cell("B1", "Note")}</row>`,
  ];
  for (let row = 2; row <= rows; row += 1) {
    body.push(
      `<row r="${row}">${cell(`A${row}`, row % 2 ? "North" : "South")}${cell(`B${row}`, `A longer note for row ${row}`)}</row>`,
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

/** No origin private file system: outputs are gathered in memory. */
const inMemory = { storage: undefined, locks: undefined };

describe("splitFile", () => {
  it.each([
    ["keeping the workbook", { column: "Region" }],
    ["compact", { column: "Region", sheet: "Log" }],
  ])(
    "reads the file in pieces and gives the byte surface's outputs, %s",
    async (_, options) => {
      const bytes = await workbook(12_000);
      expect(bytes.length).toBeGreaterThan(1024 * 1024);
      const reference = await splitWorkbookBytes({
        ...options,
        input: { name: "log.xlsx", bytes },
      });

      const { result, outputs } = await splitFile(
        guardedFile("log.xlsx", bytes),
        options,
        controls(),
        inMemory,
        new Set(),
      );

      expect(result).toEqual(reference.result);
      expect(outputs.map((output) => output.name)).toEqual(
        reference.outputs.map((output) => output.name),
      );
      for (const [index, output] of outputs.entries()) {
        expect(new Uint8Array(await output.blob.arrayBuffer())).toEqual(
          reference.outputs[index]!.bytes,
        );
      }
      const plan = await planSplitFile(guardedFile("log.xlsx", bytes), options);
      expect(plan.metrics.groups).toBe(2);
    },
    60_000,
  );

  it("fails as an unreadable file, not a damaged workbook", async () => {
    const bytes = await workbook(20);
    const blob = new Blob([bytes.slice()]);
    const gone = {
      size: blob.size,
      slice: () => ({
        arrayBuffer: () => Promise.reject(new Error("NotReadableError")),
      }),
    } as unknown as Blob;
    await expect(
      splitFile(
        { name: "log.xlsx", file: gone },
        { column: "Region" },
        controls(),
        inMemory,
        new Set(),
      ),
    ).rejects.toMatchObject({ code: "FILE_UNREADABLE" });
  });

  it("answers a cancel while a read of the file is still waiting", async () => {
    const bytes = await workbook(20);
    const stalled = {
      size: bytes.length,
      slice: () => ({ arrayBuffer: () => new Promise<never>(() => undefined) }),
    } as unknown as Blob;
    const controller = new AbortController();
    const split = splitFile(
      { name: "log.xlsx", file: stalled },
      { column: "Region" },
      { signal: controller.signal, onProgress: () => undefined },
      inMemory,
      new Set(),
    );
    controller.abort();
    await expect(split).rejects.toMatchObject({ code: "OPERATION_ABORTED" });
  });
});
