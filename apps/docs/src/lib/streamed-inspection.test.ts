import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { describeWorkbook } from "@consultchimps/xlsx";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import type { NamedFile } from "./operation-tasks";
import { inspectFile } from "./streamed-inspection";

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
    `<row r="1">${cell("A1", "Report")}</row>`,
    `<row r="3">${cell("A3", "Case_ID")}${cell("B3", "Note")}${cell("C3", "Region")}</row>`,
  ];
  for (let row = 4; row <= rows; row += 1) {
    body.push(
      `<row r="${row}">${cell(`A${row}`, `C-${row}`)}${cell(`B${row}`, `A longer note for row ${row}`)}${cell(`C${row}`, row % 2 ? "North" : "South")}</row>`,
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

describe("inspectFile", () => {
  it("reads the file in pieces and gives the command line's description", async () => {
    const bytes = await workbook(12_000);
    expect(bytes.length).toBeGreaterThan(1024 * 1024);
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-docs-"));
    try {
      const file = path.join(directory, "log.xlsx");
      await writeFile(file, bytes);
      const reference = await describeWorkbook(file);

      const outcome = await inspectFile(guardedFile("log.xlsx", bytes), {});

      expect(outcome).toEqual(reference);
      expect(outcome.description.sheets[0]).toMatchObject({
        dataRowCount: 11_997,
        headerRow: 3,
      });
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  }, 60_000);

  it("reads a CSV file in pieces and gives the command line's description", async () => {
    const lines = ["Case_ID;Amount"];
    for (let row = 1; row <= 80_000; row += 1) {
      lines.push(`C-${row};${row},5`);
    }
    const bytes = new TextEncoder().encode(lines.join("\r\n"));
    expect(bytes.length).toBeGreaterThan(1024 * 1024);
    const csv = { numbers: true, decimalSeparator: "," } as const;
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-docs-"));
    try {
      const file = path.join(directory, "cases.csv");
      await writeFile(file, bytes);
      const reference = await describeWorkbook(file, { csv });

      const outcome = await inspectFile(guardedFile("cases.csv", bytes), {
        csv,
      });

      expect(outcome).toEqual(reference);
      expect(outcome.description.csv).toEqual({
        encoding: "utf-8",
        delimiter: ";",
      });
      expect(outcome.description.sheets[0]?.columns[1]?.sampleValues[0]).toBe(
        1.5,
      );
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  }, 60_000);

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
      inspectFile({ name: "log.xlsx", file: gone }, {}),
    ).rejects.toMatchObject({
      code: "FILE_UNREADABLE",
      details: { source: "log.xlsx" },
    });
  });

  it("answers a cancel while a read of the file is still waiting", async () => {
    const bytes = await workbook(20);
    const stalled = {
      size: bytes.length,
      slice: () => ({ arrayBuffer: () => new Promise<never>(() => undefined) }),
    } as unknown as Blob;
    const controller = new AbortController();
    const inspection = inspectFile(
      { name: "log.xlsx", file: stalled },
      { signal: controller.signal },
    );
    controller.abort();
    await expect(inspection).rejects.toMatchObject({
      code: "OPERATION_ABORTED",
    });
  });

  it.each([
    ["uses", true],
    ["never uses", false],
  ])(
    "fails as an unreadable file whichever single read the browser refuses, when the sheet %s the shared strings",
    async (_, usesStrings) => {
      const zip = new JSZip();
      const source = await JSZip.loadAsync(await workbook(200));
      for (const [name, entry] of Object.entries(source.files)) {
        zip.file(name, await entry.async("uint8array"));
      }
      // Shared strings, which the reader loads apart from the worksheet.
      zip.file(
        "xl/sharedStrings.xml",
        `<sst xmlns="${MAIN}"><si><t>Report</t></si></sst>`,
      );
      zip.file(
        "[Content_Types].xml",
        (await zip.file("[Content_Types].xml")!.async("string")).replace(
          "</Types>",
          '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
        ),
      );
      if (usesStrings) {
        zip.file(
          "xl/worksheets/sheet1.xml",
          (await zip.file("xl/worksheets/sheet1.xml")!.async("string")).replace(
            `<c r="A1" t="inlineStr"><is><t>Report</t></is></c>`,
            `<c r="A1" t="s"><v>0</v></c>`,
          ),
        );
      }
      const bytes = await zip.generateAsync({ type: "uint8array" });
      const blob = new Blob([bytes.slice()]);
      const failingAt = (refused: number) => {
        let reads = 0;
        return {
          size: blob.size,
          slice: (start: number, end: number) => {
            reads += 1;
            return reads === refused
              ? { arrayBuffer: () => Promise.reject(new Error("NotReadable")) }
              : blob.slice(start, end);
          },
        } as unknown as Blob;
      };

      let total = 0;
      const counted = {
        size: blob.size,
        slice: (start: number, end: number) => {
          total += 1;
          return blob.slice(start, end);
        },
      } as unknown as Blob;
      await inspectFile({ name: "log.xlsx", file: counted }, {});
      expect(total).toBeGreaterThan(2);

      for (let refused = 1; refused <= total; refused += 1) {
        await expect(
          inspectFile({ name: "log.xlsx", file: failingAt(refused) }, {}),
        ).rejects.toMatchObject({ code: "FILE_UNREADABLE" });
      }
    },
  );
});
