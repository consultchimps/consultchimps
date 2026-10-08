/**
 * Inspection reads each worksheet once, as a stream (ADR 0006): never the
 * whole workbook at once, and never more of a worksheet's rows than the header
 * rule still needs.
 */
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import type { RandomAccessSource } from "@consultchimps/core";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { blobSource, describeWorkbookBytes } from "../src/bytes.js";
import { describeWorkbook } from "../src/index.js";
import { StreamedWorkbook } from "../src/operations/consolidate/reader.js";
import { describeStreamedWorkbook } from "../src/operations/describe.js";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MIB = 1024 * 1024;

/** A one-sheet workbook, stored rather than compressed. */
async function workbookBytes(sheetData: string): Promise<Uint8Array> {
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
    `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<worksheet xmlns="${MAIN}"><sheetData>${sheetData}</sheetData></worksheet>`,
  );
  return zip.generateAsync({ type: "uint8array" });
}

function row(number: number, values: readonly string[]): string {
  const cells = values
    .map((value, column) =>
      value === ""
        ? ""
        : `<c r="${String.fromCharCode(65 + column)}${number}" t="inlineStr"><is><t>${value}</t></is></c>`,
    )
    .join("");
  return `<row r="${number}">${cells}</row>`;
}

/** A source that refuses any read over 1 MiB, and records the largest. */
function guardedSource(
  name: string,
  bytes: Uint8Array,
): { source: RandomAccessSource; largestRead: () => number } {
  const inner = blobSource(name, new Blob([bytes.slice()]));
  let largest = 0;
  return {
    largestRead: () => largest,
    source: {
      name,
      size: inner.size,
      readAt: (offset, length, signal) => {
        largest = Math.max(largest, length);
        if (length > MIB) {
          throw new Error(`Read ${length} bytes of ${inner.size} at once`);
        }
        return inner.readAt(offset, length, signal);
      },
    },
  };
}

describe("describing a workbook read in pieces", () => {
  it("never reads more than 1 MiB at once, and answers as the file surface does", async () => {
    const rows = [row(1, ["Case_ID", "Region", "Note"])];
    for (let number = 2; number <= 8_000; number += 1) {
      rows.push(
        row(number, [
          `R-${number}`,
          number % 2 ? "North" : "South",
          `Note for row ${number}`,
        ]),
      );
    }
    const bytes = await workbookBytes(rows.join(""));
    expect(bytes.length).toBeGreaterThan(MIB);

    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));
    try {
      const file = path.join(directory, "large.xlsx");
      await writeFile(file, bytes);
      const reference = await describeWorkbook(file);

      const guarded = guardedSource("large.xlsx", bytes);
      const workbook = await StreamedWorkbook.open(guarded.source, {
        file: "large.xlsx",
        source: "large.xlsx",
        details: { source: "large.xlsx" },
      });
      const outcome = await describeStreamedWorkbook(
        workbook,
        "large.xlsx",
        {},
        "memory",
      );

      expect(guarded.largestRead()).toBeLessThanOrEqual(MIB);
      expect(outcome).toEqual(reference);
      expect(outcome.description.sheets[0]).toMatchObject({
        dataRowCount: 7_999,
        headerRow: 1,
        rowCount: 8_000,
      });
      expect(outcome.description.sheets[0]!.columns[0]!.sampleValues).toEqual([
        "R-2",
        "R-3",
        "R-4",
        "R-5",
        "R-6",
      ]);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  }, 60_000);

  it("samples in row order across the rows held for the header rule and the rows after", async () => {
    // A one-name header over three-column rows reads as the header, since its
    // first record sits right below it; that is only settled a dozen rows
    // down, so the samples come partly from rows held until then.
    const rows = [row(1, ["Name", "", ""])];
    for (let number = 2; number <= 40; number += 1) {
      const name =
        number <= 12 ? (number % 2 ? "a" : "b") : "cdef"[number % 4]!;
      rows.push(row(number, [name, number <= 20 ? "x" : "y", String(number)]));
    }
    const { description, result } = await describeWorkbookBytes({
      name: "late.xlsx",
      bytes: await workbookBytes(rows.join("")),
    });

    const sheet = description.sheets[0]!;
    expect(sheet.headerRow).toBe(1);
    expect(sheet.dataRowCount).toBe(39);
    expect(sheet.columns.map((column) => column.header)).toEqual([
      "Name",
      "column_2",
      "column_3",
    ]);
    expect(sheet.columns[0]!.sampleValues).toEqual(["b", "a", "d", "e", "f"]);
    expect(sheet.columns[1]!.sampleValues).toEqual(["x", "y"]);
    expect(sheet.columns[2]!.sampleValues).toEqual(["2", "3", "4", "5", "6"]);
    expect(result.metrics.dataRows).toBe(39);
  });
});
