/**
 * CSV inputs to splitting and merging (ADR 0007): a CSV file splits compactly,
 * since it has no workbook to keep, and merges as one worksheet named after
 * the file, on both surfaces with the same bytes.
 */
import { mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { ConsultChimpsError } from "@consultchimps/core";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import {
  blobSource,
  mergeWorkbookSources,
  mergeWorkbooksBytes,
  splitWorkbookBytes,
} from "../src/bytes.js";
import { mergeWorkbooks, splitWorkbookByColumn } from "../src/index.js";
import { sheetNamesOf, sheetRows, sheetXml } from "./support/read-workbook.js";
import { buildWorkbookFixture } from "./support/workbook-fixture.js";

const utf8 = (text: string): Uint8Array<ArrayBuffer> =>
  new TextEncoder().encode(text);

const ORDERS = utf8(
  "Region;Amount;Placed\r\nNorth;1,5;2025-01-31\r\nSouth;2;2025-02-01\r\nNorth;3;2025-02-02\r\n",
);

async function failure(
  run: () => Promise<unknown>,
): Promise<ConsultChimpsError> {
  try {
    await run();
  } catch (error) {
    if (error instanceof ConsultChimpsError) return error;
    throw error;
  }
  throw new Error("The run was expected to fail.");
}

describe("splitting a CSV file", () => {
  it("writes one compact workbook per value, reading values as asked", async () => {
    const outcome = await splitWorkbookBytes({
      input: { name: "orders.csv", bytes: ORDERS },
      column: "Region",
      csv: { numbers: true, decimalSeparator: ",", dates: "iso" },
    });
    expect(outcome.outputs.map((output) => output.name)).toEqual([
      "orders-North.xlsx",
      "orders-South.xlsx",
    ]);
    const north = outcome.outputs[0]!.bytes;
    expect(await sheetNamesOf(north)).toEqual(["orders"]);
    expect(await sheetRows(north, "orders")).toEqual([
      ["Region", "Amount", "Placed"],
      ["North", 1.5, "2025-01-31T00:00:00.000Z"],
      ["North", 3, "2025-02-02T00:00:00.000Z"],
    ]);
    expect(outcome.result.metrics).toMatchObject({
      groups: 2,
      inputRows: 3,
      outputFiles: 2,
    });
  });

  it("keeps every field text by default", async () => {
    const outcome = await splitWorkbookBytes({
      input: { name: "orders.csv", bytes: ORDERS },
      column: "Region",
    });
    expect((await sheetRows(outcome.outputs[1]!.bytes, "orders"))[1]).toEqual([
      "South",
      "2",
      "2025-02-01",
    ]);
  });

  it("refuses to keep a workbook a CSV file does not have", async () => {
    const error = await failure(() =>
      splitWorkbookBytes({
        input: { name: "orders.csv", bytes: ORDERS },
        column: "Region",
        preserveWorkbook: true,
      }),
    );
    expect(error.code).toBe("XLSX_SPLIT_CSV_PRESERVE");
  });
});

describe("merging CSV files", () => {
  it("makes each CSV file one worksheet named after it, the grid as it is", async () => {
    const book = await buildWorkbookFixture({
      sheets: [{ name: "North", rows: [["Region"], ["North"]] }],
    });
    const outcome = await mergeWorkbooksBytes({
      inputs: [
        { name: "north.xlsx", bytes: book },
        {
          name: "south orders.csv",
          bytes: utf8("Report\n\nRegion,Placed\nSouth,2025-01-31\n"),
        },
      ],
      csv: { dates: "iso" },
    });
    const merged = outcome.outputs[0]!.bytes;
    expect(await sheetNamesOf(merged)).toEqual([
      "North",
      "south orders",
      "Sheet Index",
    ]);
    expect(await sheetRows(merged, "south orders")).toEqual([
      ["Report", null],
      [null, null],
      ["Region", "Placed"],
      ["South", "2025-01-31T00:00:00.000Z"],
    ]);
    // A real date, and no filter: the CSV's grid, not a table.
    const xml = await sheetXml(merged, "south orders");
    expect(xml).toMatch(/<c r="B4" s="\d+"><v>45688<\/v><\/c>/u);
    expect(xml).not.toContain("autoFilter");
    expect(outcome.result.metrics).toMatchObject({
      csvInputFiles: 1,
      inputFiles: 2,
      outputSheets: 2,
    });
  });
  it("refuses a field longer than a cell holds, in a merge and a split", async () => {
    const long = utf8(`Region;Note\nNorth;${"x".repeat(32_768)}\n`);
    const merge = await failure(() =>
      mergeWorkbooksBytes({ inputs: [{ name: "notes.csv", bytes: long }] }),
    );
    expect(merge.code).toBe("XLSX_OUTPUT_TOO_LARGE");
    const split = await failure(() =>
      splitWorkbookBytes({
        input: { name: "notes.csv", bytes: long },
        column: "Region",
      }),
    );
    expect(split.code).toBe("XLSX_OUTPUT_TOO_LARGE");
  });

  it("makes an empty CSV file an empty worksheet", async () => {
    for (const bytes of [new Uint8Array(), utf8("\n\n")]) {
      const outcome = await mergeWorkbooksBytes({
        inputs: [{ name: "empty.csv", bytes }],
      });
      const rows = await sheetRows(outcome.outputs[0]!.bytes, "empty");
      expect(rows.flat().every((value) => value === null)).toBe(true);
    }
  });

  it("stops converting a CSV file when cancelled", async () => {
    const controller = new AbortController();
    let reads = 0;
    const source = blobSource(
      "big.csv",
      new Blob([utf8("a\n".repeat(200_000))]),
    );
    const error = await failure(() =>
      mergeWorkbookSources({
        inputs: [
          {
            ...source,
            readAt: (offset, length) => {
              reads += 1;
              if (reads === 10) controller.abort();
              return source.readAt(offset, length);
            },
          },
        ],
        signal: controller.signal,
        output: {
          write: () => undefined,
          flush: () => Promise.resolve(),
          abort: () => Promise.resolve(),
        },
      }),
    );
    expect(error.code).toBe("OPERATION_ABORTED");
    expect(reads).toBeLessThan(20);
  });
});

describe("CSV files on the command line surface", () => {
  let folder: string;

  beforeEach(async () => {
    folder = await mkdtemp(path.join(tmpdir(), "csv-split-merge-"));
  });

  afterEach(async () => {
    await rm(folder, { recursive: true, force: true });
  });

  it("splits to the bytes the byte surface writes", async () => {
    const input = path.join(folder, "orders.csv");
    await writeFile(input, ORDERS);
    const outputDirectory = path.join(folder, "out");
    await splitWorkbookByColumn({
      input,
      outputDirectory,
      column: "Region",
      csv: { numbers: true, decimalSeparator: "," },
    });
    const bytes = await splitWorkbookBytes({
      input: { name: "orders.csv", bytes: ORDERS },
      column: "Region",
      csv: { numbers: true, decimalSeparator: "," },
    });
    expect((await readdir(outputDirectory)).sort()).toEqual([
      "orders-North.xlsx",
      "orders-South.xlsx",
    ]);
    for (const output of bytes.outputs) {
      expect(
        Buffer.from(output.bytes).equals(
          await readFile(path.join(outputDirectory, output.name)),
        ),
      ).toBe(true);
    }
  });

  it("merges to the bytes the byte surface writes", async () => {
    const csv = path.join(folder, "orders.csv");
    await writeFile(csv, ORDERS);
    const output = path.join(folder, "merged.xlsx");
    const second = path.join(folder, "more orders.csv");
    await writeFile(second, ORDERS);
    const result = await mergeWorkbooks([csv, second], output, {
      csv: { numbers: true, decimalSeparator: "," },
    });
    expect(result.metrics.csvInputFiles).toBe(2);
    const bytes = await mergeWorkbooksBytes({
      inputs: [
        { name: "orders.csv", bytes: ORDERS },
        { name: "more orders.csv", bytes: ORDERS },
      ],
      csv: { numbers: true, decimalSeparator: "," },
    });
    expect(
      Buffer.from(bytes.outputs[0]!.bytes).equals(await readFile(output)),
    ).toBe(true);
  });
});
