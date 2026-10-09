/**
 * CSV inputs to consolidation and inspection (ADR 0007): a CSV file is one
 * worksheet named after the file, read through the same header rule and the
 * same writer as a workbook, on both surfaces.
 */
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { ConsultChimpsError } from "@consultchimps/core";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import {
  blobSource,
  consolidateWorkbookSources,
  consolidateWorkbooksBytes,
  describeWorkbookBytes,
  describeWorkbookSource,
  suggestColumnMappingSources,
} from "../src/bytes.js";
import { consolidateWorkbooks, describeWorkbook } from "../src/index.js";
import {
  assertFitsWorksheet,
  type ConsolidationPlan,
} from "../src/operations/consolidate/consolidate.js";
import { sheetRows, sheetXml } from "./support/read-workbook.js";
import { buildWorkbookFixture } from "./support/workbook-fixture.js";

const utf8 = (text: string): Uint8Array<ArrayBuffer> =>
  new TextEncoder().encode(text);

async function workbook(): Promise<Uint8Array> {
  return buildWorkbookFixture({
    sheets: [
      {
        name: "North",
        rows: [
          ["Region", "Amount"],
          ["North", 10],
        ],
      },
    ],
  });
}

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

describe("consolidating CSV files", () => {
  it("stacks a CSV file with a workbook, naming its worksheet after the file", async () => {
    const outcome = await consolidateWorkbooksBytes({
      inputs: [
        { name: "north.xlsx", bytes: await workbook() },
        {
          name: "south sales.csv",
          bytes: utf8("Region,Amount\r\nSouth,7\r\n"),
        },
      ],
    });
    const rows = await sheetRows(outcome.outputs[0]!.bytes, "Consolidated");
    expect(rows).toEqual([
      ["Region", "Amount", "_source_file", "_source_sheet", "_source_row"],
      ["North", 10, "north.xlsx", "North", 2],
      ["South", "7", "south sales.csv", "south sales", 2],
    ]);
    expect(outcome.result.metrics).toMatchObject({
      csvInputFiles: 1,
      inputFiles: 2,
      inputTables: 2,
      outputRows: 2,
    });
  });

  it("reads numbers and dates only when asked, and writes dates as Excel dates", async () => {
    const input = {
      name: "orders.csv",
      bytes: utf8("Order;Amount;Placed\n0042;1.234,50;31/01/2025\n"),
    };
    const plain = await consolidateWorkbooksBytes({ inputs: [input] });
    expect(
      (await sheetRows(plain.outputs[0]!.bytes, "Consolidated"))[1]?.slice(
        0,
        3,
      ),
    ).toEqual(["0042", "1.234,50", "31/01/2025"]);

    const typed = await consolidateWorkbooksBytes({
      inputs: [input],
      csv: { numbers: true, decimalSeparator: ",", dates: "dmy" },
    });
    const bytes = typed.outputs[0]!.bytes;
    expect((await sheetRows(bytes, "Consolidated"))[1]?.slice(0, 3)).toEqual([
      "0042",
      1234.5,
      "2025-01-31T00:00:00.000Z",
    ]);
    // A real date: a serial number with a date style, not text.
    expect(await sheetXml(bytes, "Consolidated")).toMatch(
      /<c r="C2" s="\d+"><v>45688<\/v><\/c>/u,
    );
  });

  it("applies the shared header rule to a CSV file as to a worksheet", async () => {
    const grid = [
      ["Quarterly report", "", "", ""],
      ["", "", "", ""],
      ["Region", "Units", "", "Amount"],
      ["East", "1", "", "3"],
      ["West", "2", "", "4"],
    ];
    const csv = await consolidateWorkbooksBytes({
      inputs: [
        {
          name: "Report.csv",
          bytes: utf8(grid.map((row) => row.join(",")).join("\n")),
        },
      ],
    });
    const book = await consolidateWorkbooksBytes({
      inputs: [
        {
          name: "Report.xlsx",
          bytes: await buildWorkbookFixture({
            sheets: [
              {
                name: "Report",
                rows: grid.map((row) =>
                  row.map((value) => (value === "" ? null : value)),
                ),
              },
            ],
          }),
        },
      ],
    });
    expect(csv.result.metrics).toEqual({
      ...book.result.metrics,
      csvInputFiles: 1,
    });
    expect(csv.result.metrics.skippedSpacerColumns).toBe(1);
    const rows = async (bytes: Uint8Array): Promise<unknown[][]> =>
      (await sheetRows(bytes, "Consolidated")).map((row) => [
        ...row.slice(0, -3),
        row.at(-1),
      ]);
    expect(await rows(csv.outputs[0]!.bytes)).toEqual(
      await rows(book.outputs[0]!.bytes),
    );
  });

  it("reports a CSV file read as Windows-1252", async () => {
    const outcome = await consolidateWorkbooksBytes({
      inputs: [
        { name: "old.csv", bytes: Uint8Array.from([0x61, 0x0a, 0xe9, 0x0a]) },
      ],
    });
    expect(outcome.result.warnings).toContain(
      "old.csv is not valid UTF-8 and has no byte order mark, so it was read as Windows-1252. If its accented letters look wrong, choose its encoding and run again.",
    );
  });

  it("refuses unusable CSV options before reading anything", async () => {
    let reads = 0;
    const source = blobSource("a.csv", new Blob([utf8("a\n1\n")]));
    const error = await failure(() =>
      consolidateWorkbookSources({
        inputs: [
          {
            ...source,
            readAt: (offset, length) => {
              reads += 1;
              return source.readAt(offset, length);
            },
          },
        ],
        csv: { delimiter: ";;" },
        output: {
          write: () => undefined,
          flush: () => Promise.resolve(),
          abort: () => Promise.resolve(),
        },
      }),
    );
    expect(error.code).toBe("XLSX_CSV_INVALID_OPTION");
    expect(reads).toBe(0);
  });

  it("refuses a table no worksheet can hold, before writing a byte", async () => {
    let writes = 0;
    const tall = utf8(`n\n${"1\n".repeat(1_048_576)}`);
    const error = await failure(() =>
      consolidateWorkbookSources({
        inputs: [blobSource("tall.csv", new Blob([tall]))],
        output: {
          write: () => {
            writes += 1;
          },
          flush: () => Promise.resolve(),
          abort: () => Promise.resolve(),
        },
      }),
    );
    expect(error.code).toBe("XLSX_OUTPUT_TOO_LARGE");
    expect(error.details).toEqual({ rows: 1_048_577, columns: 4 });
    expect(writes).toBe(0);

    // A draft needs no worksheet, so it is still offered.
    const suggestion = await suggestColumnMappingSources({
      inputs: [blobSource("tall.csv", new Blob([tall]))],
    });
    expect(suggestion?.mapping.columns).toEqual([]);
  });

  it("holds exactly what a worksheet holds", () => {
    const plan = (rowCount: number, columns: number) =>
      ({
        rowCount,
        columns: new Array<string>(columns).fill("c"),
      }) as unknown as ConsolidationPlan;
    expect(() => {
      assertFitsWorksheet(plan(1_048_575, 16_384));
    }).not.toThrow();
    expect(() => {
      assertFitsWorksheet(plan(1_048_576, 1));
    }).toThrow(/1,048,577 rows/u);
    expect(() => {
      assertFitsWorksheet(plan(1, 16_385));
    }).toThrow(/16,385 columns/u);
  });
});

describe("CSV files on the command line surface", () => {
  let folder: string;

  beforeEach(async () => {
    folder = await mkdtemp(path.join(tmpdir(), "csv-consolidate-"));
  });

  afterEach(async () => {
    await rm(folder, { recursive: true, force: true });
  });

  it("writes the bytes the browser surface writes", async () => {
    const csv = utf8("Region,Amount\nWest,4\n");
    const book = await workbook();
    await writeFile(path.join(folder, "west.csv"), csv);
    await writeFile(path.join(folder, "north.xlsx"), book);
    const output = path.join(folder, "out.xlsx");
    const result = await consolidateWorkbooks({
      inputs: [path.join(folder, "north.xlsx"), path.join(folder, "west.csv")],
      output,
      csv: { numbers: true },
    });
    expect(result.metrics.csvInputFiles).toBe(1);

    const chunks: Uint8Array[] = [];
    await consolidateWorkbookSources({
      inputs: [
        blobSource("north.xlsx", new Blob([book as Uint8Array<ArrayBuffer>])),
        blobSource("west.csv", new Blob([csv])),
      ],
      csv: { numbers: true },
      output: {
        write: (chunk) => {
          chunks.push(chunk);
        },
        flush: () => Promise.resolve(),
        abort: () => Promise.resolve(),
      },
    });
    expect(Buffer.concat(chunks).equals(await readFile(output))).toBe(true);
  });

  it("describes a CSV file the way the byte surface does", async () => {
    const csv = utf8("Region;Amount\nEast;1,5\n");
    await writeFile(path.join(folder, "east.csv"), csv);
    const fromFile = await describeWorkbook(path.join(folder, "east.csv"), {
      csv: { numbers: true, decimalSeparator: "," },
    });
    const fromBlob = await describeWorkbookSource(
      blobSource("east.csv", new Blob([csv])),
      { csv: { numbers: true, decimalSeparator: "," } },
    );
    expect(fromBlob).toEqual(fromFile);
    expect(fromFile.description).toMatchObject({
      source: "east.csv",
      csv: { encoding: "utf-8", delimiter: ";" },
      excelTables: [],
      namedRanges: [],
      sheets: [
        {
          name: "east",
          visibility: "visible",
          rowCount: 2,
          columnCount: 2,
          columns: [
            { header: "Region", index: 0, sampleValues: ["East"] },
            { header: "Amount", index: 1, sampleValues: [1.5] },
          ],
        },
      ],
    });
  });
});

describe("inspecting CSV files", () => {
  it("reports how the file was read, and its warnings", async () => {
    const outcome = await describeWorkbookBytes({
      name: "old.csv",
      bytes: Uint8Array.from([0x61, 0x09, 0x62, 0x0a, 0xe9, 0x09, 0x31, 0x0a]),
    });
    expect(outcome.description.csv).toEqual({
      encoding: "windows-1252",
      delimiter: "\t",
    });
    expect(outcome.description.sheets[0]?.columns).toEqual([
      { header: "a", index: 0, sampleValues: ["é"] },
      { header: "b", index: 1, sampleValues: ["1"] },
    ]);
    expect(outcome.result.warnings).toHaveLength(1);
  });

  it("refuses unusable CSV options whatever the input is", async () => {
    const error = await failure(async () =>
      describeWorkbookBytes(
        { name: "north.xlsx", bytes: await workbook() },
        { csv: { decimalSeparator: "," } },
      ),
    );
    expect(error.code).toBe("XLSX_CSV_INVALID_OPTION");
  });

  it("describes a workbook without a csv entry", async () => {
    const outcome = await describeWorkbookBytes({
      name: "north.xlsx",
      bytes: await workbook(),
    });
    expect(outcome.description).not.toHaveProperty("csv");
  });
});
