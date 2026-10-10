/**
 * CSV output (ADR 0007): what a CSV file holds, and which operations write
 * one, on both surfaces with the same bytes.
 */
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { ConsultChimpsError } from "@consultchimps/core";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import {
  consolidateWorkbooksBytes,
  mergeWorkbooksBytes,
  planSplitWorkbookBytes,
  splitWorkbookBytes,
} from "../src/bytes.js";
import { CsvWorkbook } from "../src/csv/reader.js";
import { CsvTableWriter, guardedCsvText } from "../src/csv/writer.js";
import { bytesSource } from "../src/package/index.js";
import {
  consolidateWorkbooks,
  mergeWorkbooks,
  splitWorkbookByColumn,
} from "../src/index.js";
import { CellDate } from "../src/package/cell-date.js";
import { CellError } from "../src/package/cell-error.js";
import type { WritableCellValue } from "../src/package/table-writer.js";
import { buildWorkbookFixture } from "./support/workbook-fixture.js";

const utf8 = (text: string): Uint8Array<ArrayBuffer> =>
  new TextEncoder().encode(text);
// The byte order mark is kept, so a test sees it.
const text = (bytes: Uint8Array): string =>
  new TextDecoder("utf-8", { ignoreBOM: true }).decode(bytes);
const BOM = "\ufeff";

function written(
  columns: string[],
  rows: WritableCellValue[][],
  options: { bom?: boolean; header?: boolean } = {},
): string {
  const chunks: Uint8Array[] = [];
  const writer = new CsvTableWriter({
    columns,
    ...options,
    onChunk: (chunk) => chunks.push(chunk),
  });
  for (const row of rows) writer.writeRow(row);
  writer.finish();
  return new TextDecoder("utf-8", { ignoreBOM: true }).decode(
    Buffer.concat(chunks),
  );
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

describe("the CSV writer", () => {
  it("quotes as RFC 4180 asks, ends every row with CRLF and starts with a BOM", () => {
    expect(
      written(
        ["Name", "Note"],
        [
          ["a,b", 'say "hi"'],
          ["two\nlines", " padded "],
          ["plain", null],
        ],
      ),
    ).toBe(
      `${BOM}Name,Note\r\n"a,b","say ""hi"""\r\n"two\nlines"," padded "\r\nplain,\r\n`,
    );
  });

  it("leaves the BOM out when asked", () => {
    expect(written(["a"], [["1"]], { bom: false })).toBe("a\r\n1\r\n");
  });

  it("writes dates as ISO, numbers, booleans and error cells as text", () => {
    expect(
      written(
        ["d", "t", "ms", "n", "b", "e"],
        [
          [
            new CellDate(45688, false, "2025-01-31T00:00:00.000Z"),
            new CellDate(45688.5, true, "2025-01-31T12:00:00.000Z"),
            new CellDate(45688.5, true, "2025-01-31T12:00:00.250Z"),
            -1234.5,
            true,
            new CellError("#N/A"),
          ],
        ],
        { bom: false },
      ).split("\r\n")[1],
    ).toBe(
      "2025-01-31,2025-01-31T12:00:00,2025-01-31T12:00:00.250,-1234.5,TRUE,#N/A",
    );
  });

  it("guards text a spreadsheet would run as a formula", () => {
    expect(guardedCsvText("=SUM(A1:A2)")).toBe("'=SUM(A1:A2)");
    expect(guardedCsvText("+cmd")).toBe("'+cmd");
    expect(guardedCsvText("-2+3")).toBe("'-2+3");
    expect(guardedCsvText("@risk")).toBe("'@risk");
    expect(guardedCsvText("\tx")).toBe("'\tx");
    expect(guardedCsvText("-12.5")).toBe("-12.5");
    expect(guardedCsvText("+1,234.50")).toBe("+1,234.50");
    expect(guardedCsvText("-")).toBe("-");
    expect(guardedCsvText("+00123")).toBe("'+00123");
    expect(guardedCsvText("-1234567890123456")).toBe("'-1234567890123456");
    expect(guardedCsvText("-0.5")).toBe("-0.5");
    expect(guardedCsvText("+1,2,3")).toBe("'+1,2,3");
    expect(guardedCsvText("-1..2")).toBe("'-1..2");
    expect(guardedCsvText("-1.234,5")).toBe("-1.234,5");
    expect(guardedCsvText("a=b")).toBe("a=b");
    expect(written(["=x"], [["=1+1"]], { bom: false })).toBe(
      "'=x\r\n'=1+1\r\n",
    );
  });
});

describe("the CSV writer at scale", () => {
  it("joins batches and reads back to the same values", async () => {
    const rows: WritableCellValue[][] = [];
    for (let row = 0; row < 2_500; row += 1) {
      rows.push([
        `r\r\n${row}`,
        `a,b "${row}"`,
        row % 3 === 0 ? "" : `line\nbreak`,
      ]);
    }
    const csv = written(["x", "y", "z"], rows);
    const book = await CsvWorkbook.open(
      bytesSource("back.csv", new TextEncoder().encode(csv)),
      { file: "back.csv", source: "back.csv", details: {} },
    );
    const back: unknown[][] = [];
    await book.readWorksheet(book.sheets[0]!, {
      begin: () => undefined,
      row: (row, cells) => {
        const dense: unknown[] = [null, null, null];
        for (const cell of cells) dense[cell.column] = cell.value;
        back[row] = dense;
      },
    });
    expect(back.length).toBe(2_501);
    expect(back.slice(1)).toEqual(
      rows.map((row) => row.map((value) => (value === "" ? null : value))),
    );
  });

  it("quotes a lone empty field and dates a workbook cannot count", () => {
    expect(written(["a"], [[""], ["x"]], { bom: false })).toBe(
      `a\r\n""\r\nx\r\n`,
    );
    expect(written(["d"], [["1850-01-01T00:00:00.000Z"]], { bom: false })).toBe(
      `d\r\n1850-01-01\r\n`,
    );
    expect(guardedCsvText(" =HYPERLINK()")).toBe("' =HYPERLINK()");
    expect(guardedCsvText(" -12")).toBe(" -12");
  });
});

describe("the CSV writer's row count", () => {
  it("refuses fewer or more rows than declared", () => {
    const writer = () =>
      new CsvTableWriter({
        columns: ["a"],
        rowCount: 1,
        onChunk: () => undefined,
      });
    const short = writer();
    expect(() => short.finish()).toThrow(/0 rows were written but 1/u);
    const long = writer();
    long.writeRow(["x"]);
    expect(() => long.writeRow(["y"])).toThrow(/More rows/u);
  });
});

describe("consolidating to CSV", () => {
  const north = utf8("Region;Amount;Placed\nNorth;1,5;31/01/2025\n");

  it("writes CSV when the output is named .csv, or when asked", async () => {
    const byName = await consolidateWorkbooksBytes({
      inputs: [{ name: "north.csv", bytes: north }],
      outputName: "all.csv",
      csv: { numbers: true, decimalSeparator: ",", dates: "dmy" },
    });
    expect(byName.outputs[0]!.name).toBe("all.csv");
    expect(byName.outputs[0]!.mediaType).toBe("text/csv");
    expect(text(byName.outputs[0]!.bytes)).toBe(
      `${BOM}Region,Amount,Placed,_source_file,_source_sheet,_source_row\r\nNorth,1.5,2025-01-31,north.csv,north,2\r\n`,
    );
    const asked = await consolidateWorkbooksBytes({
      inputs: [{ name: "north.csv", bytes: north }],
      outputFormat: "csv",
      csvBom: false,
    });
    expect(asked.outputs[0]!.name).toBe("consolidated.csv");
    expect(text(asked.outputs[0]!.bytes).startsWith("Region")).toBe(true);
  });

  it("needs no legal worksheet name for a CSV output", async () => {
    const outcome = await consolidateWorkbooksBytes({
      inputs: [{ name: "north.csv", bytes: north }],
      outputFormat: "csv",
      outputSheetName: "History",
    });
    expect(outcome.outputs[0]!.name).toBe("consolidated.csv");
  });

  it("refuses a format that contradicts the output's name", async () => {
    const error = await failure(() =>
      consolidateWorkbooksBytes({
        inputs: [{ name: "north.csv", bytes: north }],
        outputName: "all.xlsx",
        outputFormat: "csv",
      }),
    );
    expect(error.code).toBe("XLSX_OUTPUT_FORMAT_INVALID");
  });

  it("writes CSV beyond a worksheet's limits, which bind workbooks only", async () => {
    const outcome = await consolidateWorkbooksBytes({
      inputs: [
        { name: "notes.csv", bytes: utf8(`Note\n${"x".repeat(40_000)}\n`) },
      ],
      outputFormat: "csv",
      addSourceColumns: false,
    });
    expect(text(outcome.outputs[0]!.bytes)).toContain("x".repeat(40_000));
  });
});

describe("splitting to CSV", () => {
  const orders = utf8("Region,Amount\nNorth,1\nSouth,2\nNorth,3\n");

  it("splits a CSV file into CSV files by default", async () => {
    const outcome = await splitWorkbookBytes({
      input: { name: "orders.csv", bytes: orders },
      column: "Region",
    });
    expect(outcome.outputs.map((output) => output.name)).toEqual([
      "orders-North.csv",
      "orders-South.csv",
    ]);
    expect(outcome.outputs[0]!.mediaType).toBe("text/csv");
    expect(text(outcome.outputs[0]!.bytes)).toBe(
      `${BOM}Region,Amount\r\nNorth,1\r\nNorth,3\r\n`,
    );
  });

  it("writes a CSV split beyond a worksheet's limits", async () => {
    const outcome = await splitWorkbookBytes({
      input: {
        name: "notes.csv",
        bytes: utf8(`Region,Note\nNorth,${"x".repeat(40_000)}\n`),
      },
      column: "Region",
    });
    expect(text(outcome.outputs[0]!.bytes)).toContain("x".repeat(40_000));
  });

  it("splits a workbook into CSV files when asked, and refuses to keep it", async () => {
    const book = await buildWorkbookFixture({
      sheets: [
        {
          name: "Orders",
          rows: [
            ["Region", "Amount"],
            ["North", 1],
            ["South", 2],
          ],
        },
      ],
    });
    const outcome = await splitWorkbookBytes({
      input: { name: "orders.xlsx", bytes: book },
      column: "Region",
      outputFormat: "csv",
      csvBom: false,
    });
    expect(outcome.outputs.map((output) => output.name)).toEqual([
      "orders-North.csv",
      "orders-South.csv",
    ]);
    expect(text(outcome.outputs[1]!.bytes)).toBe(
      "Region,Amount\r\nSouth,2\r\n",
    );
    const error = await failure(() =>
      splitWorkbookBytes({
        input: { name: "orders.xlsx", bytes: book },
        column: "Region",
        outputFormat: "csv",
        preserveWorkbook: true,
      }),
    );
    expect(error.code).toBe("XLSX_SPLIT_CSV_PRESERVE");
  });
});

describe("merging to CSV", () => {
  it("is refused, since a merge's tabs cannot be one CSV file", async () => {
    const error = await failure(async () =>
      mergeWorkbooksBytes({
        inputs: [
          {
            name: "north.xlsx",
            bytes: await buildWorkbookFixture({
              sheets: [{ name: "North", rows: [["a"]] }],
            }),
          },
        ],
        outputName: "all.csv",
      }),
    );
    expect(error.code).toBe("XLSX_OUTPUT_FORMAT_INVALID");
  });
});

describe("CSV output on the command line surface", () => {
  let folder: string;

  beforeEach(async () => {
    folder = await mkdtemp(path.join(tmpdir(), "csv-output-"));
  });

  afterEach(async () => {
    await rm(folder, { recursive: true, force: true });
  });

  it("writes the bytes the byte surface writes", async () => {
    const orders = utf8("Region,Amount\nNorth,1\nSouth,2\n");
    await writeFile(path.join(folder, "orders.csv"), orders);
    const output = path.join(folder, "all.csv");
    const result = await consolidateWorkbooks({
      inputs: [path.join(folder, "orders.csv")],
      output,
    });
    expect(result.artifacts[0]!.mediaType).toBe("text/csv");
    const bytes = await consolidateWorkbooksBytes({
      inputs: [{ name: "orders.csv", bytes: orders }],
      outputName: "all.csv",
    });
    expect(
      Buffer.from(bytes.outputs[0]!.bytes).equals(await readFile(output)),
    ).toBe(true);

    await splitWorkbookByColumn({
      input: path.join(folder, "orders.csv"),
      outputDirectory: path.join(folder, "split"),
      column: "Region",
    });
    const split = await splitWorkbookBytes({
      input: { name: "orders.csv", bytes: orders },
      column: "Region",
    });
    for (const file of split.outputs) {
      expect(
        Buffer.from(file.bytes).equals(
          await readFile(path.join(folder, "split", file.name)),
        ),
      ).toBe(true);
    }
  });

  it("refuses a merge named .csv before writing", async () => {
    await writeFile(path.join(folder, "orders.csv"), utf8("a\n1\n"));
    const error = await failure(() =>
      mergeWorkbooks(
        [path.join(folder, "orders.csv")],
        path.join(folder, "all.csv"),
      ),
    );
    expect(error.code).toBe("XLSX_OUTPUT_FORMAT_INVALID");
  });
});

describe("splitting every worksheet to CSV", () => {
  async function book(): Promise<Uint8Array<ArrayBuffer>> {
    return (await buildWorkbookFixture({
      sheets: [
        {
          name: "Orders",
          rows: [
            ["Region", "Amount"],
            ["North", 1],
            ["South", 2],
          ],
        },
        { name: "Notes", rows: [["Text"], ["no region here"]] },
        {
          name: "Returns",
          rows: [
            ["Region", "Units"],
            ["North", 5],
          ],
          state: "hidden",
        },
      ],
    })) as Uint8Array<ArrayBuffer>;
  }

  it("writes one CSV file per worksheet and value, named for both", async () => {
    const input = { name: "orders.xlsx", bytes: await book() };
    const outcome = await splitWorkbookBytes({
      input,
      column: "Region",
      outputFormat: "csv",
      csvBom: false,
      includeHiddenSheets: true,
    });
    expect(outcome.outputs.map((output) => output.name)).toEqual([
      "orders-North - Orders.csv",
      "orders-South - Orders.csv",
      "orders-North - Returns.csv",
    ]);
    expect(text(outcome.outputs[2]!.bytes)).toBe("Region,Units\r\nNorth,5\r\n");
    expect(outcome.result.metrics).toMatchObject({
      outputFiles: 3,
      groups: 2,
      sheetsFiltered: 2,
      inputRows: 3,
    });
    expect(outcome.result.warnings).toContain(
      'Worksheet "Notes" has no rows under column "Region", so no CSV file was written for it.',
    );

    const plan = await planSplitWorkbookBytes({
      input,
      column: "Region",
      outputFormat: "csv",
      includeHiddenSheets: true,
    });
    expect(plan.outputs.map((output) => output.path)).toEqual(
      outcome.outputs.map((output) => output.name),
    );
  });

  it("leaves a hidden worksheet out unless asked for", async () => {
    const outcome = await splitWorkbookBytes({
      input: { name: "orders.xlsx", bytes: await book() },
      column: "Region",
      outputFormat: "csv",
    });
    expect(outcome.outputs.map((output) => output.name)).toEqual([
      "orders-North.csv",
      "orders-South.csv",
    ]);
  });

  it("keeps one-source names when one worksheet carries the column", async () => {
    const outcome = await splitWorkbookBytes({
      input: {
        name: "orders.xlsx",
        bytes: (await buildWorkbookFixture({
          sheets: [
            { name: "Orders", rows: [["Region"], ["North"]] },
            { name: "Notes", rows: [["Text"], ["x"]] },
          ],
        })) as Uint8Array<ArrayBuffer>,
      },
      column: "Region",
      outputFormat: "csv",
    });
    expect(outcome.outputs.map((output) => output.name)).toEqual([
      "orders-North.csv",
    ]);
  });

  it("tells apart names that come out the same, and leaves out a sheet with no groups", async () => {
    const outcome = await splitWorkbookBytes({
      input: {
        name: "orders.xlsx",
        bytes: (await buildWorkbookFixture({
          sheets: [
            { name: "Orders", rows: [["Region"], ["a/b"], ["a-b"]] },
            {
              name: "Blank",
              rows: [
                ["Region", "N"],
                [null, 1],
              ],
            },
            { name: "Re:turns".replace(":", ""), rows: [["Region"], ["a/b"]] },
          ],
        })) as Uint8Array<ArrayBuffer>,
      },
      column: "Region",
      outputFormat: "csv",
      includeBlank: false,
    });
    expect(outcome.outputs.map((output) => output.name)).toEqual([
      "orders-a-b - Orders.csv",
      "orders-a-b - Orders-2.csv",
      "orders-a-b - Returns.csv",
    ]);
    expect(outcome.result.warnings.join(" ")).toContain('Worksheet "Blank"');
    // The blank sheet's row was read and skipped, so the metrics count it.
    expect(outcome.result.metrics).toMatchObject({
      inputRows: 4,
      skippedRows: 1,
      sheetsFiltered: 3,
    });
  });

  it("refuses for no groups when every sheet with the column is blank", async () => {
    const error = await failure(async () =>
      splitWorkbookBytes({
        input: {
          name: "orders.xlsx",
          bytes: (await buildWorkbookFixture({
            sheets: [
              { name: "Notes", rows: [["Text"], ["x"]] },
              {
                name: "Blank",
                rows: [
                  ["Region", "N"],
                  [null, 1],
                ],
              },
            ],
          })) as Uint8Array<ArrayBuffer>,
        },
        column: "Region",
        outputFormat: "csv",
        includeBlank: false,
      }),
    );
    expect(error.code).toBe("XLSX_SPLIT_NO_GROUPS");
  });

  it("stops before the next worksheet once cancelled", async () => {
    const controller = new AbortController();
    controller.abort();
    const error = await failure(async () =>
      planSplitWorkbookBytes({
        input: { name: "orders.xlsx", bytes: await book() },
        column: "Region",
        outputFormat: "csv",
        signal: controller.signal,
      }),
    );
    expect(error.code).toBe("OPERATION_ABORTED");
  });

  it("refuses a column no worksheet carries", async () => {
    const error = await failure(async () =>
      splitWorkbookBytes({
        input: { name: "orders.xlsx", bytes: await book() },
        column: "Country",
        outputFormat: "csv",
      }),
    );
    expect(error.code).toBe("XLSX_SPLIT_COLUMN_NOT_FOUND");
  });

  it("writes the same files on the command line surface", async () => {
    const folder = await mkdtemp(path.join(tmpdir(), "csv-sheets-"));
    try {
      const bytes = await book();
      await writeFile(path.join(folder, "orders.xlsx"), bytes);
      await splitWorkbookByColumn({
        input: path.join(folder, "orders.xlsx"),
        outputDirectory: path.join(folder, "out"),
        column: "Region",
        outputFormat: "csv",
      });
      const byte = await splitWorkbookBytes({
        input: { name: "orders.xlsx", bytes },
        column: "Region",
        outputFormat: "csv",
      });
      for (const output of byte.outputs) {
        expect(
          Buffer.from(output.bytes).equals(
            await readFile(path.join(folder, "out", output.name)),
          ),
        ).toBe(true);
      }
    } finally {
      await rm(folder, { recursive: true, force: true });
    }
  });
});
