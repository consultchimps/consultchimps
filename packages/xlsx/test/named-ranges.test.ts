import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import {
  readWorkbookNamedRanges,
  splitWorkbookByColumn,
} from "../src/index.js";

const temporaryDirectories: string[] = [];

async function createTemporaryDirectory(): Promise<string> {
  const directory = await mkdtemp(
    path.join(tmpdir(), "consultchimps-xlsx-ranges-"),
  );
  temporaryDirectories.push(directory);
  return directory;
}

async function createWorkbookWithNames(
  filePath: string,
  sheets: Array<[string, Array<Array<null | number | string>>]>,
  names: Array<{ Name: string; Ref: string }>,
): Promise<void> {
  const workbook = XLSX.utils.book_new();
  for (const [sheetName, rows] of sheets) {
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(rows),
      sheetName,
    );
  }
  workbook.Workbook = { Names: names };
  await writeFile(
    filePath,
    XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
  );
}

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { force: true, recursive: true })),
  );
});

describe("readWorkbookNamedRanges", () => {
  it("reads bounded named ranges and ignores built-in names", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    await createWorkbookWithNames(
      input,
      [
        [
          "Clients",
          [
            ["Quarterly report", null, null],
            ["Client", "Region", "Amount"],
            ["A", "North", 10],
            ["B", "South", 20],
            ["C", "North", 30],
            ["Total", null, 60],
          ],
        ],
      ],
      [
        { Name: "ClientRange", Ref: "Clients!$A$2:$C$5" },
        { Name: "_xlnm.Print_Area", Ref: "Clients!$A$1:$C$6" },
      ],
    );

    const ranges = await readWorkbookNamedRanges(input);
    expect(ranges).toHaveLength(1);
    expect(ranges[0]).toMatchObject({
      columns: ["Client", "Region", "Amount"],
      rangeName: "ClientRange",
      rangeRef: "A2:C5",
      rows: [
        { Amount: 10, Client: "A", Region: "North" },
        { Amount: 20, Client: "B", Region: "South" },
        { Amount: 30, Client: "C", Region: "North" },
      ],
      source: { file: "clients.xlsx", firstDataRow: 3, sheet: "Clients" },
    });
  });

  it("resolves quoted sheet names in range references", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    await createWorkbookWithNames(
      input,
      [
        [
          "Client Data",
          [
            ["Client", "Region"],
            ["A", "North"],
          ],
        ],
      ],
      [{ Name: "Data", Ref: "'Client Data'!$A$1:$B$2" }],
    );

    const ranges = await readWorkbookNamedRanges(input);
    expect(ranges).toHaveLength(1);
    expect(ranges[0]?.source?.sheet).toBe("Client Data");
  });
});

describe("splitWorkbookByColumn with a named range", () => {
  it("splits the named range and ignores cells outside it", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    const output = path.join(directory, "split");
    await createWorkbookWithNames(
      input,
      [
        [
          "Clients",
          [
            ["Quarterly report", null, null],
            ["Client", "Region", "Amount"],
            ["A", "North", 10],
            ["B", "South", 20],
            ["C", "North", 30],
            ["Total", null, 60],
          ],
        ],
      ],
      [{ Name: "ClientRange", Ref: "Clients!$A$2:$C$5" }],
    );

    const result = await splitWorkbookByColumn({
      input,
      outputDirectory: output,
      column: "Region",
      range: "clientrange",
    });

    expect(result.metrics).toEqual({
      groups: 2,
      inputFiles: 1,
      inputRows: 3,
      outputFiles: 2,
      outputRows: 3,
      skippedRows: 0,
    });

    const north = XLSX.read(
      await readFile(path.join(output, "clients-North.xlsx")),
      { type: "buffer" },
    );
    expect(
      XLSX.utils.sheet_to_json(north.Sheets.Clients!, { raw: true }),
    ).toEqual([
      { Amount: 10, Client: "A", Region: "North" },
      { Amount: 30, Client: "C", Region: "North" },
    ]);
  });

  it("rejects conflicting or unsupported named-range options", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    await createWorkbookWithNames(
      input,
      [
        [
          "Clients",
          [
            ["Client", "Region"],
            ["A", "North"],
          ],
        ],
      ],
      [{ Name: "ClientRange", Ref: "Clients!$A$1:$B$2" }],
    );

    await expect(
      splitWorkbookByColumn({
        input,
        outputDirectory: path.join(directory, "conflict"),
        column: "Region",
        range: "ClientRange",
        table: "ClientData",
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_TABLE_RANGE_CONFLICT" });

    await expect(
      splitWorkbookByColumn({
        input,
        outputDirectory: path.join(directory, "header"),
        column: "Region",
        headerRow: 2,
        range: "ClientRange",
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_RANGE_HEADER_ROW" });

    await expect(
      splitWorkbookByColumn({
        input,
        outputDirectory: path.join(directory, "preserve"),
        column: "Region",
        preserveWorkbook: true,
        range: "ClientRange",
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_PRESERVE_REQUIRES_TABLE" });

    await expect(
      splitWorkbookByColumn({
        input,
        outputDirectory: path.join(directory, "missing"),
        column: "Region",
        range: "MissingRange",
      }),
    ).rejects.toThrowError(
      /Named range "MissingRange" was not found or has no data rows/,
    );
  });
});
