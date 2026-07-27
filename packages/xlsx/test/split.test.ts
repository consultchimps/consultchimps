import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import { splitWorkbookByColumn } from "../src/index.js";

const temporaryDirectories: string[] = [];

async function createTemporaryDirectory(): Promise<string> {
  const directory = await mkdtemp(
    path.join(tmpdir(), "consultchimps-xlsx-split-"),
  );
  temporaryDirectories.push(directory);
  return directory;
}

async function createWorkbook(
  filePath: string,
  sheets: Array<[string, Array<Array<boolean | null | number | string>>]>,
): Promise<void> {
  const workbook = XLSX.utils.book_new();
  for (const [sheetName, rows] of sheets) {
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(rows),
      sheetName,
    );
  }
  await writeFile(
    filePath,
    XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
  );
}

async function readRows(
  filePath: string,
  sheetName: string,
): Promise<Array<Record<string, unknown>>> {
  const workbook = XLSX.read(await readFile(filePath), { type: "buffer" });
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]!, {
    defval: null,
    raw: true,
  });
}

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { force: true, recursive: true })),
  );
});

describe("splitWorkbookByColumn", () => {
  it("writes one workbook per typed value with safe, stable filenames", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    const output = path.join(directory, "split");
    await createWorkbook(input, [
      [
        "Clients",
        [
          ["Client", "Region", "Amount"],
          ["A", "North", 10],
          ["B", "South", 20],
          ["C", "North", 30],
          ["D", null, 40],
          ["E", "North/West", 50],
          ["F", "North:West", 60],
          ["G", 1, 70],
          ["H", "1", 80],
        ],
      ],
    ]);

    const result = await splitWorkbookByColumn(input, output, {
      column: " region ",
    });

    expect(result.metrics).toEqual({
      groups: 7,
      inputFiles: 1,
      inputRows: 8,
      outputFiles: 7,
      outputRows: 8,
      skippedRows: 0,
    });
    expect(
      result.artifacts.map((artifact) => path.basename(artifact.path)),
    ).toEqual([
      "clients-North.xlsx",
      "clients-South.xlsx",
      "clients-blank.xlsx",
      "clients-North-West.xlsx",
      "clients-North-West-2.xlsx",
      "clients-1.xlsx",
      "clients-1-2.xlsx",
    ]);
    expect(
      await readRows(path.join(output, "clients-North.xlsx"), "Clients"),
    ).toEqual([
      { Amount: 10, Client: "A", Region: "North" },
      { Amount: 30, Client: "C", Region: "North" },
    ]);
    expect(
      await readRows(path.join(output, "clients-blank.xlsx"), "Clients"),
    ).toEqual([{ Amount: 40, Client: "D", Region: null }]);
    expect(
      await readRows(path.join(output, "clients-1.xlsx"), "Clients"),
    ).toEqual([{ Amount: 70, Client: "G", Region: 1 }]);
    expect(
      await readRows(path.join(output, "clients-1-2.xlsx"), "Clients"),
    ).toEqual([{ Amount: 80, Client: "H", Region: "1" }]);
  });

  it("requires an explicit worksheet when more than one table is available", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "regions.xlsx");
    await createWorkbook(input, [
      [
        "Current",
        [
          ["Client", "Region"],
          ["A", "North"],
        ],
      ],
      [
        "Archive",
        [
          ["Client", "Region"],
          ["B", "South"],
        ],
      ],
    ]);

    await expect(
      splitWorkbookByColumn(input, path.join(directory, "ambiguous"), {
        column: "Region",
      }),
    ).rejects.toThrowError(/multiple non-empty worksheets/);

    const result = await splitWorkbookByColumn(
      input,
      path.join(directory, "selected"),
      {
        column: "Region",
        sheet: "archive",
      },
    );
    expect(result.artifacts).toHaveLength(1);
    expect(await readRows(result.artifacts[0]!.path, "Archive")).toEqual([
      { Client: "B", Region: "South" },
    ]);
  });

  it("preflights every output before writing and can skip blank values", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    const output = path.join(directory, "split");
    await createWorkbook(input, [
      [
        "Clients",
        [
          ["Client", "Region"],
          ["A", "North"],
          ["B", null],
          ["C", "South"],
        ],
      ],
    ]);

    await mkdir(output);
    await createWorkbook(path.join(output, "clients-South.xlsx"), [
      [
        "Existing",
        [
          ["Do", "Not"],
          ["Replace", "Me"],
        ],
      ],
    ]);

    await expect(
      splitWorkbookByColumn(input, output, {
        column: "Region",
        includeBlank: false,
      }),
    ).rejects.toThrowError(/Output already exists/);
    await expect(
      readFile(path.join(output, "clients-North.xlsx")),
    ).rejects.toThrow();

    const result = await splitWorkbookByColumn(input, output, {
      column: "Region",
      includeBlank: false,
      overwrite: true,
    });
    expect(result.metrics.skippedRows).toBe(1);
    expect(result.warnings).toEqual([
      'Skipped 1 row with blank values in "Region".',
    ]);
  });
});
