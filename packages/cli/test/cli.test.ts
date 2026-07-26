import { execFile } from "node:child_process";
import { mkdtemp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { PDFDocument } from "pdf-lib";
import { afterEach, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

const execFileAsync = promisify(execFile);
const cliPath = fileURLToPath(new URL("../dist/index.js", import.meta.url));
const temporaryDirectories: string[] = [];

interface CliResult {
  exitCode: number;
  stderr: string;
  stdout: string;
}

async function createTemporaryDirectory(): Promise<string> {
  const directory = await mkdtemp(
    path.join(tmpdir(), "consultchimps-cli-test-"),
  );
  temporaryDirectories.push(directory);
  await mkdir(path.join(directory, "inputs"));
  await mkdir(path.join(directory, "outputs"));
  return directory;
}

async function runCli(
  args: string[],
  expectedExitCode = 0,
): Promise<CliResult> {
  let outcome: CliResult;

  try {
    const result = await execFileAsync(process.execPath, [cliPath, ...args], {
      encoding: "utf8",
    });
    outcome = { exitCode: 0, stderr: result.stderr, stdout: result.stdout };
  } catch (error) {
    const processError = error as Error & {
      code?: number;
      stderr?: string;
      stdout?: string;
    };
    const exitCode =
      typeof processError.code === "number" ? processError.code : 1;
    outcome = {
      exitCode,
      stderr: processError.stderr ?? "",
      stdout: processError.stdout ?? "",
    };
  }

  expect(
    outcome.exitCode,
    outcome.stderr || `Command exited with ${outcome.exitCode}`,
  ).toBe(expectedExitCode);
  return outcome;
}

async function writeWorkbook(
  filePath: string,
  sheets: Array<[string, Array<Array<string | number>>]>,
  hiddenSheets: string[] = [],
): Promise<void> {
  const workbook = XLSX.utils.book_new();
  for (const [sheetName, rows] of sheets) {
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(rows),
      sheetName,
    );
  }
  workbook.Workbook = {
    Sheets: workbook.SheetNames.map((sheetName) => ({
      Hidden: hiddenSheets.includes(sheetName) ? 1 : 0,
      name: sheetName,
    })),
  };
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

describe("consultchimps CLI", () => {
  it("consolidates workbook globs through the built command", async () => {
    const directory = await createTemporaryDirectory();
    const inputs = path.join(directory, "inputs");
    const output = path.join(directory, "outputs", "consolidated.xlsx");

    await writeWorkbook(
      path.join(inputs, "north.xlsx"),
      [
        [
          "North",
          [
            ["Client", "Amount"],
            ["A", 10],
            ["B", 20],
          ],
        ],
        [
          "Hidden",
          [
            ["Client", "Amount"],
            ["SHOULD_NOT_APPEAR", 999],
          ],
        ],
      ],
      ["Hidden"],
    );
    await writeWorkbook(path.join(inputs, "south.xlsx"), [
      [
        "South",
        [
          ["Amount", "Status", "client"],
          [30, "Open", "C"],
        ],
      ],
      [
        "West",
        [
          ["Client", "Amount", "Region"],
          ["D", 40, "West"],
        ],
      ],
    ]);

    const command = await runCli([
      "--json",
      "sheets",
      "consolidate",
      path.join(inputs, "*.xlsx"),
      "--output",
      output,
    ]);
    expect(JSON.parse(command.stdout).metrics).toEqual({
      inputFiles: 2,
      inputTables: 3,
      outputColumns: 7,
      outputRows: 4,
    });

    const workbook = XLSX.read(await readFile(output), { type: "buffer" });
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(
      workbook.Sheets.Consolidated!,
      { defval: null },
    );
    expect(rows.map((row) => row.Client)).toEqual(["A", "B", "C", "D"]);
    expect(rows.some((row) => row.Client === "SHOULD_NOT_APPEAR")).toBe(false);
    expect(rows[2]).toMatchObject({
      _source_file: "south.xlsx",
      _source_row: 2,
      _source_sheet: "South",
    });
    expect(rows[3]).toMatchObject({ Region: "West" });

    const overwrite = await runCli(
      [
        "--json",
        "sheets",
        "consolidate",
        path.join(inputs, "*.xlsx"),
        "--output",
        output,
      ],
      1,
    );
    expect(overwrite.stderr).toContain("FILES_OUTPUT_EXISTS");
    await expect(
      runCli([
        "--json",
        "sheets",
        "consolidate",
        path.join(inputs, "*.xlsx"),
        "--output",
        output,
        "--force",
      ]),
    ).resolves.toMatchObject({ exitCode: 0 });
  });

  it("splits and merges PDFs through the built command", async () => {
    const directory = await createTemporaryDirectory();
    const inputs = path.join(directory, "inputs");
    const outputs = path.join(directory, "outputs");
    const sourcePath = path.join(inputs, "source.pdf");
    const pagesDirectory = path.join(outputs, "pages");
    const mergedPath = path.join(outputs, "merged.pdf");

    const source = await PDFDocument.create();
    source.addPage([300, 400]);
    source.addPage([400, 500]);
    source.addPage([500, 600]);
    await writeFile(sourcePath, await source.save());

    const split = await runCli([
      "--json",
      "pdf",
      "split",
      sourcePath,
      "--output",
      pagesDirectory,
    ]);
    expect(JSON.parse(split.stdout).metrics).toEqual({
      inputFiles: 1,
      outputFiles: 3,
      pages: 3,
    });

    const expectedSizes = [
      { height: 400, width: 300 },
      { height: 500, width: 400 },
      { height: 600, width: 500 },
    ];
    for (let index = 0; index < expectedSizes.length; index += 1) {
      const filename = `source-page-${String(index + 1).padStart(3, "0")}.pdf`;
      const pageDocument = await PDFDocument.load(
        await readFile(path.join(pagesDirectory, filename)),
      );
      expect(pageDocument.getPageCount()).toBe(1);
      expect(pageDocument.getPage(0).getSize()).toEqual(expectedSizes[index]);
    }

    const merge = await runCli([
      "--json",
      "pdf",
      "merge",
      path.join(pagesDirectory, "*.pdf"),
      "--output",
      mergedPath,
    ]);
    expect(JSON.parse(merge.stdout).metrics).toEqual({
      inputFiles: 3,
      outputFiles: 1,
      pages: 3,
    });

    const merged = await PDFDocument.load(await readFile(mergedPath));
    expect(merged.getPageCount()).toBe(3);
    expect(merged.getPages().map((page) => page.getSize())).toEqual(
      expectedSizes,
    );

    const overwrite = await runCli(
      [
        "--json",
        "pdf",
        "merge",
        path.join(pagesDirectory, "*.pdf"),
        "--output",
        mergedPath,
      ],
      1,
    );
    expect(overwrite.stderr).toContain("FILES_OUTPUT_EXISTS");
    await expect(
      runCli([
        "--json",
        "pdf",
        "merge",
        path.join(pagesDirectory, "*.pdf"),
        "--output",
        mergedPath,
        "--force",
      ]),
    ).resolves.toMatchObject({ exitCode: 0 });
  });
});
