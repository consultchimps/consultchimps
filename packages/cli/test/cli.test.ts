import { execFile } from "node:child_process";
import {
  copyFile,
  mkdtemp,
  mkdir,
  readFile,
  readdir,
  rm,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import JSZip from "jszip";
import { PDFDocument } from "pdf-lib";
import { afterEach, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import { readWorkbookExcelTables } from "@consultchimps/xlsx";

const execFileAsync = promisify(execFile);
const cliPath = fileURLToPath(new URL("../dist/index.js", import.meta.url));
const structuredTableFixture = fileURLToPath(
  new URL("../../xlsx/test/fixtures/structured-table.xlsx", import.meta.url),
);
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
  sheets: Array<[string, Array<Array<boolean | null | number | string>>]>,
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

async function writePowerPointTemplate(filePath: string): Promise<void> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
  );
  zip.file(
    "ppt/presentation.xml",
    '<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>',
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>',
  );
  zip.file(
    "ppt/slides/slide1.xml",
    '<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Profile"/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:rPr b="1"/><a:t>{{client_name}}: {{revenue}}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
  );
  await writeFile(
    filePath,
    await zip.generateAsync({ compression: "DEFLATE", type: "nodebuffer" }),
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
  it(
    "shows output options in top-level and command help",
    { timeout: 30_000 },
    async () => {
      const topLevelHelp = await runCli(["--help"]);
      expect(topLevelHelp.stdout).toContain(
        'consultchimps sheets consolidate "inputs/*.xlsx" -o combined.xlsx',
      );
      expect(topLevelHelp.stdout).toContain(
        "consultchimps pdf split report.pdf -o pages",
      );
      expect(topLevelHelp.stdout).toContain("consultchimps pptx populate");

      const sheetsHelp = await runCli(["sheets", "--help"]);
      expect(sheetsHelp.stdout).toContain(
        "consultchimps sheets split clients.xlsx -c Region -o by-region",
      );

      const consolidateHelp = await runCli(["sheets", "consolidate", "--help"]);
      expect(consolidateHelp.stdout).toContain("-o, --output <path>");
      expect(consolidateHelp.stdout).toContain(
        "where to save the new consolidated .xlsx workbook",
      );
      expect(consolidateHelp.stdout).toContain("Examples:");
      expect(consolidateHelp.stdout).toContain("What happens:");
      expect(consolidateHelp.stdout).toContain(
        "Your original workbooks are never changed.",
      );

      const pdfSplitHelp = await runCli(["pdf", "split", "--help"]);
      expect(pdfSplitHelp.stdout).toContain("-o, --output <directory>");
      expect(pdfSplitHelp.stdout).toContain(
        "folder where the separate page files will be saved",
      );
      expect(pdfSplitHelp.stdout).toContain(
        "consultchimps pdf split report.pdf -o pages",
      );
      expect(pdfSplitHelp.stdout).toContain("What happens:");

      const pptxHelp = await runCli(["pptx", "--help"]);
      expect(pptxHelp.stdout).toContain("pptx inspect-template");
      expect(pptxHelp.stdout).toContain("pptx populate");

      const pptxPopulateHelp = await runCli(["pptx", "populate", "--help"]);
      expect(pptxPopulateHelp.stdout).toContain("--template <path>");
      expect(pptxPopulateHelp.stdout).toContain("--data <path>");
      expect(pptxPopulateHelp.stdout).toContain("{{field_name}}");
      expect(pptxPopulateHelp.stdout).toContain(
        "Source files are never changed.",
      );
    },
  );

  it("inspects and populates a PowerPoint template through the built command", async () => {
    const directory = await createTemporaryDirectory();
    const template = path.join(directory, "inputs", "profile.pptx");
    const data = path.join(directory, "inputs", "companies.xlsx");
    const output = path.join(directory, "outputs", "profiles.pptx");
    await writePowerPointTemplate(template);
    await writeWorkbook(data, [
      [
        "Companies",
        [
          ["client_name", "revenue"],
          ["Company A", "$12.4M"],
          ["Company B", "$8.7M"],
        ],
      ],
    ]);

    const inspection = await runCli([
      "--json",
      "pptx",
      "inspect-template",
      template,
    ]);
    expect(JSON.parse(inspection.stdout)).toMatchObject({
      placeholderOccurrences: 2,
      placeholders: [
        { name: "client_name", occurrences: 1 },
        { name: "revenue", occurrences: 1 },
      ],
      slideNumber: 1,
    });

    const populated = await runCli([
      "--json",
      "pptx",
      "populate",
      "--template",
      template,
      "--data",
      data,
      "--output",
      output,
    ]);
    expect(JSON.parse(populated.stdout)).toMatchObject({
      metrics: {
        generatedSlides: 2,
        replacements: 4,
      },
      operation: "pptx.populate",
    });
    const outputZip = await JSZip.loadAsync(await readFile(output));
    const generatedText = await Promise.all(
      ["ppt/slides/slide2.xml", "ppt/slides/slide3.xml"].map((entry) =>
        outputZip.file(entry)!.async("string"),
      ),
    );
    expect(generatedText[0]).toContain("Company A: $12.4M");
    expect(generatedText[1]).toContain("Company B: $8.7M");
  });

  it("explains successful tasks and recoverable errors in plain language", async () => {
    const directory = await createTemporaryDirectory();
    const sourcePath = path.join(directory, "inputs", "source.pdf");
    const pagesDirectory = path.join(directory, "outputs", "pages");
    const source = await PDFDocument.create();
    source.addPage([300, 400]);
    source.addPage([400, 500]);
    await writeFile(sourcePath, await source.save());

    const success = await runCli([
      "pdf",
      "split",
      sourcePath,
      "--output",
      pagesDirectory,
    ]);
    expect(success.stdout).toContain(
      "SUCCESS: ConsultChimps finished your task.",
    );
    expect(success.stdout).toContain("Your PDF split is complete.");
    expect(success.stdout).toContain(
      "ConsultChimps read 2 pages from the source PDF.",
    );
    expect(success.stdout).toContain("Your original PDF was not changed.");
    expect(success.stdout).toContain("Files created:");
    expect(success.stdout).toContain("What you can do next:");

    const error = await runCli(
      ["pdf", "split", sourcePath, "--output", pagesDirectory],
      1,
    );
    expect(error.stderr).toContain(
      "ERROR: ConsultChimps could not finish your task.",
    );
    expect(error.stderr).toContain("Choose a different output filename");
    expect(error.stderr).toContain(
      "If you intentionally want to replace the existing output",
    );
    expect(error.stderr).toContain("FILES_OUTPUT_EXISTS");
  });

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

  it("splits one workbook into files grouped by a column", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "inputs", "clients.xlsx");
    const output = path.join(directory, "outputs", "regions");
    await writeWorkbook(input, [
      [
        "Clients",
        [
          ["Client", "Region", "Amount"],
          ["A", "North", 10],
          ["B", "South", 20],
          ["C", "North", 30],
          ["D", null, 40],
        ],
      ],
    ]);

    const command = await runCli([
      "--json",
      "sheets",
      "split",
      input,
      "--column",
      "Region",
      "--output",
      output,
      "--prefix",
      "client-region",
    ]);
    expect(JSON.parse(command.stdout).metrics).toEqual({
      groups: 3,
      inputFiles: 1,
      inputRows: 4,
      outputFiles: 3,
      outputRows: 4,
      skippedRows: 0,
    });
    expect((await readdir(output)).sort()).toEqual([
      "client-region-North.xlsx",
      "client-region-South.xlsx",
      "client-region-blank.xlsx",
    ]);

    const north = XLSX.read(
      await readFile(path.join(output, "client-region-North.xlsx")),
      { type: "buffer" },
    );
    expect(
      XLSX.utils.sheet_to_json(north.Sheets.Clients!, {
        defval: null,
        raw: true,
      }),
    ).toEqual([
      { Amount: 10, Client: "A", Region: "North" },
      { Amount: 30, Client: "C", Region: "North" },
    ]);

    const overwrite = await runCli(
      [
        "--json",
        "sheets",
        "split",
        input,
        "--column",
        "Region",
        "--output",
        output,
        "--prefix",
        "client-region",
      ],
      1,
    );
    expect(overwrite.stderr).toContain("FILES_OUTPUT_EXISTS");
    await expect(
      runCli([
        "--json",
        "sheets",
        "split",
        input,
        "--column",
        "Region",
        "--output",
        output,
        "--prefix",
        "client-region",
        "--force",
      ]),
    ).resolves.toMatchObject({ exitCode: 0 });
  });

  it("splits a named Excel Table through the built command", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "inputs", "clients.xlsx");
    const output = path.join(directory, "outputs", "regions");
    await copyFile(structuredTableFixture, input);

    const command = await runCli([
      "--json",
      "sheets",
      "split",
      input,
      "--table",
      "ClientData",
      "--column",
      "Region",
      "--no-preserve-workbook",
      "--output",
      output,
    ]);

    expect(JSON.parse(command.stdout).metrics).toEqual({
      groups: 2,
      inputFiles: 1,
      inputRows: 3,
      outputFiles: 2,
      outputRows: 3,
      skippedRows: 0,
    });
    expect((await readdir(output)).sort()).toEqual([
      "clients-North.xlsx",
      "clients-South.xlsx",
    ]);

    const north = XLSX.read(
      await readFile(path.join(output, "clients-North.xlsx")),
      { type: "buffer" },
    );
    expect(
      XLSX.utils.sheet_to_json(north.Sheets.Clients!, {
        defval: null,
        raw: true,
      }),
    ).toEqual([
      { Amount: 10, Client: "A", Region: "North" },
      { Amount: 30, Client: "C", Region: "North" },
    ]);
  });

  it("preserves the complete workbook by default while splitting an Excel Table", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "inputs", "clients.xlsx");
    const output = path.join(directory, "outputs", "preserved-regions");
    await copyFile(structuredTableFixture, input);

    const command = await runCli([
      "--json",
      "sheets",
      "split",
      input,
      "--table",
      "ClientData",
      "--column",
      "Region",
      "--output",
      output,
    ]);

    expect(JSON.parse(command.stdout).metrics).toMatchObject({
      groups: 2,
      outputFiles: 2,
      outputRows: 3,
    });
    const northPath = path.join(output, "clients-North.xlsx");
    expect(await readWorkbookExcelTables(northPath)).toMatchObject([
      {
        excelTableRange: "B4:D7",
        rows: [
          { Amount: 10, Client: "A", Region: "North" },
          { Amount: 30, Client: "C", Region: "North" },
        ],
      },
    ]);
    const northWorkbook = XLSX.read(await readFile(northPath), {
      type: "buffer",
    });
    expect(northWorkbook.SheetNames).toEqual(["Cover", "Clients"]);
    expect(northWorkbook.Sheets.Clients?.G4?.v).toBe(
      "Cells outside ClientData",
    );
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
