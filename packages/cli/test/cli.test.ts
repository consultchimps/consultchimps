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

interface JsonErrorEnvelope {
  error: { code: string | null; message: string };
  ok: false;
}

interface JsonSuccessEnvelope {
  ok: true;
  result: Record<string, unknown>;
}

// --json promises exactly one JSON object on a single line of stdout, so
// parsing raw stdout and checking it is one line is itself part of the
// contract under test rather than mere test convenience.
function parseSingleJsonLine(stdout: string): unknown {
  expect(stdout.endsWith("\n")).toBe(true);
  const body = stdout.slice(0, -1);
  expect(body).not.toContain("\n");
  return JSON.parse(body);
}

function parseJsonSuccess(stdout: string): Record<string, unknown> {
  const envelope = parseSingleJsonLine(stdout) as JsonSuccessEnvelope;
  expect(envelope.ok).toBe(true);
  return envelope.result;
}

function parseJsonError(stdout: string): JsonErrorEnvelope["error"] {
  const envelope = parseSingleJsonLine(stdout) as JsonErrorEnvelope;
  expect(envelope.ok).toBe(false);
  return envelope.error;
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
  sheets: Array<
    [string, Array<Array<XLSX.CellObject | boolean | null | number | string>>]
  >,
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
    XLSX.write(workbook, {
      bookType: "xlsx",
      cellStyles: true,
      type: "buffer",
    }),
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
  it("shows output options in top-level and command help", async () => {
    const topLevelHelp = await runCli(["--help"]);
    expect(topLevelHelp.stdout).toContain(
      'consultchimps sheets consolidate "inputs/*.xlsx" -o combined.xlsx',
    );
    expect(topLevelHelp.stdout).toContain(
      "consultchimps pdf split report.pdf -o pages",
    );
    expect(topLevelHelp.stdout).toContain("consultchimps pptx populate");
    expect(topLevelHelp.stdout).toContain("--json");
    expect(topLevelHelp.stdout).toContain("Automation with --json:");
    expect(topLevelHelp.stdout).toContain('{"ok":true,"result":...}');
    expect(topLevelHelp.stdout).toContain(
      "consultchimps --json pdf split report.pdf -o pages",
    );

    // --help and --version terminate through Commander's exit override, so they
    // must still succeed rather than be treated as usage failures.
    const version = await runCli(["--version"]);
    expect(version.stdout.trim()).toMatch(/^\d+\.\d+\.\d+/);
    expect(version.stderr).toBe("");

    const sheetsHelp = await runCli(["sheets", "--help"]);
    expect(sheetsHelp.stdout).toContain(
      "consultchimps sheets split clients.xlsx -c Region -o by-region",
    );

    const mergeHelp = await runCli(["sheets", "merge", "--help"]);
    expect(mergeHelp.stdout).toContain("-o, --output <path>");
    expect(mergeHelp.stdout).toContain("--no-index");
    expect(mergeHelp.stdout).toContain("--values");
    expect(mergeHelp.stdout).toContain(
      'consultchimps sheets merge "inputs/*.xlsx" --values -o all-sheets.xlsx',
    );

    const consolidateHelp = await runCli(["sheets", "consolidate", "--help"]);
    expect(consolidateHelp.stdout).toContain("-o, --output <path>");
    expect(consolidateHelp.stdout).toContain("--values");
    expect(consolidateHelp.stdout).toContain("--map <file>");
    expect(consolidateHelp.stdout).toContain("--suggest-map <file>");
    expect(consolidateHelp.stdout).toContain(
      "where to save the new consolidated .xlsx workbook",
    );
    expect(consolidateHelp.stdout).toContain("Examples:");
    expect(consolidateHelp.stdout).toContain("What happens:");
    expect(consolidateHelp.stdout).toContain(
      "Your original workbooks are never changed.",
    );

    const splitHelp = await runCli(["sheets", "split", "--help"]);
    expect(splitHelp.stdout).toContain("--values");
    expect(splitHelp.stdout).toContain("preserving formatting");

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
  });

  it("distinguishes merge from consolidate in sheets help", async () => {
    // Commander hard-wraps help output at ~80 columns, so full sentences are
    // asserted against whitespace-normalized stdout.
    const mergeHelp = await runCli(["sheets", "merge", "--help"]);
    const mergeHelpText = mergeHelp.stdout.replace(/\s+/g, " ");
    expect(mergeHelpText).toContain(
      "copy every worksheet from multiple Excel workbooks into one workbook, keeping each sheet separate",
    );
    expect(mergeHelp.stdout).toContain(
      "When you want one combined sheet instead of separate tabs:",
    );
    expect(mergeHelpText).toContain(
      "Use consultchimps sheets consolidate to stack the rows from every worksheet into a single sheet, matching columns by header.",
    );

    const consolidateHelp = await runCli(["sheets", "consolidate", "--help"]);
    const consolidateHelpText = consolidateHelp.stdout.replace(/\s+/g, " ");
    expect(consolidateHelpText).toContain(
      "stack the rows from every worksheet into one combined sheet, matching columns by header",
    );
    expect(consolidateHelp.stdout).toContain(
      "When you want each worksheet kept as its own tab instead:",
    );
    expect(consolidateHelpText).toContain(
      "Use consultchimps sheets merge to copy every worksheet into one workbook without combining any rows.",
    );

    const sheetsHelp = await runCli(["sheets", "--help"]);
    expect(sheetsHelp.stdout).toContain(
      'consultchimps sheets merge "inputs/*.xlsx" -o all-sheets.xlsx',
    );
  });

  it("wraps --json success and failure in a single-line envelope", async () => {
    const directory = await createTemporaryDirectory();
    const sourcePath = path.join(directory, "inputs", "source.pdf");
    const pagesDirectory = path.join(directory, "outputs", "pages");
    const source = await PDFDocument.create();
    source.addPage([300, 400]);
    source.addPage([400, 500]);
    await writeFile(sourcePath, await source.save());

    const success = await runCli([
      "--json",
      "pdf",
      "split",
      sourcePath,
      "--output",
      pagesDirectory,
    ]);
    // Machine-readable mode must not leak any of the human explanation.
    expect(success.stdout).not.toContain("SUCCESS: ConsultChimps");
    const result = parseJsonSuccess(success.stdout);
    expect(result).toMatchObject({
      metrics: { inputFiles: 1, outputFiles: 2, pages: 2 },
      operation: "pdf.split",
      warnings: [],
    });
    expect(
      (result.artifacts as Array<{ path: string }>).map((artifact) =>
        path.basename(artifact.path),
      ),
    ).toEqual(["source-page-001.pdf", "source-page-002.pdf"]);

    // Re-running collides with the outputs just written, which is a stable,
    // published failure rather than an unexpected crash.
    const failure = await runCli(
      ["--json", "pdf", "split", sourcePath, "--output", pagesDirectory],
      1,
    );
    expect(failure.stdout).not.toContain("ERROR: ConsultChimps");
    const error = parseJsonError(failure.stdout);
    expect(error.code).toBe("FILES_OUTPUT_EXISTS");
    expect(typeof error.message).toBe("string");
    expect(error.message.length).toBeGreaterThan(0);
  });

  it("reports usage errors through the --json envelope", async () => {
    // An unknown option fails inside Commander before any command action runs.
    const unknownOption = await runCli(["--json", "--bogus"], 1);
    const optionError = parseJsonError(unknownOption.stdout);
    expect(optionError.code).toBe("CLI_USAGE");
    expect(optionError.message).toContain("--bogus");

    // A missing required option fails on the subcommand rather than the root
    // program, which only produces an envelope if the subcommand inherited the
    // exit override.
    const missingOption = await runCli(
      ["--json", "sheets", "split", "clients.xlsx"],
      1,
    );
    const missingError = parseJsonError(missingOption.stdout);
    expect(missingError.code).toBe("CLI_USAGE");
    expect(missingError.message).toContain("--column");

    // Without --json, Commander keeps reporting usage errors as prose on
    // stderr with nothing on stdout.
    const humanError = await runCli(["sheets", "split", "clients.xlsx"], 1);
    expect(humanError.stdout).toBe("");
    expect(humanError.stderr).toContain(
      "required option '-c, --column <name>' not specified",
    );
    // Commander quotes the argument it could not parse, and an argument is
    // untrusted: a shell expanding a pattern can hand over a filename that
    // begins with a dash and carries a terminal escape. Commander writes that
    // prose itself, before it throws, so it is escaped on the way out.
    const escapeCharacter = String.fromCharCode(0x1b);
    const craftedArgument = await runCli(
      ["sheets", "inspect", `--x${escapeCharacter}[31mclients.xlsx`],
      1,
    );
    expect(craftedArgument.stderr).not.toContain(escapeCharacter);
    expect(craftedArgument.stderr).toContain(
      "unknown option '--x\\u001B[31mclients.xlsx'",
    );
    // The prose keeps its own line breaks, so it stays readable.
    expect(craftedArgument.stderr.endsWith("\n")).toBe(true);
  });

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
    expect(parseJsonSuccess(inspection.stdout)).toMatchObject({
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
    expect(parseJsonSuccess(populated.stdout)).toMatchObject({
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

  it("describes a workbook without creating or changing anything", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "inputs", "clients.xlsx");
    await copyFile(structuredTableFixture, input);
    const before = await readFile(input);

    const help = await runCli(["sheets", "inspect", "--help"]);
    expect(help.stdout).toContain("--sheet <name>");
    expect(help.stdout).toContain("--header-row <number>");
    expect(help.stdout).toContain("--hidden");
    expect(help.stdout).toContain("--samples <number>");
    expect(help.stdout).toContain("consultchimps sheets inspect clients.xlsx");
    expect(help.stdout.replace(/\s+/g, " ")).toContain(
      "No file is created and nothing in the workbook is changed.",
    );

    const command = await runCli(["sheets", "inspect", input]);
    expect(command.stdout).toContain("Excel workbook inspection: clients.xlsx");
    // The header row this fixture resolves without help is its report title,
    // which is exactly the mistake an inspection exists to expose before a
    // consolidation matches columns on it.
    expect(command.stdout).toContain("2. Clients (visible)");
    expect(command.stdout).toContain("Used range: 8 rows by 7 columns");
    expect(command.stdout).toContain("Header row: 1");
    expect(command.stdout).toContain("Data rows below the header: 7");
    expect(command.stdout).toContain(
      "1. ClientData on worksheet Clients (B4:D8)",
    );
    // Each table column name is quoted, so a header spelled "City, State" can
    // never read as two columns.
    expect(command.stdout).toContain('Columns: "Client", "Region", "Amount"');
    expect(command.stdout).toContain("Named ranges:");
    // The description is rendered beside the result the messages package
    // explains, and an inspection reports no artifacts because it writes none.
    expect(command.stdout).toContain(
      "Your Excel workbook inspection is complete.",
    );
    expect(command.stdout).toContain(
      "Nothing was created or changed. An inspection only reads the workbook.",
    );
    expect(command.stdout).toContain("No files were created.");

    // Naming the real header row moves every reported column, and nothing is
    // written either way. The options come first here, as the usage line says
    // they may: a repeatable --sheet takes one name each, so the workbook is
    // still there to be read afterwards.
    const withHeaderRow = await runCli([
      "sheets",
      "inspect",
      "--sheet",
      "Clients",
      "--header-row",
      "4",
      "--samples",
      "2",
      input,
    ]);
    expect(withHeaderRow.stdout).toContain("Header row: 4");
    expect(withHeaderRow.stdout).toContain('2. Client: "A", "B"');
    expect(withHeaderRow.stdout).toContain('3. Region: "North", "South"');
    expect(withHeaderRow.stdout).toContain("4. Amount: 10, 20");

    expect(await readdir(path.join(directory, "inputs"))).toEqual([
      "clients.xlsx",
    ]);
    expect(await readdir(path.join(directory, "outputs"))).toEqual([]);
    expect(await readFile(input)).toEqual(before);
  });

  it("returns the whole inspection outcome through the --json envelope", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "inputs", "review-log.xlsx");
    await writeWorkbook(
      input,
      [
        [
          "Reviews",
          [
            ["Region", "Owner", "Score"],
            ["North", "Ana", 10],
            ["South", "Ben", 20],
            ["North", "Ana", 10],
          ],
        ],
        ["Archive", [["Note"], ["Archived"]]],
      ],
      ["Archive"],
    );

    const command = await runCli([
      "--json",
      "sheets",
      "inspect",
      input,
      "--samples",
      "2",
    ]);
    // --json carries the description as well as the result: the counts alone
    // would drop the sheet names, headers, and samples an inspection is for.
    expect(parseJsonSuccess(command.stdout)).toMatchObject({
      description: {
        excelTables: [],
        namedRanges: [],
        sheets: [
          {
            columnCount: 3,
            columns: [
              { header: "Region", index: 0, sampleValues: ["North", "South"] },
              { header: "Owner", index: 1, sampleValues: ["Ana", "Ben"] },
              // Stored values keep their type, so the number 10 is not "10".
              { header: "Score", index: 2, sampleValues: [10, 20] },
            ],
            dataRowCount: 3,
            headerRow: 1,
            name: "Reviews",
            rowCount: 4,
            visibility: "visible",
          },
        ],
        source: "review-log.xlsx",
      },
      result: {
        artifacts: [],
        metrics: {
          dataRows: 3,
          excelTables: 0,
          headerColumns: 3,
          hiddenWorksheets: 0,
          namedRanges: 0,
          worksheets: 1,
        },
        operation: "sheets.inspect",
        warnings: [
          "1 worksheet is hidden and was not described. Include hidden worksheets to describe it.",
        ],
      },
    });
    expect(command.stderr).toBe("");

    const withHidden = await runCli([
      "--json",
      "sheets",
      "inspect",
      input,
      "--hidden",
    ]);
    const outcome = parseJsonSuccess(withHidden.stdout) as {
      description: { sheets: Array<{ name: string; visibility: string }> };
      result: { metrics: Record<string, number>; warnings: string[] };
    };
    expect(outcome.description.sheets.map((sheet) => sheet.name)).toEqual([
      "Reviews",
      "Archive",
    ]);
    expect(outcome.description.sheets[1]?.visibility).toBe("hidden");
    expect(outcome.result.metrics.hiddenWorksheets).toBe(1);
    expect(outcome.result.warnings).toEqual([]);
  });

  it("shows control characters from a crafted workbook rather than sending them to the terminal", async () => {
    // Written as code points so this file carries no control character of its
    // own, which would defeat the point of the assertions below. The escape is
    // what a workbook's cells can carry; the C1 control byte is what its
    // worksheet names can, because the strict parser the workbook part is read
    // with rejects a character reference for an escape outright.
    const escapeCharacter = String.fromCharCode(0x1b);
    const c1Character = String.fromCharCode(0x9b);
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "inputs", "crafted.xlsx");
    await writeWorkbook(input, [
      [
        "SheetNAMEMARK",
        [
          ["RegionMARKER", "Owner", "Notes"],
          ["NorthMARKER", "Ana", 'a "quoted" value \\ backslash'],
        ],
      ],
      // An empty worksheet earns the "no header row" warning, which quotes the
      // worksheet name back at the reader.
      ["EmptyNAMEMARK", [[null]]],
    ]);

    // Excel's own writer stores a control character as the literal text
    // `_x001b_`, so the reachable path is a hand-built package. Patching the
    // markers into one produces exactly that workbook without committing a
    // binary: the worksheet part carries the header and the values, and the
    // workbook part the worksheet name, which the CLI also narrates as
    // progress on stderr.
    const zip = await JSZip.loadAsync(await readFile(input));
    const sheetXml = await zip
      .file("xl/worksheets/sheet1.xml")!
      .async("string");
    expect(sheetXml).toContain("MARKER");
    zip.file(
      "xl/worksheets/sheet1.xml",
      sheetXml.replaceAll("MARKER", "&#27;[31m"),
    );
    const workbookXml = await zip.file("xl/workbook.xml")!.async("string");
    expect(workbookXml).toContain("NAMEMARK");
    zip.file(
      "xl/workbook.xml",
      workbookXml.replaceAll("NAMEMARK", `${c1Character}[31m`),
    );
    await writeFile(
      input,
      await zip.generateAsync({ compression: "DEFLATE", type: "nodebuffer" }),
    );

    const command = await runCli(["sheets", "inspect", input]);
    // The report is written to a terminal, where a control character is an
    // instruction: it must arrive as visible text instead. Progress narration
    // carries the worksheet name too, so it is held to the same rule.
    expect(command.stdout).not.toContain(escapeCharacter);
    expect(command.stdout).not.toContain(c1Character);
    expect(command.stderr).not.toContain(c1Character);
    expect(command.stdout).toContain("1. Sheet\\u009B[31m (visible)");
    expect(command.stderr).toContain(
      "Describing worksheets 1/2: Sheet\\u009B[31m",
    );
    expect(command.stdout).toContain("1. Region\\u001B[31m:");
    // The library's warnings quote the worksheet name back at the reader, and
    // the messages package writes them to stdout, so they are escaped as well.
    expect(command.stdout).toContain(
      'No header row was found in "Empty\\u009B[31m"',
    );
    expect(command.stdout).toContain('"North\\u001B[31m"');
    // The quotes are the sample's boundary, so a quote inside the text is
    // escaped and a backslash is doubled: the value must not read as two
    // samples, and an escape must not read as text the workbook holds.
    expect(command.stdout).toContain(
      '3. Notes: "a \\"quoted\\" value \\\\ backslash"',
    );

    // The structured result is data rather than terminal output, and JSON's own
    // escaping already makes a control character inert, so it keeps the values
    // the workbook holds.
    const json = await runCli(["--json", "sheets", "inspect", input]);
    const outcome = parseJsonSuccess(json.stdout) as {
      description: {
        sheets: Array<{ columns: Array<{ header: string }>; name: string }>;
      };
    };
    expect(outcome.description.sheets[0]?.name).toBe(`Sheet${c1Character}[31m`);
    expect(outcome.description.sheets[0]?.columns[0]?.header).toBe(
      `Region${escapeCharacter}[31m`,
    );
  });

  it("surfaces the library's refusals when inspecting a workbook", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "inputs", "review-log.xlsx");
    await writeWorkbook(input, [
      [
        "Reviews",
        [
          ["Region", "Owner"],
          ["North", "Ana"],
        ],
      ],
    ]);

    // The sample bound belongs to the library, which refuses an out-of-range
    // request with a stable code rather than quietly clamping it.
    const tooManySamples = await runCli(
      ["--json", "sheets", "inspect", input, "--samples", "50"],
      1,
    );
    expect(parseJsonError(tooManySamples.stdout)).toEqual({
      code: "XLSX_INVALID_SAMPLE_LIMIT",
      message: "The sample value count must be a whole number from 0 to 5.",
    });

    const missingSheet = await runCli(
      ["--json", "sheets", "inspect", input, "--sheet", "Nowhere"],
      1,
    );
    expect(parseJsonError(missingSheet.stdout).code).toBe(
      "XLSX_WORKSHEET_NOT_FOUND",
    );

    // A header row the library would reject must reach it rather than being
    // rounded down here, which would describe a row nobody asked for.
    // Blank text is the one value Number reads as a number, so it is refused
    // rather than quietly meaning "no samples".
    const blankSamples = await runCli(
      ["--json", "sheets", "inspect", input, "--samples", ""],
      1,
    );
    expect(parseJsonError(blankSamples.stdout).code).toBe(
      "XLSX_INVALID_SAMPLE_LIMIT",
    );

    for (const headerRow of ["1.5", "3junk", "0"]) {
      const rejected = await runCli(
        ["--json", "sheets", "inspect", input, "--header-row", headerRow],
        1,
      );
      expect(parseJsonError(rejected.stdout)).toEqual({
        code: "XLSX_INVALID_HEADER_ROW",
        message:
          "The header row must be a positive whole number counted from 1.",
      });
    }

    const missingFile = await runCli(
      ["sheets", "inspect", path.join(directory, "inputs", "absent.xlsx")],
      1,
    );
    expect(missingFile.stderr).toContain(
      "ERROR: ConsultChimps could not finish your task.",
    );
    expect(missingFile.stderr).toContain("FILES_NOT_FOUND");
    expect(missingFile.stdout).toBe("");
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
    expect(parseJsonSuccess(command.stdout).metrics).toMatchObject({
      inputFiles: 2,
      inputTables: 3,
      outputColumns: 7,
      outputRows: 4,
    });
    // --json promises machine-readable output only: no progress narration may
    // reach stderr on success.
    expect(command.stderr).toBe("");

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

    // Without --json the CLI narrates progress on stderr while it works; the
    // non-TTY test process receives it as plain newline-delimited lines.
    const force = await runCli([
      "sheets",
      "consolidate",
      path.join(inputs, "*.xlsx"),
      "--output",
      output,
      "--force",
    ]);
    expect(force.stderr).toContain("Reading workbooks 1/2: north.xlsx");
    expect(force.stderr).toContain("Reading workbooks 2/2: south.xlsx");
    expect(force.stderr).toContain("Writing output 1/1: consolidated.xlsx");
  });

  it("applies a column mapping and drafts one through the built command", async () => {
    const directory = await createTemporaryDirectory();
    const inputs = path.join(directory, "inputs");
    await writeWorkbook(path.join(inputs, "north.xlsx"), [
      [
        "Cases",
        [
          ["Case ID", "Failed Checks", "Region"],
          ["R-1", 5, "north"],
        ],
      ],
    ]);
    await writeWorkbook(path.join(inputs, "south.xlsx"), [
      [
        "Cases",
        [
          ["Reference", "Failed_Checks", "Region"],
          ["R-2", 7, "south"],
        ],
      ],
    ]);

    // --suggest-map drafts the equivalence groups and still consolidates; the
    // draft is evidence for a second run, never applied for the user.
    const draft = path.join(directory, "drafts", "mapping.json");
    const suggested = await runCli([
      "sheets",
      "consolidate",
      path.join(inputs, "*.xlsx"),
      "--output",
      path.join(directory, "suggested.xlsx"),
      "--suggest-map",
      draft,
    ]);
    expect(suggested.stdout).toContain(
      "drafted a column mapping proposing 1 canonical column",
    );
    expect(suggested.stdout).toContain("Type: Column mapping file");
    expect(suggested.stdout).toContain(
      "Review and edit the drafted column mapping",
    );
    expect(suggested.stderr).toContain(
      "Writing mapping draft 1/1: mapping.json",
    );
    expect(JSON.parse(await readFile(draft, "utf8"))).toEqual({
      version: 1,
      columns: [{ name: "Failed Checks", aliases: [] }],
    });

    // The never-overwrite rule covers the draft exactly as it covers a
    // workbook: a second run refuses rather than replacing it.
    const refused = await runCli(
      [
        "--json",
        "sheets",
        "consolidate",
        path.join(inputs, "*.xlsx"),
        "--output",
        path.join(directory, "again.xlsx"),
        "--suggest-map",
        draft,
      ],
      1,
    );
    expect(parseJsonError(refused.stdout).code).toBe("FILES_OUTPUT_EXISTS");

    const mapping = path.join(directory, "mapping.json");
    await writeFile(
      mapping,
      JSON.stringify({
        version: 1,
        columns: [
          { name: "Case_ID", aliases: ["Case ID", "Reference"] },
          { name: "Failed Checks", aliases: [] },
        ],
      }),
      "utf8",
    );
    const output = path.join(directory, "mapped.xlsx");
    const mapped = await runCli([
      "sheets",
      "consolidate",
      path.join(inputs, "*.xlsx"),
      "--output",
      output,
      "--map",
      mapping,
      "--no-source",
    ]);
    expect(mapped.stdout).toContain(
      "1 column did not match the column mapping and kept its own name",
    );
    expect(mapped.stdout).toContain('"Region"');

    const workbook = XLSX.read(await readFile(output), { type: "buffer" });
    expect(
      XLSX.utils.sheet_to_json(workbook.Sheets.Consolidated!, {
        defval: null,
        header: 1,
        raw: true,
      }),
    ).toEqual([
      ["Case_ID", "Failed Checks", "Region"],
      ["R-1", 5, "north"],
      ["R-2", 7, "south"],
    ]);

    // The two options describe different reviews of the same headers, and
    // ADR 0002 fixed no meaning for combining them, so the run refuses.
    const combined = await runCli(
      [
        "--json",
        "sheets",
        "consolidate",
        path.join(inputs, "*.xlsx"),
        "--output",
        path.join(directory, "both.xlsx"),
        "--map",
        mapping,
        "--suggest-map",
        path.join(directory, "second-draft.json"),
      ],
      1,
    );
    expect(parseJsonError(combined.stdout).code).toBe(
      "XLSX_MAPPING_SUGGEST_CONFLICT",
    );

    const badMapping = path.join(directory, "broken.json");
    await writeFile(badMapping, "{ not JSON", "utf8");
    const broken = await runCli(
      [
        "sheets",
        "consolidate",
        path.join(inputs, "*.xlsx"),
        "--output",
        path.join(directory, "broken.xlsx"),
        "--map",
        badMapping,
      ],
      1,
    );
    expect(broken.stderr).toContain("XLSX_MAPPING_FILE_INVALID");
    expect(broken.stderr).toContain(
      "Nothing was created or changed: a column mapping is checked and applied before any output is written.",
    );
  });

  it("merges worksheets through the built command", async () => {
    const directory = await createTemporaryDirectory();
    const inputs = path.join(directory, "inputs");
    const output = path.join(directory, "outputs", "merged.xlsx");
    await writeWorkbook(
      path.join(inputs, "north.xlsx"),
      [
        [
          "Summary",
          [
            ["Amount", "Tax", "Total"],
            [100, 5, { f: "A2+B2", t: "n", v: 105, z: "$#,##0.00" }],
          ],
        ],
        ["Private", [["Amount"], [100]]],
      ],
      ["Private"],
    );
    await writeWorkbook(path.join(inputs, "south.xlsx"), [
      ["Summary", [["Region"], ["South"]]],
    ]);

    const command = await runCli([
      "--json",
      "sheets",
      "merge",
      path.join(inputs, "*.xlsx"),
      "--output",
      output,
    ]);

    expect(parseJsonSuccess(command.stdout).metrics).toMatchObject({
      hiddenSheets: 1,
      inputFiles: 2,
      outputSheets: 3,
    });
    // --json suppresses the stderr progress narration entirely on success.
    expect(command.stderr).toBe("");
    const workbook = XLSX.read(await readFile(output), { type: "buffer" });
    expect(workbook.SheetNames).toEqual([
      "Summary",
      "Private",
      "Summary (2)",
      "Sheet Index",
    ]);
    expect(workbook.Sheets.Summary?.C2?.f).toBe("A2+B2");

    const humanOutput = path.join(directory, "outputs", "human.xlsx");
    const human = await runCli([
      "sheets",
      "merge",
      path.join(inputs, "*.xlsx"),
      "--output",
      humanOutput,
      "--no-index",
      "--values",
    ]);
    expect(human.stdout).toContain("Your Excel workbook merge is complete.");
    // Human mode narrates each merged input and the final write on stderr.
    expect(human.stderr).toContain("Merging inputs 1/2: north.xlsx");
    expect(human.stderr).toContain("Merging inputs 2/2: south.xlsx");
    expect(human.stderr).toContain("Writing output 1/1: human.xlsx");
    expect(human.stdout).toContain("3 worksheets");
    expect(human.stdout).toContain(
      "1 source worksheet was hidden in the merged workbook.",
    );
    expect(human.stdout).not.toContain('see the visible "Sheet Index"');
    const valuesWorkbook = XLSX.read(await readFile(humanOutput), {
      cellStyles: true,
      type: "buffer",
    });
    expect(valuesWorkbook.Sheets.Summary?.C2).toMatchObject({
      v: 105,
      z: "$#,##0.00",
    });
    expect(valuesWorkbook.Sheets.Summary?.C2?.f).toBeUndefined();
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
      "--output-dir",
      output,
      "--prefix",
      "client-region",
    ]);
    expect(parseJsonSuccess(command.stdout).metrics).toMatchObject({
      groups: 2,
      inputFiles: 1,
      inputRows: 4,
      outputFiles: 2,
      outputRows: 3,
      sheetsFiltered: 1,
      skippedRows: 1,
    });
    expect((await readdir(output)).sort()).toEqual([
      "client-region-North.xlsx",
      "client-region-South.xlsx",
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

    expect(parseJsonSuccess(command.stdout).metrics).toMatchObject({
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

    expect(parseJsonSuccess(command.stdout).metrics).toMatchObject({
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

  it("splits a named range through the built command", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "inputs", "clients.xlsx");
    const output = path.join(directory, "outputs", "range-regions");
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([
        ["Quarterly report", null],
        ["Client", "Region"],
        ["A", "North"],
        ["B", "South"],
      ]),
      "Clients",
    );
    workbook.Workbook = {
      Names: [{ Name: "ClientRange", Ref: "Clients!$A$2:$B$4" }],
    };
    await writeFile(
      input,
      XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
    );

    const command = await runCli([
      "--json",
      "sheets",
      "split",
      input,
      "--range",
      "ClientRange",
      "--column",
      "Region",
      "--output",
      output,
    ]);

    expect(parseJsonSuccess(command.stdout).metrics).toMatchObject({
      groups: 2,
      inputRows: 2,
      outputFiles: 2,
    });
    expect((await readdir(output)).sort()).toEqual([
      "clients-North.xlsx",
      "clients-South.xlsx",
    ]);
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
    expect(parseJsonSuccess(split.stdout).metrics).toEqual({
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
    expect(parseJsonSuccess(merge.stdout).metrics).toEqual({
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
