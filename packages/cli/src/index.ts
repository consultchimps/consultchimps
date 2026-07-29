#!/usr/bin/env node

import { readFileSync } from "node:fs";
import path from "node:path";

import {
  isConsultChimpsError,
  type OperationResult,
} from "@consultchimps/core";
import { discoverFiles } from "@consultchimps/files";
import { mergePdfs, splitPdf } from "@consultchimps/pdf";
import {
  consolidateWorkbooks,
  splitWorkbookByColumn,
} from "@consultchimps/xlsx";
import { Command } from "commander";

import { formatHumanError, formatHumanResult } from "./human-output.js";

interface GlobalOptions {
  json?: boolean;
}

interface PackageMetadata {
  version: string;
}

interface ConsolidateOptions {
  force?: boolean;
  headerRow?: number;
  hidden?: boolean;
  output: string;
  outputSheet?: string;
  sheet?: string[];
  source?: boolean;
}

interface SheetSplitOptions {
  column: string;
  force?: boolean;
  headerRow?: number;
  hidden?: boolean;
  output?: string;
  prefix?: string;
  preserveWorkbook?: boolean;
  sheet?: string;
  skipBlank?: boolean;
  table?: string;
}

interface SplitOptions {
  force?: boolean;
  output?: string;
  prefix?: string;
}

interface MergeOptions {
  force?: boolean;
  output: string;
}

function positiveInteger(value: string): number {
  const parsed = Number.parseInt(value, 10);
  if (!Number.isInteger(parsed) || parsed < 1) {
    throw new Error("Expected a positive integer.");
  }
  return parsed;
}

function printResult(result: OperationResult, json: boolean): void {
  if (json) {
    process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
    return;
  }

  process.stdout.write(formatHumanResult(result));
}

const program = new Command();
const packageMetadata = JSON.parse(
  readFileSync(new URL("../package.json", import.meta.url), "utf8"),
) as PackageMetadata;

program
  .name("consultchimps")
  .description(
    "Clear, local-first tools that explain how they process your spreadsheets and PDFs.",
  )
  .version(packageMetadata.version)
  .option(
    "--json",
    "print structured JSON for automation instead of the detailed explanation",
  )
  .addHelpText(
    "after",
    `
Quick start:
  consultchimps sheets consolidate "inputs/*.xlsx" -o combined.xlsx
  consultchimps sheets split clients.xlsx -c Region -o by-region
  consultchimps pdf split report.pdf -o pages
  consultchimps pdf merge "inputs/*.pdf" -o combined.pdf

Run consultchimps help <command> or append --help to a command for all options.
`,
  );

const sheets = program
  .command("sheets")
  .description(
    "combine or divide Excel workbooks without changing the original files",
  )
  .addHelpText(
    "after",
    `
Examples:
  consultchimps sheets consolidate "inputs/*.xlsx" -o combined.xlsx
  consultchimps sheets split clients.xlsx -c Region -o by-region

Safety:
  Your original Excel workbooks are not changed. ConsultChimps creates new
  output files and refuses to replace existing outputs unless you use --force.

Run consultchimps sheets help <command> for all command options.
`,
  );

sheets
  .command("consolidate")
  .description(
    "combine visible, non-empty worksheets into one new Excel workbook",
  )
  .argument(
    "<inputs...>",
    'Excel files, folders, or quoted patterns such as "inputs/*.xlsx"',
  )
  .requiredOption(
    "-o, --output <path>",
    "where to save the new consolidated .xlsx workbook",
  )
  .option(
    "--sheet <names...>",
    "include only worksheets with these exact names",
  )
  .option(
    "--header-row <number>",
    "row containing column names, counted from 1",
    positiveInteger,
  )
  .option("--hidden", "include hidden worksheets as well as visible ones")
  .option(
    "--no-source",
    "leave out columns that identify each row's source file, worksheet, and row",
  )
  .option(
    "--output-sheet <name>",
    "name of the worksheet created in the new workbook",
    "Consolidated",
  )
  .option(
    "-f, --force",
    "replace the output file if it already exists; use with care",
  )
  .addHelpText(
    "after",
    `
Examples:
  consultchimps sheets consolidate "inputs/*.xlsx" -o combined.xlsx
  consultchimps sheets consolidate north.xlsx south.xlsx --output combined.xlsx

What happens:
  1. ConsultChimps finds the matching Excel files.
  2. It reads every selected, non-empty worksheet.
  3. It matches columns by header name and combines all data rows.
  4. It writes one new workbook and explains exactly what was created.

Your original workbooks are never changed.
`,
  )
  .action(async (inputs: string[], options: ConsolidateOptions) => {
    const inputPaths = await discoverFiles(inputs, { extensions: [".xlsx"] });
    const result = await consolidateWorkbooks(inputPaths, options.output, {
      addSourceColumns: options.source !== false,
      headerRow: options.headerRow,
      includeHiddenSheets: options.hidden === true,
      outputSheetName: options.outputSheet,
      overwrite: options.force === true,
      sheets: options.sheet,
    });
    printResult(result, program.opts<GlobalOptions>().json === true);
  });

sheets
  .command("split")
  .description(
    "create one new Excel workbook for each distinct value in a selected column",
  )
  .argument("<input>", "the source .xlsx workbook to divide")
  .requiredOption(
    "-c, --column <name>",
    "column whose values decide which rows go into each new workbook",
  )
  .option(
    "-o, --output <directory>",
    "folder where the new workbooks will be saved",
  )
  .option("--sheet <name>", "exact name of the worksheet to divide")
  .option(
    "--table <name>",
    "use this named Excel Table instead of the worksheet's full used range",
  )
  .option(
    "--header-row <number>",
    "row containing column names, counted from 1",
    positiveInteger,
  )
  .option("--hidden", "allow the selected worksheet to be hidden")
  .option(
    "--preserve-workbook",
    "keep the full workbook layout and replace only the selected Excel Table rows",
  )
  .option(
    "--skip-blank",
    "do not create an output group for rows with a blank split-column value",
  )
  .option(
    "--prefix <name>",
    "text to place at the start of each output filename",
  )
  .option(
    "-f, --force",
    "replace matching output files that already exist; use with care",
  )
  .addHelpText(
    "after",
    `
Examples:
  consultchimps sheets split clients.xlsx -c Region -o by-region
  consultchimps sheets split clients.xlsx --table ClientData --column Region --preserve-workbook

What happens:
  1. ConsultChimps reads the selected worksheet or named Excel Table.
  2. It groups rows using the values in --column.
  3. It creates a clearly named workbook for every distinct group.
  4. It reports every created file, skipped row, and warning.

Your original workbook is never changed.
`,
  )
  .action(async (input: string, options: SheetSplitOptions) => {
    const inputPaths = await discoverFiles([input], {
      extensions: [".xlsx"],
    });
    if (inputPaths.length !== 1) {
      throw new Error(
        `Expected exactly one input workbook; found ${inputPaths.length}.`,
      );
    }

    const inputPath = inputPaths[0];
    if (!inputPath) {
      throw new Error("No input workbook was found.");
    }

    const outputDirectory =
      options.output ??
      path.join(path.dirname(inputPath), `${path.parse(inputPath).name}-split`);
    const result = await splitWorkbookByColumn(inputPath, outputDirectory, {
      column: options.column,
      filenamePrefix: options.prefix,
      headerRow: options.headerRow,
      includeBlank: options.skipBlank !== true,
      includeHiddenSheets: options.hidden === true,
      overwrite: options.force === true,
      preserveWorkbook: options.preserveWorkbook === true,
      sheet: options.sheet,
      table: options.table,
    });
    printResult(result, program.opts<GlobalOptions>().json === true);
  });

const pdf = program
  .command("pdf")
  .description("split or combine PDF documents without changing the originals")
  .addHelpText(
    "after",
    `
Examples:
  consultchimps pdf split report.pdf -o pages
  consultchimps pdf merge "inputs/*.pdf" -o combined.pdf

Safety:
  Your original PDF documents are not changed. ConsultChimps creates new output
  files and refuses to replace existing outputs unless you use --force.

Run consultchimps pdf help <command> for all command options.
`,
  );

pdf
  .command("split")
  .description("create one new PDF file for every page in a source PDF")
  .argument("<input>", "the source PDF document to divide")
  .option(
    "-o, --output <directory>",
    "folder where the separate page files will be saved",
  )
  .option(
    "--prefix <name>",
    "text to place at the start of each output filename",
  )
  .option(
    "-f, --force",
    "replace matching output files that already exist; use with care",
  )
  .addHelpText(
    "after",
    `
Example:
  consultchimps pdf split report.pdf -o pages

What happens:
  ConsultChimps creates one clearly numbered PDF for every page, lists every
  new file, and leaves the source PDF unchanged.
`,
  )
  .action(async (input: string, options: SplitOptions) => {
    const [inputPath] = await discoverFiles([input], { extensions: [".pdf"] });
    if (!inputPath) {
      throw new Error("No input PDF was found.");
    }
    const outputDirectory =
      options.output ??
      path.join(path.dirname(inputPath), `${path.parse(inputPath).name}-pages`);
    const result = await splitPdf(inputPath, outputDirectory, {
      filenamePrefix: options.prefix,
      overwrite: options.force === true,
    });
    printResult(result, program.opts<GlobalOptions>().json === true);
  });

pdf
  .command("merge")
  .description("combine several PDF documents into one new PDF")
  .argument(
    "<inputs...>",
    'PDF files, folders, or quoted patterns such as "inputs/*.pdf"',
  )
  .requiredOption(
    "-o, --output <path>",
    "where to save the new combined PDF document",
  )
  .option(
    "-f, --force",
    "replace the output PDF if it already exists; use with care",
  )
  .addHelpText(
    "after",
    `
Examples:
  consultchimps pdf merge "inputs/*.pdf" -o combined.pdf
  consultchimps pdf merge first.pdf second.pdf --output combined.pdf

What happens:
  ConsultChimps reads the matching PDFs in their resolved order, copies every
  page into one new document, reports the final page count, and leaves every
  source PDF unchanged.
`,
  )
  .action(async (inputs: string[], options: MergeOptions) => {
    const inputPaths = await discoverFiles(inputs, { extensions: [".pdf"] });
    const result = await mergePdfs(inputPaths, options.output, {
      overwrite: options.force === true,
    });
    printResult(result, program.opts<GlobalOptions>().json === true);
  });

try {
  await program.parseAsync(process.argv);
} catch (error) {
  if (isConsultChimpsError(error)) {
    const json = program.opts<GlobalOptions>().json === true;
    const details = json
      ? `\n${JSON.stringify({ code: error.code, details: error.details }, null, 2)}`
      : "";
    process.stderr.write(
      json
        ? `consultchimps: ${error.message}${details}\n`
        : formatHumanError(error.message, error.code),
    );
  } else {
    const message = error instanceof Error ? error.message : String(error);
    process.stderr.write(
      program.opts<GlobalOptions>().json
        ? `consultchimps: ${message}\n`
        : formatHumanError(message),
    );
  }
  process.exitCode = 1;
}
