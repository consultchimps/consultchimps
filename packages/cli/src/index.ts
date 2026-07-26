#!/usr/bin/env node

import path from "node:path";

import {
  isConsultChimpsError,
  type OperationResult,
} from "@consultchimps/core";
import { discoverFiles } from "@consultchimps/files";
import { mergePdfs, splitPdf } from "@consultchimps/pdf";
import { consolidateWorkbooks } from "@consultchimps/xlsx";
import { Command } from "commander";

interface GlobalOptions {
  json?: boolean;
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

  process.stdout.write(
    [
      `${result.operation} completed.`,
      ...Object.entries(result.metrics).map(
        ([name, value]) => `  ${name}: ${value}`,
      ),
      ...result.artifacts.map((artifact) => `  output: ${artifact.path}`),
      ...result.warnings.map((warning) => `  warning: ${warning}`),
      "",
    ].join("\n"),
  );
}

const program = new Command();
program
  .name("consultchimps")
  .description("Composable, local-first operations tools for consultants.")
  .version("0.1.0")
  .option("--json", "print machine-readable JSON");

const sheets = program
  .command("sheets")
  .description("work with spreadsheet files");

sheets
  .command("consolidate")
  .description("combine visible, non-empty worksheets into one Excel table")
  .argument("<inputs...>", "files, directories, or glob patterns")
  .requiredOption("-o, --output <path>", "output .xlsx file")
  .option("--sheet <names...>", "include only these worksheet names")
  .option("--header-row <number>", "one-based header row", positiveInteger)
  .option("--hidden", "include hidden worksheets")
  .option("--no-source", "omit source file, sheet, and row columns")
  .option("--output-sheet <name>", "output worksheet name", "Consolidated")
  .option("-f, --force", "replace an existing output file")
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

const pdf = program.command("pdf").description("work with PDF files");

pdf
  .command("split")
  .description("write each PDF page to a separate file")
  .argument("<input>", "input PDF file")
  .option("-o, --output <directory>", "output directory")
  .option("--prefix <name>", "output filename prefix")
  .option("-f, --force", "replace existing output files")
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
  .description("combine PDF files in resolved input order")
  .argument("<inputs...>", "files, directories, or glob patterns")
  .requiredOption("-o, --output <path>", "output PDF file")
  .option("-f, --force", "replace an existing output file")
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
    const details = program.opts<GlobalOptions>().json
      ? `\n${JSON.stringify({ code: error.code, details: error.details }, null, 2)}`
      : "";
    process.stderr.write(`consultchimps: ${error.message}${details}\n`);
  } else {
    const message = error instanceof Error ? error.message : String(error);
    process.stderr.write(`consultchimps: ${message}\n`);
  }
  process.exitCode = 1;
}
