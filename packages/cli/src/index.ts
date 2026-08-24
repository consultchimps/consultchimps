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
  inspectPowerPointTemplate,
  populatePowerPointTemplate,
} from "@consultchimps/pptx";
import {
  consolidateWorkbooks,
  mergeWorkbooks,
  splitWorkbookByColumn,
} from "@consultchimps/xlsx";
import {
  CLI_VOCABULARY,
  formatHumanError,
  formatHumanResult,
} from "@consultchimps/messages";
import { Command, CommanderError } from "commander";

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
  values?: boolean;
}

interface SheetMergeOptions {
  force?: boolean;
  index: boolean;
  output: string;
  values?: boolean;
}

interface SheetSplitOptions {
  column: string;
  force?: boolean;
  headerRow?: number;
  hidden?: boolean;
  output?: string;
  outputDir?: string;
  prefix?: string;
  preserveWorkbook?: boolean;
  range?: string;
  sheet?: string;
  skipBlank?: boolean;
  strict?: boolean;
  table?: string;
  values?: boolean;
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

interface PptxPopulateOptions {
  data: string;
  force?: boolean;
  headerRow?: number;
  output: string;
  sheet?: string;
  template: string;
  templateSlide?: number;
}

interface PptxInspectOptions {
  templateSlide?: number;
}

function positiveInteger(value: string): number {
  const parsed = Number.parseInt(value, 10);
  if (!Number.isInteger(parsed) || parsed < 1) {
    throw new Error("Expected a positive integer.");
  }
  return parsed;
}

// The --json envelope is a stable machine-readable contract: exactly one JSON
// object on a single line, so a consumer can parse stdout without first
// separating prose from data. "ok" discriminates the two shapes, which lets
// automation branch on success or failure without inspecting the exit code.
function jsonEnvelope(payload: unknown): string {
  return `${JSON.stringify(payload)}\n`;
}

function printJsonResult(result: unknown): void {
  process.stdout.write(jsonEnvelope({ ok: true, result }));
}

function printJsonFailure(message: string, code: string | null): void {
  const envelope = jsonEnvelope({ ok: false, error: { message, code } });
  // stdout stays parseable on its own; stderr repeats the same object so the
  // CLI still reports failures on stderr as every other command does.
  process.stdout.write(envelope);
  process.stderr.write(envelope);
}

function printResult<TMetric extends string>(
  result: OperationResult<TMetric>,
  json: boolean,
): void {
  if (json) {
    printJsonResult(result);
    return;
  }

  process.stdout.write(
    formatHumanResult(result, { vocabulary: CLI_VOCABULARY }),
  );
}

// Usage errors such as an unknown option carry no library error code, so they
// report one stable code of their own rather than null, which is reserved for
// genuinely unexpected failures.
const USAGE_ERROR_CODE = "CLI_USAGE";

// --json has to be read from argv rather than program.opts() because a usage
// error is thrown before Commander finishes populating the parsed options. A
// literal "--json" supplied as an option *value* would false-positive here,
// which is an acceptable trade for making parse failures machine-readable.
const jsonRequested = process.argv.includes("--json");

const program = new Command();

// Commander exits the process itself on a usage error, which would bypass the
// --json envelope and leave stdout empty. exitOverride makes it throw a
// CommanderError so every failure reaches the single handler at the end of this
// file. This and configureOutput must both be set before any .command() call,
// because subcommands copy the exit callback and output configuration from
// their parent at creation time.
program.exitOverride();

if (jsonRequested) {
  // Commander writes its usage prose to stderr before throwing. Silence it so
  // JSON mode emits nothing but the envelope. Help and version text goes
  // through writeOut and is deliberately left alone.
  program.configureOutput({ writeErr: () => {} });
}

const packageMetadata = JSON.parse(
  readFileSync(new URL("../package.json", import.meta.url), "utf8"),
) as PackageMetadata;

program
  .name("consultchimps")
  .description(
    "Clear, local-first tools that explain how they process your spreadsheets, presentations, and PDFs.",
  )
  .version(packageMetadata.version)
  .option(
    "--json",
    "print one line of machine-readable JSON for automation instead of the detailed explanation",
  )
  .addHelpText(
    "after",
    `
Quick start:
  consultchimps sheets consolidate "inputs/*.xlsx" -o combined.xlsx
  consultchimps sheets merge "inputs/*.xlsx" -o all-sheets.xlsx
  consultchimps sheets split clients.xlsx -c Region -o by-region
  consultchimps pptx populate --template profile.pptx --data clients.xlsx --sheet Clients --template-slide 1 -o profiles.pptx
  consultchimps pdf split report.pdf -o pages
  consultchimps pdf merge "inputs/*.pdf" -o combined.pdf

Automation with --json:
  Place --json before the command to replace the explanation with one line of
  JSON on stdout. Success prints {"ok":true,"result":...} and failure prints
  {"ok":false,"error":{"message":...,"code":...}} while keeping the nonzero
  exit code. Nothing else is written to stdout, so the output can be piped
  straight into a JSON parser.

  consultchimps --json pdf split report.pdf -o pages

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
  consultchimps sheets merge "inputs/*.xlsx" -o all-sheets.xlsx
  consultchimps sheets split clients.xlsx -c Region -o by-region

Safety:
  Your original Excel workbooks are not changed. ConsultChimps creates new
  output files and refuses to replace existing outputs unless you use --force.

Run consultchimps sheets help <command> for all command options.
`,
  );

sheets
  .command("merge")
  .description(
    "copy every worksheet from multiple Excel workbooks into one workbook, keeping each sheet separate",
  )
  .argument(
    "<inputs...>",
    'Excel files, folders, or quoted patterns such as "inputs/*.xlsx"',
  )
  .requiredOption("-o, --output <path>", "where to save the new workbook")
  .option("--no-index", "do not add the visible Sheet Index worksheet")
  .option(
    "--values",
    "replace formulas with their stored values while preserving formatting",
  )
  .option(
    "-f, --force",
    "replace the output file if it already exists; use with care",
  )
  .addHelpText(
    "after",
    `
Examples:
  consultchimps sheets merge "inputs/*.xlsx" --values -o all-sheets.xlsx
  consultchimps sheets merge north.xlsx south.xlsx --output all-sheets.xlsx

Every source worksheet remains a separate tab. Sheet Index records source names
and hidden/visible status. Duplicate tab names receive a suffix. --values
removes formulas but always retains cell and workbook formatting.

When you want one combined sheet instead of separate tabs:
  Use consultchimps sheets consolidate to stack the rows from every worksheet
  into a single sheet, matching columns by header.
`,
  )
  .action(async (inputs: string[], options: SheetMergeOptions) => {
    const inputPaths = await discoverFiles(inputs, { extensions: [".xlsx"] });
    const result = await mergeWorkbooks(inputPaths, options.output, {
      includeSheetIndex: options.index,
      overwrite: options.force === true,
      values: options.values === true,
    });
    printResult(result, program.opts<GlobalOptions>().json === true);
  });

sheets
  .command("consolidate")
  .description(
    "stack the rows from every worksheet into one combined sheet, matching columns by header",
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
    "--values",
    "write stored values instead of formulas while preserving output formatting",
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
Consolidation already writes stored values rather than copying formulas;
--values makes that requirement explicit.

When you want each worksheet kept as its own tab instead:
  Use consultchimps sheets merge to copy every worksheet into one workbook
  without combining any rows.
`,
  )
  .action(async (inputs: string[], options: ConsolidateOptions) => {
    const inputPaths = await discoverFiles(inputs, { extensions: [".xlsx"] });
    const result = await consolidateWorkbooks({
      inputs: inputPaths,
      output: options.output,
      addSourceColumns: options.source !== false,
      headerRow: options.headerRow,
      includeHiddenSheets: options.hidden === true,
      outputSheetName: options.outputSheet,
      overwrite: options.force === true,
      sheets: options.sheet,
      values: options.values === true,
    });
    printResult(result, program.opts<GlobalOptions>().json === true);
  });

sheets
  .command("split")
  .description(
    "create one new Excel workbook for each distinct value in a selected column",
  )
  .argument("<input>", "the source .xlsx or .xlsm workbook to divide")
  .requiredOption(
    "-c, --column <name>",
    "column whose values decide which rows go into each new workbook",
  )
  .option(
    "-o, --output <directory>",
    "folder where the new workbooks will be saved",
  )
  .option(
    "--output-dir <directory>",
    "folder where the new workbooks will be saved (alias for --output)",
  )
  .option("--sheet <name>", "exact name of the worksheet to divide")
  .option(
    "--table <name>",
    "use this named Excel Table instead of the worksheet's full used range (preferred)",
  )
  .option(
    "--range <name>",
    "use this named range instead of the worksheet's full used range",
  )
  .option(
    "--header-row <number>",
    "row containing column names, counted from 1",
    positiveInteger,
  )
  .option("--hidden", "allow the selected worksheet to be hidden")
  .option(
    "--preserve-workbook",
    "keep the full workbook layout (default without a selector and for --table)",
  )
  .option(
    "--no-preserve-workbook",
    "write plain data-only workbooks using the single-source split mode",
  )
  .option(
    "--values",
    "replace formulas with their stored values while preserving formatting",
  )
  .option(
    "--strict",
    "match split values exactly, including case, whitespace, and value type",
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
  consultchimps sheets split clients.xlsx -c Region --output-dir by-region
  consultchimps sheets split clients.xlsx --table ClientData --column Region --values
  consultchimps sheets split clients.xlsx --range ClientRange --column Region

What happens:
  1. By default, ConsultChimps finds --column in every worksheet.
  2. It collects distinct non-blank values across all matching worksheets.
  3. It copies the complete workbook once per value and removes other rows.
  4. Worksheets without --column are copied unchanged.

Matching trims surrounding whitespace, ignores case, and treats ordinary
numeric text like the equivalent number. Use --strict for exact matching.
Use --sheet, --table, or --range for the legacy single-source split mode.
Use --no-preserve-workbook only when a compact, data-only result is wanted.

--values removes formulas while retaining their stored results and all
formatting in a preserved workbook. A formula without a stored result becomes
a formatted blank cell and is reported as a warning.

Pro tip: before a table split, prepare the workbook exactly as you want to
deliver it - set each sheet's zoom, place the cursor on cell A1 so every
file opens consistently, add any cover sheet, and save.

Your original workbook is never changed.
`,
  )
  .action(async (input: string, options: SheetSplitOptions) => {
    if (options.output && options.outputDir) {
      throw new Error(
        "Choose either --output or --output-dir; they name the same destination option.",
      );
    }
    const inputPaths = await discoverFiles([input], {
      extensions: [".xlsx", ".xlsm"],
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
      options.outputDir ??
      options.output ??
      path.join(path.dirname(inputPath), `${path.parse(inputPath).name}-split`);
    const result = await splitWorkbookByColumn({
      input: inputPath,
      outputDirectory,
      column: options.column,
      filenamePrefix: options.prefix,
      headerRow: options.headerRow,
      includeBlank: options.skipBlank !== true,
      includeHiddenSheets: options.hidden === true,
      overwrite: options.force === true,
      preserveWorkbook: options.preserveWorkbook,
      range: options.range,
      sheet: options.sheet,
      table: options.table,
      strict: options.strict === true,
      values: options.values === true,
    });
    printResult(result, program.opts<GlobalOptions>().json === true);
  });

const pptx = program
  .command("pptx")
  .description(
    "inspect or populate PowerPoint templates without changing the source files",
  )
  .addHelpText(
    "after",
    `
Examples:
  consultchimps pptx inspect-template profile.pptx
  consultchimps pptx populate --template profile.pptx --data clients.xlsx -o profiles.pptx

Safety:
  Your source PowerPoint template and Excel workbook are not changed.
  ConsultChimps creates one new presentation and refuses to replace an existing
  output unless you use --force.
`,
  );

pptx
  .command("inspect-template")
  .description(
    "list text placeholders on one PowerPoint template slide without creating a file",
  )
  .argument("<template>", "the source .pptx template to inspect")
  .option(
    "--template-slide <number>",
    "template slide number, counted from 1 (default: 1)",
    positiveInteger,
  )
  .addHelpText(
    "after",
    `
Example:
  consultchimps pptx inspect-template profile.pptx

The report identifies valid {{field_name}} placeholders, malformed placeholder
braces, and unsupported placeholder placements. Split-run placeholders are
supported.
`,
  )
  .action(async (template: string, options: PptxInspectOptions) => {
    const inspection = await inspectPowerPointTemplate(template, {
      templateSlide: options.templateSlide,
    });
    if (program.opts<GlobalOptions>().json === true) {
      printJsonResult(inspection);
      return;
    }

    const lines = [
      `PowerPoint template slide ${inspection.slideNumber}`,
      `Placeholder occurrences: ${inspection.placeholderOccurrences}`,
      "Placeholders:",
      ...(inspection.placeholders.length > 0
        ? inspection.placeholders.map(
            (placeholder) =>
              `  - ${placeholder.name}: ${placeholder.occurrences}`,
          )
        : ["  - None"]),
      `Malformed placeholder locations: ${inspection.malformedPlaceholderCount}`,
      `Unsupported split-run placeholders: ${
        inspection.unsupportedSplitRunPlaceholders.join(", ") || "None"
      }`,
      `Unsupported placeholder placements: ${
        inspection.unsupportedPlacementPlaceholders.join(", ") || "None"
      }`,
      "",
    ];
    process.stdout.write(lines.join("\n"));
  });

pptx
  .command("populate")
  .description(
    "create one populated PowerPoint slide for every nonempty Excel data row",
  )
  .requiredOption(
    "--template <path>",
    "source .pptx file containing {{field_name}} placeholders",
  )
  .requiredOption("--data <path>", "source .xlsx workbook containing the data")
  .option(
    "--sheet <name>",
    "exact worksheet name containing the data (default: first worksheet)",
  )
  .option(
    "--template-slide <number>",
    "template slide number, counted from 1 (default: 1)",
    positiveInteger,
  )
  .requiredOption(
    "-o, --output <path>",
    "where to save the new populated .pptx presentation",
  )
  .option(
    "--header-row <number>",
    "row containing field names, counted from 1",
    positiveInteger,
  )
  .option(
    "-f, --force",
    "replace the output presentation if it already exists; use with care",
  )
  .addHelpText(
    "after",
    `
Example:
  consultchimps pptx populate --template profile.pptx --data clients.xlsx --output profiles.pptx

Put placeholders such as {{client_name}} or Revenue: {{revenue}} in ordinary
text shapes on the template slide. Each nonempty row below the Excel header
creates one slide, in worksheet order. The first worksheet and first slide are
used unless you select them. Empty cells become empty text.

The output contains only the generated slides. Source files are never changed.
`,
  )
  .action(async (options: PptxPopulateOptions) => {
    const result = await populatePowerPointTemplate({
      headerRow: options.headerRow,
      outputPath: options.output,
      overwrite: options.force === true,
      templatePath: options.template,
      templateSlide: options.templateSlide,
      workbookPath: options.data,
      worksheet: options.sheet,
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
    const result = await splitPdf({
      input: inputPath,
      outputDirectory,
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
    const result = await mergePdfs({
      inputs: inputPaths,
      output: options.output,
      overwrite: options.force === true,
    });
    printResult(result, program.opts<GlobalOptions>().json === true);
  });

try {
  await program.parseAsync(process.argv);
} catch (error) {
  if (error instanceof CommanderError) {
    // With exitOverride active Commander reports --help and --version as errors
    // too. Those are successful terminations whose output Commander has already
    // written, so there is nothing to add and nothing to fail.
    if (error.exitCode !== 0) {
      if (jsonRequested) {
        printJsonFailure(error.message, USAGE_ERROR_CODE);
      }
      // Without --json Commander has already written its own usage prose, so
      // only its exit status needs carrying over.
      process.exitCode = error.exitCode;
    }
  } else {
    const json = program.opts<GlobalOptions>().json === true;
    const expected = isConsultChimpsError(error);
    const message = expected
      ? error.message
      : error instanceof Error
        ? error.message
        : String(error);
    // Only expected operational failures carry a published, stable error code.
    // Unexpected errors report null rather than inventing a code an automation
    // could come to depend on.
    const code = expected ? error.code : null;

    if (json) {
      printJsonFailure(message, code);
    } else {
      process.stderr.write(
        formatHumanError(message, expected ? error.code : undefined, {
          vocabulary: CLI_VOCABULARY,
        }),
      );
    }
    process.exitCode = 1;
  }
}
