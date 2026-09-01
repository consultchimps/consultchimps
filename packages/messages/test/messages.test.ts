import type { OperationResult } from "@consultchimps/core";
import { describe, expect, it } from "vitest";

import {
  CLI_VOCABULARY,
  formatHumanError,
  formatHumanResult,
  GENERIC_VOCABULARY,
  type MessageVocabulary,
} from "../src/index.js";

const cli = { vocabulary: CLI_VOCABULARY } as const;

function result(
  operation: string,
  metrics: Record<string, number>,
  paths: string[],
  warnings: string[] = [],
): OperationResult {
  return {
    operation,
    artifacts: paths.map((path) => ({ kind: "file", path })),
    metrics,
    warnings,
  };
}

describe("human-readable CLI output", () => {
  it.each([
    {
      expected: [
        "Your Excel consolidation is complete.",
        "2 Excel files",
        "3 visible worksheets",
        "10 data rows",
        "5 columns",
        "Your original Excel files were not changed.",
      ],
      value: result(
        "sheets.consolidate",
        {
          inputFiles: 2,
          inputTables: 3,
          outputColumns: 5,
          outputRows: 10,
        },
        ["combined.xlsx"],
      ),
    },
    {
      expected: [
        "Your Excel workbook split is complete.",
        "10 data rows",
        "3 distinct groups",
        "3 separate Excel workbooks",
        "1 row was skipped",
        "Your original Excel workbook was not changed.",
      ],
      value: result(
        "sheets.split-by-column",
        {
          groups: 3,
          inputFiles: 1,
          inputRows: 10,
          outputFiles: 3,
          outputRows: 9,
          skippedRows: 1,
        },
        ["North.xlsx", "South.xlsx", "West.xlsx"],
        ['Skipped 1 row with a blank value in "Region".'],
      ),
    },
    {
      expected: [
        "Your PDF split is complete.",
        "3 pages",
        "3 separate PDF files",
        "Your original PDF was not changed.",
      ],
      value: result("pdf.split", { inputFiles: 1, outputFiles: 3, pages: 3 }, [
        "page-001.pdf",
        "page-002.pdf",
        "page-003.pdf",
      ]),
    },
    {
      expected: [
        "Your PDF merge is complete.",
        "2 PDF files",
        "6 pages",
        "Your original PDF files were not changed.",
      ],
      value: result("pdf.merge", { inputFiles: 2, outputFiles: 1, pages: 6 }, [
        "combined.pdf",
      ]),
    },
  ])("explains $value.operation in plain language", ({ expected, value }) => {
    const output = formatHumanResult(value, cli);

    expect(output).toContain("SUCCESS: ConsultChimps finished your task.");
    expect(output).toContain("What ConsultChimps did:");
    expect(output).toContain("Detailed results:");
    expect(output).toContain("Files created:");
    expect(output).toContain("Warnings:");
    expect(output).toContain("What you can do next:");
    expect(output).toContain("listed below");
    expected.forEach((text) => expect(output).toContain(text));
  });

  it("explains worksheet merges without exposing raw metric names", () => {
    const output = formatHumanResult(
      result(
        "sheets.merge",
        { hiddenSheets: 1, inputFiles: 2, outputSheets: 4 },
        ["merged.xlsx"],
        ["1 source worksheet was hidden."],
      ),
      cli,
    );

    expect(output).toContain("Your Excel workbook merge is complete.");
    expect(output).toContain("4 worksheets");
    expect(output).toContain("2 Excel files");
    expect(output).toContain("1 source worksheet was hidden");
    expect(output).toContain("Your original Excel files were not changed.");
    expect(output).not.toContain("outputSheets:");
  });

  it("states clearly when a successful operation has no warnings", () => {
    const output = formatHumanResult(
      result("pdf.merge", { inputFiles: 2, outputFiles: 1, pages: 6 }, [
        "combined.pdf",
      ]),
      cli,
    );

    expect(output).toContain(
      "None. ConsultChimps did not detect any recoverable problems during this task.",
    );
  });

  it("explains how to recover from an existing output error", () => {
    const output = formatHumanError(
      "The output file already exists.",
      "FILES_OUTPUT_EXISTS",
      cli,
    );

    expect(output).toContain(
      "ERROR: ConsultChimps could not finish your task.",
    );
    expect(output).toContain("What went wrong:");
    expect(output).toContain("Choose a different output filename");
    expect(output).toContain("rerun the command with --force");
    expect(output).toContain("Error reference:");
    expect(output).toContain("FILES_OUTPUT_EXISTS");
  });

  it.each([
    {
      code: "OPERATION_ABORTED",
      expected: "The task was cancelled before it finished",
    },
    {
      code: "FILES_NOT_FOUND",
      expected: "Check that every file or folder path is spelled correctly",
    },
    {
      code: "FILES_INPUT_OVERWRITE",
      expected: "ConsultChimps protects source files",
    },
    {
      code: "XLSX_NO_TABLES",
      expected: "at least one visible worksheet with data",
    },
    {
      code: "XLSX_SPLIT_NO_TABLE",
      expected: "Check the workbook, worksheet, table, column name",
    },
    {
      code: "TABLE_MAPPING_COLUMN_COLLISION",
      expected:
        "Nothing was created or changed: a column mapping is checked and applied before any output is written.",
    },
    {
      code: "XLSX_MAPPING_DATE_NOT_TEXT",
      expected: "Open the column mapping and check the canonical column names",
    },
    {
      code: "PDF_NO_PAGES",
      expected: "Confirm that every source file is a readable PDF",
    },
    {
      code: "PPTX_NO_DATA_ROWS",
      expected: "consultchimps pptx inspect-template",
    },
    {
      code: "SOMETHING_UNEXPECTED",
      expected: "keep the error reference below when asking for support",
    },
  ])(
    "offers recovery steps tailored to $code",
    ({ code, expected }: { code: string; expected: string }) => {
      const output = formatHumanError("Something failed.", code, cli);
      expect(output).toContain(expected);
      expect(output).toContain(code);
    },
  );

  it("explains a consolidation that mapped columns and drafted a mapping", () => {
    const mapped = formatHumanResult(
      result(
        "sheets.consolidate",
        {
          inputFiles: 2,
          inputTables: 2,
          outputColumns: 5,
          outputRows: 10,
          suggestedColumns: 0,
          unmappedColumns: 2,
        },
        ["combined.xlsx"],
        ['2 columns did not match the column mapping: "Region", "Owner".'],
      ),
      cli,
    );
    expect(mapped).toContain(
      "2 columns did not match the column mapping and kept their own names",
    );
    expect(mapped).toContain(
      "Columns that did not match the column mapping: 2",
    );
    expect(mapped).not.toContain("drafted a column mapping");

    const drafted = formatHumanResult(
      {
        operation: "sheets.consolidate",
        artifacts: [
          { kind: "file", path: "combined.xlsx" },
          {
            kind: "file",
            mediaType: "application/json",
            path: "mapping.json",
          },
        ],
        metrics: {
          inputFiles: 2,
          inputTables: 2,
          outputColumns: 5,
          outputRows: 10,
          suggestedColumns: 3,
          unmappedColumns: 0,
        },
        warnings: [],
      },
      cli,
    );
    expect(drafted).toContain(
      "It also drafted a column mapping proposing 3 canonical columns, and applied none of them.",
    );
    expect(drafted).toContain(
      "Review and edit the drafted column mapping listed below",
    );
    expect(drafted).not.toContain("did not match the column mapping and");

    // A plain consolidation adds neither sentence; the two counts still
    // appear among the detailed results, reporting zero.
    const plain = formatHumanResult(
      result(
        "sheets.consolidate",
        {
          inputFiles: 2,
          inputTables: 2,
          outputColumns: 5,
          outputRows: 10,
          suggestedColumns: 0,
          unmappedColumns: 0,
        },
        ["combined.xlsx"],
      ),
      cli,
    );
    expect(plain).not.toContain("did not match the column mapping and");
    expect(plain).not.toContain("drafted a column mapping");
    expect(plain).toContain("Columns that did not match the column mapping: 0");

    // A draft over headers that already agree is a real file proposing
    // nothing, so the reader is still told what it is and that it was not
    // applied. The count cannot answer that; the artifact can.
    const empty = formatHumanResult(
      {
        operation: "sheets.consolidate",
        artifacts: [
          { kind: "file", path: "combined.xlsx" },
          {
            kind: "file",
            mediaType: "application/json",
            path: "mapping.json",
          },
        ],
        metrics: {
          inputFiles: 2,
          inputTables: 2,
          outputColumns: 5,
          outputRows: 10,
          suggestedColumns: 0,
          unmappedColumns: 0,
        },
        warnings: [],
      },
      cli,
    );
    expect(empty).toContain(
      "It also drafted a column mapping. Every header was already spelled the same way, so the draft proposes no canonical columns.",
    );
    expect(empty).toContain("Review and edit the drafted column mapping");
  });

  it("names a written column mapping as its own kind of file", () => {
    const output = formatHumanResult(
      {
        operation: "sheets.consolidate",
        artifacts: [
          {
            kind: "file",
            mediaType: "application/json",
            path: "mapping.json",
          },
        ],
        metrics: {},
        warnings: [],
      },
      cli,
    );
    expect(output).toContain("Type: Column mapping file");
  });

  it("omits the error reference when no code is available", () => {
    const output = formatHumanError("Something failed.", undefined, cli);
    expect(output).toContain("Read the message above");
    expect(output).toContain(
      "Run the command again with --help if you need examples.",
    );
    expect(output).not.toContain("Error reference:");
  });

  it("explains a template inspection without pointing at output files", () => {
    const output = formatHumanResult(
      result(
        "pptx.inspect-template",
        {
          malformedPlaceholderLocations: 1,
          placeholderFields: 2,
          placeholderOccurrences: 3,
          unsupportedPlacementPlaceholders: 0,
          unsupportedSplitRunPlaceholders: 0,
        },
        [],
        ["Slide 1 has 1 location with malformed placeholder braces."],
      ),
      cli,
    );

    expect(output).toContain(
      "Your PowerPoint template inspection is complete.",
    );
    expect(output).toContain(
      "ConsultChimps found 2 distinct placeholder fields on the inspected template slide, used 3 times in total.",
    );
    expect(output).toContain("Nothing was created or changed.");

    // Every metric reads as plain language, never as its internal name.
    expect(output).toContain("Locations with malformed placeholder braces: 1");
    expect(output).toContain("Distinct placeholder fields: 2");
    expect(output).toContain("Placeholders split across text runs: 0");
    expect(output).not.toContain("malformedPlaceholderLocations:");
    expect(output).not.toContain("unsupportedSplitRunPlaceholders:");

    // An inspection creates nothing, so the reader is never told to open a
    // file that does not exist.
    expect(output).toContain("No files were created.");
    expect(output).not.toContain("Open the files");

    // The names are not in this text, only the counts are, so the next step
    // sends the reader to the inspection report rather than "above".
    expect(output).toContain(
      "Read the placeholder names from the inspection report that accompanies this result",
    );
    expect(output).not.toContain("placeholder field listed above");
    expect(output).toContain(
      "1. Slide 1 has 1 location with malformed placeholder braces.",
    );
    expect(output).toContain(
      "Run consultchimps pptx populate --help for a complete example.",
    );
  });

  it("explains a workbook inspection without pointing at output files", () => {
    const output = formatHumanResult(
      result(
        "sheets.inspect",
        {
          dataRows: 12,
          excelTables: 1,
          headerColumns: 5,
          hiddenWorksheets: 1,
          namedRanges: 0,
          worksheets: 2,
        },
        [],
        [
          "1 worksheet is hidden and was not described. Include hidden worksheets to describe it.",
        ],
      ),
      cli,
    );

    expect(output).toContain("Your Excel workbook inspection is complete.");
    expect(output).toContain(
      "ConsultChimps described 2 worksheets, holding 5 columns and 12 data rows in total.",
    );
    expect(output).toContain("It also found 1 Excel Table and 0 named ranges.");
    expect(output).toContain("Nothing was created or changed.");

    // Every metric reads as plain language, never as its internal name.
    expect(output).toContain("Worksheets described: 2");
    expect(output).toContain("Columns described across worksheets: 5");
    expect(output).toContain("Data rows described: 12");
    expect(output).toContain("Excel Tables found: 1");
    expect(output).toContain("Named ranges found: 0");
    expect(output).toContain("Hidden worksheets described: 1");
    expect(output).not.toContain("headerColumns:");
    expect(output).not.toContain("hiddenWorksheets:");

    // An inspection creates nothing, so the reader is never told to open a
    // file that does not exist.
    expect(output).toContain("No files were created.");
    expect(output).not.toContain("Open the files");
    expect(output).not.toContain("Open the new");

    // The names are not in this text, only the counts are, so the next step
    // sends the reader to the description rather than "above".
    expect(output).toContain(
      "Read the worksheet names, column headers, and sample values from the description that accompanies this result",
    );
    expect(output).toContain("1. 1 worksheet is hidden and was not described.");
  });

  it("labels artifact types and falls back for unknown operations", () => {
    const output = formatHumanResult(
      {
        operation: "future.operation",
        artifacts: [
          { kind: "directory", path: "outputs" },
          {
            kind: "file",
            path: "deck.pptx",
            mediaType:
              "application/vnd.openxmlformats-officedocument.presentationml.presentation",
          },
          { kind: "file", path: "notes.bin" },
        ],
        warnings: ["One recoverable warning."],
        metrics: { customMetric: 2 },
      },
      cli,
    );

    expect(output).toContain('"future.operation" operation successfully');
    expect(output).toContain("Folder");
    expect(output).toContain("PowerPoint presentation");
    expect(output).toContain("Type: File");
    expect(output).toContain("customMetric: 2");
    expect(output).toContain("1. One recoverable warning.");
  });

  it("states plainly when no files were created", () => {
    const output = formatHumanResult(
      result("pdf.merge", { inputFiles: 0, outputFiles: 0, pages: 0 }, []),
      cli,
    );
    expect(output).toContain("No files were created.");
  });
});

describe("interface-neutral output", () => {
  const commandLineVocabulary = ["--force", "--help", "consultchimps"];

  const errorCodes = [
    "OPERATION_ABORTED",
    "FILES_NOT_FOUND",
    "FILES_OUTPUT_EXISTS",
    "FILES_INPUT_OVERWRITE",
    "XLSX_NO_TABLES",
    "XLSX_SPLIT_NO_TABLE",
    "PDF_NO_PAGES",
    "PPTX_NO_DATA_ROWS",
    "SOMETHING_UNEXPECTED",
    undefined,
  ] as const;

  it.each(errorCodes)(
    "keeps recovery guidance free of command-line vocabulary for %s",
    (code) => {
      const output = formatHumanError("Something failed.", code);

      expect(output).toContain("ERROR: ConsultChimps could not finish");
      expect(output).toContain("What you can do:");
      commandLineVocabulary.forEach((phrase) =>
        expect(output).not.toContain(phrase),
      );
    },
  );

  it.each(errorCodes)(
    "still explains how to recover from %s without a command line",
    (code) => {
      const output = formatHumanError("Something failed.", code);
      const steps = output
        .split("\n")
        .filter((line) => line.startsWith("  - "));

      expect(steps.length).toBeGreaterThanOrEqual(2);
      steps.forEach((step) => expect(step.trim().length).toBeGreaterThan(20));
    },
  );

  it("describes overwrite recovery without naming a flag", () => {
    const output = formatHumanError(
      "The output file already exists.",
      "FILES_OUTPUT_EXISTS",
    );

    expect(output).toContain(
      "Choose a different output filename or output folder and start the task again.",
    );
    expect(output).toContain(
      "If you intentionally want to replace the existing output, allow the existing output to be replaced and try again.",
    );
    expect(output).not.toContain("--force");
  });

  it("describes template inspection without naming an executable", () => {
    const output = formatHumanError("Something failed.", "PPTX_NO_DATA_ROWS");

    expect(output).toContain(
      "Inspect the PowerPoint template to review its placeholders before populating the presentation.",
    );
    expect(output).not.toContain("consultchimps");
  });

  it("points at created files without assuming a scrolling transcript", () => {
    const output = formatHumanResult(
      result("pdf.split", { inputFiles: 1, outputFiles: 2, pages: 2 }, [
        "page-001.pdf",
        "page-002.pdf",
      ]),
    );

    expect(output).toContain(
      "Open the new PDF files shown in the list of created files",
    );
    expect(output).not.toContain("listed below");
    commandLineVocabulary.forEach((phrase) =>
      expect(output).not.toContain(phrase),
    );
  });

  it("uses the generic vocabulary when no options are supplied", () => {
    const value = result("pdf.merge", { inputFiles: 2, outputFiles: 1 }, [
      "combined.pdf",
    ]);

    expect(formatHumanResult(value)).toBe(
      formatHumanResult(value, { vocabulary: GENERIC_VOCABULARY }),
    );
    expect(formatHumanError("Something failed.", "FILES_NOT_FOUND")).toBe(
      formatHumanError("Something failed.", "FILES_NOT_FOUND", {
        vocabulary: GENERIC_VOCABULARY,
      }),
    );
  });

  it("accepts a custom vocabulary supplied by another interface", () => {
    const desktop: MessageVocabulary = {
      ...GENERIC_VOCABULARY,
      actionNoun: "job",
      artifactListReference: "in the results panel",
      retryWithOverwrite: "Switch on Replace existing files and run it again.",
    };

    const output = formatHumanError(
      "The output file already exists.",
      "FILES_OUTPUT_EXISTS",
      { vocabulary: desktop },
    );
    expect(output).toContain("Switch on Replace existing files");

    const unknown = formatHumanError("Something failed.", "MYSTERY", {
      vocabulary: desktop,
    });
    expect(unknown).toContain("folders, and job options");

    const success = formatHumanResult(
      result("pdf.merge", { inputFiles: 2, outputFiles: 1, pages: 6 }, [
        "combined.pdf",
      ]),
      { vocabulary: desktop },
    );
    expect(success).toContain("Open the new PDF in the results panel");
  });

  it("keeps the command-line vocabulary complete", () => {
    const keys = Object.keys(GENERIC_VOCABULARY).sort();
    expect(Object.keys(CLI_VOCABULARY).sort()).toEqual(keys);
    keys.forEach((key) => {
      const phrase = CLI_VOCABULARY[key as keyof MessageVocabulary];
      expect(phrase.length).toBeGreaterThan(0);
    });
  });
});
