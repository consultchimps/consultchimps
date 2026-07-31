import type { OperationResult } from "@consultchimps/core";
import { describe, expect, it } from "vitest";

import { formatHumanError, formatHumanResult } from "../src/index.js";

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
    const output = formatHumanResult(value);

    expect(output).toContain("SUCCESS: ConsultChimps finished your task.");
    expect(output).toContain("What ConsultChimps did:");
    expect(output).toContain("Detailed results:");
    expect(output).toContain("Files created:");
    expect(output).toContain("Warnings:");
    expect(output).toContain("What you can do next:");
    expected.forEach((text) => expect(output).toContain(text));
  });

  it("states clearly when a successful operation has no warnings", () => {
    const output = formatHumanResult(
      result("pdf.merge", { inputFiles: 2, outputFiles: 1, pages: 6 }, [
        "combined.pdf",
      ]),
    );

    expect(output).toContain(
      "None. ConsultChimps did not detect any recoverable problems during this task.",
    );
  });

  it("explains how to recover from an existing output error", () => {
    const output = formatHumanError(
      "The output file already exists.",
      "FILES_OUTPUT_EXISTS",
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
      const output = formatHumanError("Something failed.", code);
      expect(output).toContain(expected);
      expect(output).toContain(code);
    },
  );

  it("omits the error reference when no code is available", () => {
    const output = formatHumanError("Something failed.");
    expect(output).toContain("Read the message above");
    expect(output).not.toContain("Error reference:");
  });

  it("labels artifact types and falls back for unknown operations", () => {
    const output = formatHumanResult({
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
    });

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
    );
    expect(output).toContain("No files were created.");
  });
});
