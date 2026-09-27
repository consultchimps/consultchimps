import type { Artifact, OperationResult } from "@consultchimps/core";
import { describe, expect, it } from "vitest";

import {
  CLI_VOCABULARY,
  formatHumanError,
  formatHumanResult,
  GENERIC_VOCABULARY,
} from "../src/index.js";

/**
 * Every Power BI code rendered in both vocabularies. The rules asserted here are
 * the ones a reviewer would otherwise have to check code by code: each code is
 * explained rather than falling through to the last-resort wording, the generic
 * wording never names a flag or an executable, and nothing anywhere names a
 * table, a column, or a path from inside the file.
 */

const cli = { vocabulary: CLI_VOCABULARY } as const;
const generic = { vocabulary: GENERIC_VOCABULARY } as const;

// Every refusal the package documents, plus one it does not, which must still
// be explained as a Power BI problem rather than as an unknown one.
const PBI_CODES = [
  "PBI_INVALID_OPTIONS",
  "PBI_INVALID_CONTAINER",
  "PBI_NO_MODEL",
  "PBI_MODEL_ENCRYPTED",
  "PBI_MODEL_UNREADABLE",
  "PBI_NO_EXPORTABLE_TABLES",
  "PBI_RUNTIME_UNAVAILABLE",
  "PBI_EXPORT_LIMIT_EXCEEDED",
  "PBI_SOMETHING_ADDED_LATER",
  // The codes the command line raises itself, which only its own users can
  // receive, plus one it does not raise yet.
  "CLI_PBI_ONE_INPUT",
  "CLI_PBI_OUTPUT_INCOMPLETE",
  "CLI_PBI_OUTPUT_CLEANUP_REQUIRED",
  "CLI_PBI_UNEXPECTED_OUTPUTS",
  "CLI_PBI_SOMETHING_ADDED_LATER",
] as const;

const LAST_RESORT = "keep the error reference below when asking for support";
const COMMAND_LINE_VOCABULARY = [
  "--force",
  "--help",
  "--include-hidden",
  "--max-memory",
  "consultchimps",
];

function steps(output: string): string[] {
  return output.split("\n").filter((line) => line.startsWith("  - "));
}

function exportResult(warnings: string[] = []): OperationResult {
  const artifacts: Artifact[] = [
    {
      kind: "file",
      path: "workbook.xlsx",
      mediaType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
    { kind: "file", path: "manifest.json", mediaType: "application/json" },
  ];
  return {
    operation: "pbi.export",
    artifacts,
    metrics: {
      inputFiles: 1,
      outputFiles: 2,
      exportedTables: 3,
      exportedColumns: 17,
      exportedRows: 1250,
      outputWorksheets: 4,
    },
    warnings,
  };
}

describe("Power BI refusals", () => {
  it.each(PBI_CODES)("explains %s in both vocabularies", (code) => {
    for (const options of [cli, generic]) {
      const output = formatHumanError("The reader stopped.", code, options);
      expect(output).toContain("What you can do:");
      expect(output).toContain(code);
      // Not the last-resort wording: every PBI code has guidance of its own.
      expect(output).not.toContain(LAST_RESORT);
      const recovery = steps(output);
      expect(recovery.length).toBeGreaterThanOrEqual(2);
      for (const step of recovery)
        expect(step.trim().length).toBeGreaterThan(20);
    }
  });

  it.each(PBI_CODES)(
    "keeps the generic wording free of command-line vocabulary for %s",
    (code) => {
      const output = formatHumanError("The reader stopped.", code, generic);
      for (const phrase of COMMAND_LINE_VOCABULARY)
        expect(output).not.toContain(phrase);
    },
  );

  it.each(PBI_CODES)("names nothing from inside the file for %s", (code) => {
    for (const options of [cli, generic]) {
      // The message is the library's; the guidance is this package's. Only the
      // guidance is under test here, so the message is replaced by a marker.
      const output = formatHumanError("MESSAGE", code, options);
      const guidance = steps(output).join("\n");
      for (const forbidden of [
        "Sales",
        "Customer",
        "metadata.sqlitedb",
        "DataModel",
        ".dictionary",
        ".idf",
      ])
        expect(guidance).not.toContain(forbidden);
    }
  });

  it("offers the hidden-table option only where hidden tables were excluded", () => {
    const hidden = { ...cli, details: { hiddenExcluded: true } };
    const notHidden = { ...cli, details: { hiddenExcluded: false } };
    expect(
      formatHumanError("Nothing survived.", "PBI_NO_EXPORTABLE_TABLES", hidden),
    ).toContain("--include-hidden");
    expect(
      formatHumanError("Nothing survived.", "PBI_NO_EXPORTABLE_TABLES", {
        ...generic,
        details: { hiddenExcluded: true },
      }),
    ).toContain("turn on the option that includes hidden tables");
    // A model whose tables were all unreadable is not helped by including
    // hidden ones, and telling the reader to try it wastes their next attempt.
    expect(
      formatHumanError(
        "Nothing survived.",
        "PBI_NO_EXPORTABLE_TABLES",
        notHidden,
      ),
    ).not.toContain("--include-hidden");
    expect(
      formatHumanError("Nothing survived.", "PBI_NO_EXPORTABLE_TABLES", cli),
    ).not.toContain("--include-hidden");
    expect(
      formatHumanError("The model is encrypted.", "PBI_MODEL_ENCRYPTED", cli),
    ).not.toContain("--include-hidden");
  });

  it("says what is on disk for each of the command's own refusals", () => {
    expect(
      formatHumanError("Half written.", "CLI_PBI_OUTPUT_INCOMPLETE", cli),
    ).toContain("not a complete export");
    expect(
      formatHumanError("Litter.", "CLI_PBI_OUTPUT_CLEANUP_REQUIRED", cli),
    ).toContain("the export is complete");
    expect(formatHumanError("Two files.", "CLI_PBI_ONE_INPUT", cli)).toContain(
      "Nothing was read",
    );
    expect(
      formatHumanError("Odd outputs.", "CLI_PBI_UNEXPECTED_OUTPUTS", cli),
    ).toContain("Nothing was written");
  });

  it("offers the memory ceiling only for a capacity refusal", () => {
    expect(
      formatHumanError("Over the limit.", "PBI_EXPORT_LIMIT_EXCEEDED", cli),
    ).toContain("--max-memory");
    expect(
      formatHumanError("Over the limit.", "PBI_EXPORT_LIMIT_EXCEEDED", generic),
    ).toContain("Raise the memory the export is allowed to use");
    expect(
      formatHumanError("No model here.", "PBI_NO_MODEL", cli),
    ).not.toContain("--max-memory");
  });

  it("says nothing was written where nothing was", () => {
    for (const code of [
      "PBI_INVALID_OPTIONS",
      "PBI_MODEL_UNREADABLE",
      "PBI_NO_EXPORTABLE_TABLES",
      "PBI_EXPORT_LIMIT_EXCEEDED",
      "CLI_PBI_ONE_INPUT",
      "CLI_PBI_UNEXPECTED_OUTPUTS",
    ])
      expect(formatHumanError("The reader stopped.", code, cli)).toMatch(
        /Nothing was (written|read)|No workbook was written/,
      );
  });
});

describe("a successful Power BI export", () => {
  it("reports the six metrics and names both outputs for what they are", () => {
    const output = formatHumanResult(exportResult(), cli);
    expect(output).toContain("Your Power BI tables were exported.");
    expect(output).toContain("3 tables");
    expect(output).toContain("4 worksheets");
    expect(output).toContain("17 columns");
    expect(output).toContain("1,250 rows");
    expect(output).toContain("Power BI tables exported: 3");
    expect(output).toContain("Worksheets written: 4");
    expect(output).toContain("Type: Excel workbook");
    expect(output).toContain("Type: Power BI export manifest");
    expect(output).not.toContain("Column mapping file");
    expect(output).toContain("Your original Power BI file was not changed.");
  });

  it("sends the reader to the manifest for what was left out", () => {
    const output = formatHumanResult(exportResult(), generic);
    expect(output).toContain("Read the manifest");
    for (const phrase of COMMAND_LINE_VOCABULARY)
      expect(output).not.toContain(phrase);
  });

  it("lists the library's warnings without rewording them", () => {
    const warning =
      "2 columns were left out because this reader failed while decoding them.";
    expect(formatHumanResult(exportResult([warning]), cli)).toContain(warning);
  });

  it("still labels a column mapping as one", () => {
    const mapping = formatHumanResult(
      {
        operation: "sheets.consolidate",
        artifacts: [
          {
            kind: "file",
            path: "mapping.json",
            mediaType: "application/json",
          },
        ],
        metrics: {
          inputFiles: 1,
          inputTables: 1,
          outputColumns: 2,
          outputRows: 3,
        },
        warnings: [],
      },
      cli,
    );
    expect(mapping).toContain("Type: Column mapping file");
  });
});
