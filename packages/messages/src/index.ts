import type { Artifact, OperationResult } from "@consultchimps/core";

interface OperationExplanation {
  readonly nextSteps: readonly string[];
  readonly summary: (result: OperationResult) => readonly string[];
  readonly title: string;
}

const numberFormatter = new Intl.NumberFormat("en-US");

function metric(result: OperationResult, name: string): number {
  return result.metrics[name] ?? 0;
}

function formatNumber(value: number): string {
  return numberFormatter.format(value);
}

function quantity(
  value: number,
  singular: string,
  plural = `${singular}s`,
): string {
  return `${formatNumber(value)} ${value === 1 ? singular : plural}`;
}

const operationExplanations: Readonly<Record<string, OperationExplanation>> = {
  "sheets.consolidate": {
    title: "Your Excel consolidation is complete.",
    summary: (result) => [
      `ConsultChimps read ${quantity(metric(result, "inputFiles"), "Excel file")} and combined ${quantity(metric(result, "inputTables"), "visible worksheet")}.`,
      `The finished workbook contains ${quantity(metric(result, "outputRows"), "data row")} arranged across ${quantity(metric(result, "outputColumns"), "column")}.`,
      "Your original Excel files were not changed.",
    ],
    nextSteps: [
      "Open the new Excel workbook listed below and review the consolidated worksheet.",
      "Keep the source columns in the workbook if you need to trace a row back to its original file and worksheet.",
    ],
  },
  "sheets.split-by-column": {
    title: "Your Excel workbook split is complete.",
    summary: (result) => [
      `ConsultChimps read ${quantity(metric(result, "inputRows"), "data row")} from the source workbook.`,
      `It found ${quantity(metric(result, "groups"), "distinct group")} and created ${quantity(metric(result, "outputFiles"), "separate Excel workbook")}.`,
      `${quantity(metric(result, "outputRows"), "data row")} ${metric(result, "outputRows") === 1 ? "was" : "were"} written to the new workbooks, and ${quantity(metric(result, "skippedRows"), "row")} ${metric(result, "skippedRows") === 1 ? "was" : "were"} skipped.`,
      "Your original Excel workbook was not changed.",
    ],
    nextSteps: [
      "Open the new workbooks listed below and confirm that each file contains the expected group.",
      "If rows were skipped, review the warning section to understand why.",
    ],
  },
  "pdf.split": {
    title: "Your PDF split is complete.",
    summary: (result) => [
      `ConsultChimps read ${quantity(metric(result, "pages"), "page")} from the source PDF.`,
      `It created ${quantity(metric(result, "outputFiles"), "separate PDF file")}, with one source page in each new file.`,
      "Your original PDF was not changed.",
    ],
    nextSteps: [
      "Open the new PDF files listed below and confirm that the pages are in the expected order.",
      "The page number in each filename identifies its position in the original PDF.",
    ],
  },
  "pdf.merge": {
    title: "Your PDF merge is complete.",
    summary: (result) => [
      `ConsultChimps combined ${quantity(metric(result, "inputFiles"), "PDF file")} in the resolved input order.`,
      `The new PDF contains ${quantity(metric(result, "pages"), "page")}.`,
      "Your original PDF files were not changed.",
    ],
    nextSteps: [
      "Open the new PDF listed below and check that the documents appear in the intended order.",
      "Keep the original PDFs until you have confirmed the merged file is complete.",
    ],
  },
  "pptx.populate": {
    title: "Your PowerPoint presentation is complete.",
    summary: (result) => [
      `ConsultChimps read ${quantity(metric(result, "inputRows"), "nonempty Excel record")} and created ${quantity(metric(result, "generatedSlides"), "populated slide")} in worksheet order.`,
      `It replaced ${quantity(metric(result, "replacements"), "placeholder occurrence")} across the generated slides.`,
      "Your source PowerPoint template and Excel workbook were not changed.",
    ],
    nextSteps: [
      "Open the new PowerPoint presentation listed below and review every generated slide.",
      "Check longer replacement values for fit because this version does not shrink or truncate text automatically.",
    ],
  },
};

const metricLabels: Readonly<Record<string, string>> = {
  generatedSlides: "PowerPoint slides generated",
  groups: "Distinct groups found",
  inputFiles: "Input files read",
  inputRows: "Source data rows read",
  inputTables: "Visible worksheets combined",
  outputColumns: "Columns in the finished spreadsheet",
  outputFiles: "New files created",
  outputRows: "Data rows written",
  pages: "PDF pages processed",
  placeholderFields: "Distinct placeholder fields",
  placeholderOccurrences: "Placeholder occurrences per template slide",
  replacements: "Placeholder replacements made",
  skippedRows: "Rows skipped",
  warnings: "Warnings reported",
};

function artifactType(artifact: Artifact): string {
  if (artifact.kind === "directory") {
    return "Folder";
  }
  if (
    artifact.mediaType ===
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  ) {
    return "Excel workbook";
  }
  if (artifact.mediaType === "application/pdf") {
    return "PDF document";
  }
  if (
    artifact.mediaType ===
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  ) {
    return "PowerPoint presentation";
  }
  return "File";
}

function genericExplanation(result: OperationResult): OperationExplanation {
  return {
    title: "Your task is complete.",
    summary: () => [
      `ConsultChimps completed the "${result.operation}" operation successfully.`,
      "Review the detailed results and created files below.",
    ],
    nextSteps: [
      "Open the files listed below and confirm that they contain the expected results.",
    ],
  };
}

export function formatHumanResult(result: OperationResult): string {
  const explanation =
    operationExplanations[result.operation] ?? genericExplanation(result);
  const lines = [
    "SUCCESS: ConsultChimps finished your task.",
    "",
    explanation.title,
    "",
    "What ConsultChimps did:",
    ...explanation.summary(result).map((line) => `  - ${line}`),
    "",
    "Detailed results:",
    ...Object.entries(result.metrics).map(
      ([name, value]) =>
        `  - ${metricLabels[name] ?? name}: ${formatNumber(value)}`,
    ),
    "",
    "Files created:",
  ];

  if (result.artifacts.length === 0) {
    lines.push(
      "  - No files were created. Review the result details above for more information.",
    );
  } else {
    result.artifacts.forEach((artifact, index) => {
      lines.push(
        `  ${index + 1}. ${artifact.path}`,
        `     Type: ${artifactType(artifact)}`,
      );
    });
  }

  lines.push("", "Warnings:");
  if (result.warnings.length === 0) {
    lines.push(
      "  - None. ConsultChimps did not detect any recoverable problems during this task.",
    );
  } else {
    lines.push(
      ...result.warnings.map((warning, index) => `  ${index + 1}. ${warning}`),
    );
  }

  lines.push(
    "",
    "What you can do next:",
    ...explanation.nextSteps.map((step) => `  - ${step}`),
    "",
  );

  return lines.join("\n");
}

function recoverySteps(code: string | undefined): readonly string[] {
  if (code === "OPERATION_ABORTED") {
    return [
      "The task was cancelled before it finished; no source file was changed.",
      "Output files completed before the cancellation may remain. Review and remove them if they are not wanted.",
      "Run the command again when you are ready to complete the task.",
    ];
  }
  if (code === "FILES_NOT_FOUND") {
    return [
      "Check that every file or folder path is spelled correctly and still exists.",
      "If you used a pattern such as *.xlsx or *.pdf, place it in quotation marks and try again.",
      "Run the command with --help if you want to review the expected input format.",
    ];
  }
  if (code === "FILES_OUTPUT_EXISTS") {
    return [
      "Choose a different output filename or output folder and run the command again.",
      "If you intentionally want to replace the existing output, rerun the command with --force.",
      "Use --force only after confirming that the existing output is safe to replace.",
    ];
  }
  if (code === "FILES_INPUT_OVERWRITE") {
    return [
      "Choose an output path that is different from every source file.",
      "ConsultChimps protects source files and will not replace an input file.",
    ];
  }
  if (code === "XLSX_NO_TABLES") {
    return [
      "Open the source workbooks and confirm that they contain at least one visible worksheet with data.",
      "If the data is on a hidden worksheet, review the --hidden option in the command help.",
    ];
  }
  if (code?.startsWith("XLSX_")) {
    return [
      "Check the workbook, worksheet, table, column name, and command options mentioned in the message above.",
      "Run the command with --help to review the available spreadsheet options and examples.",
    ];
  }
  if (code?.startsWith("PDF_")) {
    return [
      "Confirm that every source file is a readable PDF and that the output location is available.",
      "Run the command with --help to review the available PDF options and examples.",
    ];
  }
  if (code?.startsWith("PPTX_")) {
    return [
      "Check the PowerPoint template slide, placeholder spelling, Excel headers, and output path mentioned above.",
      "Run consultchimps pptx inspect-template to review placeholders before populating the presentation.",
      "Run consultchimps pptx populate --help for a complete example.",
    ];
  }
  return [
    "Read the message above and check the supplied files, folders, and command options.",
    "Run the command again with --help if you need examples.",
    "If the problem continues, keep the error reference below when asking for support.",
  ];
}

export function formatHumanError(message: string, code?: string): string {
  return [
    "ERROR: ConsultChimps could not finish your task.",
    "",
    "What went wrong:",
    `  ${message}`,
    "",
    "What you can do:",
    ...recoverySteps(code).map((step) => `  - ${step}`),
    ...(code
      ? [
          "",
          "Error reference:",
          `  ${code}`,
          "  This reference can help a developer or support person identify the exact type of problem.",
        ]
      : []),
    "",
  ].join("\n");
}
