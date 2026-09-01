import type {
  WorkbookColumnDescription,
  WorkbookDescription,
} from "@consultchimps/xlsx";

/**
 * The human-readable half of `sheets inspect`.
 *
 * An inspection creates nothing, so its metrics are counts and everything a
 * reader actually came for - worksheet names, header spellings, sample values -
 * travels beside the result in the description. `@consultchimps/messages`
 * renders the result and points at "the description that accompanies this
 * result"; this module is that description, and printing it is what makes the
 * next step it names possible.
 *
 * The rendering lives in the CLI rather than in the messages package because
 * `WorkbookDescription` is an xlsx type: rendering it there would give a
 * package that depends only on `@consultchimps/core` a dependency on a format
 * adapter, to describe a structure only this command prints.
 */

type SampleValue = WorkbookColumnDescription["sampleValues"][number];

/** A count with its noun, so no line reads "1 rows". */
function quantity(value: number, singular: string): string {
  return `${value} ${value === 1 ? singular : `${singular}s`}`;
}

/**
 * One sample value, written the way the workbook stores it. Text is quoted and
 * every other stored value is written bare, because the description keeps the
 * number 1 and the text "1" apart on purpose and a reader comparing two columns
 * needs to see which one a cell holds.
 */
function formatSampleValue(value: SampleValue): string {
  return typeof value === "string" ? `"${value}"` : String(value);
}

function formatColumn(column: WorkbookColumnDescription): string {
  const samples =
    column.sampleValues.length > 0
      ? column.sampleValues.map(formatSampleValue).join(", ")
      : "no sample values";
  // The position is the column's own zero-based index, shown counted from 1 so
  // it reads like a spreadsheet column rather than an array offset.
  return `       ${column.index + 1}. ${column.header}: ${samples}`;
}

/**
 * Render a workbook description as plain text for stdout.
 *
 * The order is the description's own order, which is workbook order, so the
 * same workbook and options always produce the same report.
 */
export function formatWorkbookDescription(
  description: WorkbookDescription,
): string {
  const lines = [
    `Excel workbook inspection: ${description.source}`,
    "",
    "Worksheets:",
  ];

  if (description.sheets.length === 0) {
    lines.push("  - None. No worksheet matched the selection.");
  }
  description.sheets.forEach((sheet, index) => {
    lines.push(
      `  ${index + 1}. ${sheet.name} (${sheet.visibility})`,
      `     Used range: ${quantity(sheet.rowCount, "row")} by ${quantity(
        sheet.columnCount,
        "column",
      )}`,
      `     Header row: ${sheet.headerRow ?? "none found"}`,
      `     Data rows below the header: ${sheet.dataRowCount}`,
      "     Columns and their sample values:",
      ...(sheet.columns.length > 0
        ? sheet.columns.map(formatColumn)
        : ["       - None"]),
    );
  });

  lines.push("", "Excel Tables:");
  if (description.excelTables.length === 0) {
    lines.push("  - None found in the described worksheets.");
  }
  description.excelTables.forEach((table, index) => {
    lines.push(
      `  ${index + 1}. ${table.name} on worksheet ${table.sheet} (${table.range})`,
      `     Columns: ${table.headers.join(", ") || "None"}`,
    );
  });

  lines.push("", "Named ranges:");
  if (description.namedRanges.length === 0) {
    lines.push("  - None found in the described worksheets.");
  }
  description.namedRanges.forEach((range, index) => {
    lines.push(
      `  ${index + 1}. ${range.name} on worksheet ${range.sheet} (${range.ref})`,
    );
  });

  lines.push("");
  return lines.join("\n");
}
