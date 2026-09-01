import type {
  WorkbookColumnDescription,
  WorkbookDescription,
} from "@consultchimps/xlsx";

/**
 * The human-readable half of `sheets inspect`.
 *
 * An inspection creates nothing, so its metrics are counts and everything a
 * reader actually came for (worksheet names, header spellings, sample values)
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
 * Text from the workbook, made safe to print.
 *
 * Sheet names, headers, table names, and cell values are input, and input files
 * are untrusted. This report goes straight to a terminal, where a control
 * character is an instruction rather than a character: a header carrying an
 * escape can move the cursor, recolor or erase the lines already printed, or
 * reach a terminal feature, so a crafted workbook could make the report say
 * something the workbook does not contain. Excel's own writer stores a control
 * character as the literal text `_x001b_`, but a numeric XML character
 * reference decodes to the character itself, so a hand-built package reaches
 * this function with the real thing.
 *
 * Every control character is therefore shown as a visible `\\uXXXX` escape,
 * which also keeps a newline inside a cell from breaking the report's layout. A
 * backslash the workbook holds is doubled, so an escape in the report can never
 * be confused with text that merely looks like one. The value itself is
 * untouched: `--json` reports the raw string, where JSON's own escaping already
 * makes a control character inert.
 */
function printable(text: string): string {
  let rendered = "";
  for (const character of text) {
    const code = character.codePointAt(0) ?? 0;
    // C0 (including DEL) and C1: the ranges a terminal reads as commands.
    // The test is a comparison rather than a character class because a
    // regular expression carrying these characters is itself the thing lint
    // warns about, and the bounds say which ranges are meant.
    const isControl = code < 0x20 || (code >= 0x7f && code <= 0x9f);
    if (isControl) {
      rendered += `\\u${code.toString(16).toUpperCase().padStart(4, "0")}`;
    } else if (character === "\\") {
      rendered += "\\\\";
    } else {
      rendered += character;
    }
  }
  return rendered;
}

/**
 * One sample value, written the way the workbook stores it. Text is quoted and
 * every other stored value is written bare, because the description keeps the
 * number 1 and the text "1" apart on purpose and a reader comparing two columns
 * needs to see which one a cell holds.
 *
 * The quotes are the sample's boundary, so a quote inside the text is escaped
 * too: without that, the single stored string `North", 999` would read as the
 * text "North" followed by the number 999, which is a wrong answer about both
 * the values and their types.
 */
function formatSampleValue(value: SampleValue): string {
  if (typeof value !== "string") {
    return String(value);
  }
  return `"${printable(value).replaceAll('"', '\\"')}"`;
}

function formatColumn(column: WorkbookColumnDescription): string {
  const samples =
    column.sampleValues.length > 0
      ? column.sampleValues.map(formatSampleValue).join(", ")
      : "no sample values";
  // The position is the column's own zero-based index, shown counted from 1 so
  // it reads like a spreadsheet column rather than an array offset.
  return `       ${column.index + 1}. ${printable(column.header)}: ${samples}`;
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
    `Excel workbook inspection: ${printable(description.source)}`,
    "",
    "Worksheets:",
  ];

  if (description.sheets.length === 0) {
    lines.push("  - None. No worksheet matched the selection.");
  }
  description.sheets.forEach((sheet, index) => {
    lines.push(
      `  ${index + 1}. ${printable(sheet.name)} (${sheet.visibility})`,
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
      `  ${index + 1}. ${printable(table.name)} on worksheet ${printable(
        table.sheet,
      )} (${printable(table.range)})`,
      `     Columns: ${table.headers.map(printable).join(", ") || "None"}`,
    );
  });

  lines.push("", "Named ranges:");
  if (description.namedRanges.length === 0) {
    lines.push("  - None found in the described worksheets.");
  }
  description.namedRanges.forEach((range, index) => {
    lines.push(
      `  ${index + 1}. ${printable(range.name)} on worksheet ${printable(
        range.sheet,
      )} (${printable(range.ref)})`,
    );
  });

  lines.push("");
  return lines.join("\n");
}
