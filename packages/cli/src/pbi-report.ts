import { printable } from "./text.js";

/**
 * The human-readable half of `pbi export`.
 *
 * The result explains what the export did in counts; this is the part a reader
 * came for: which tables reached the workbook, which worksheet each one is on,
 * and what was left behind. Every name here comes from the manifest the export
 * wrote, which is the one place a name from inside the model is allowed to
 * appear, and every one of them is rendered terminal-safe: a model can name a
 * table anything at all, including a terminal escape.
 *
 * The rendering lives in the CLI rather than in `@consultchimps/messages`
 * because the manifest is a `@consultchimps/pbi` structure, and a package that
 * depends only on `@consultchimps/core` should not learn a format adapter's
 * types to print one command's report.
 */

/** Only the manifest fields this report reads, parsed from the written bytes. */
interface ReportReason {
  readonly code: string;
  readonly count: number;
}

interface ReportColumn {
  readonly name: string;
  readonly reasons: readonly ReportReason[];
}

interface ReportPart {
  readonly sheetName: string;
  readonly start: number;
  readonly end: number;
}

interface ReportTable {
  readonly name: string;
  readonly parts?: readonly ReportPart[];
  readonly reasons: readonly ReportReason[];
  readonly columns: readonly ReportColumn[];
}

export interface ExportManifestReport {
  readonly tables: readonly ReportTable[];
  readonly excludedTables: readonly ReportTable[];
  readonly unverifiedPaths: readonly ReportReason[];
}

const numberFormatter = new Intl.NumberFormat("en-US");

function count(value: number, singular: string): string {
  return `${numberFormatter.format(value)} ${value === 1 ? singular : `${singular}s`}`;
}

/** Reason codes for one line, in the manifest's own order, never reworded. */
function codes(reasons: readonly ReportReason[]): string {
  return reasons.map((reason) => reason.code).join(", ");
}

function partSummary(parts: readonly ReportPart[]): string {
  if (parts.length === 0) return "no worksheet";
  return parts
    .map(
      (part) =>
        `${printable(part.sheetName)} (rows ${numberFormatter.format(part.start + 1)} to ${numberFormatter.format(part.end)})`,
    )
    .join(", ");
}

/**
 * Parse the manifest bytes the export produced. The bytes are this process's
 * own output, so a failure to read them back is a defect rather than bad input,
 * and the caller treats it as one.
 */
export function readExportManifest(bytes: Uint8Array): ExportManifestReport {
  return JSON.parse(new TextDecoder().decode(bytes)) as ExportManifestReport;
}

export function formatPowerBiExport(manifest: ExportManifestReport): string {
  const lines: string[] = ["Tables exported:"];
  if (manifest.tables.length === 0) lines.push("  - None.");
  manifest.tables.forEach((table, index) => {
    const columnsLeftOut = table.columns.filter(
      (column) => column.reasons.length > 0,
    );
    lines.push(
      `  ${index + 1}. ${printable(table.name)}: ${partSummary(table.parts ?? [])}`,
      `     ${count(table.columns.length - columnsLeftOut.length, "column")} exported, ${count(columnsLeftOut.length, "column")} left out`,
    );
    for (const column of columnsLeftOut)
      lines.push(
        `       - ${printable(column.name)}: ${codes(column.reasons)}`,
      );
  });

  lines.push("", "Tables left out:");
  if (manifest.excludedTables.length === 0)
    lines.push("  - None. Every table in the model was exported.");
  manifest.excludedTables.forEach((table, index) => {
    lines.push(
      `  ${index + 1}. ${printable(table.name)}: ${codes(table.reasons)}`,
    );
    for (const column of table.columns.filter(
      (entry) => entry.reasons.length > 0,
    ))
      lines.push(
        `       - ${printable(column.name)}: ${codes(column.reasons)}`,
      );
  });

  lines.push("", "Readings with no verified example:");
  if (manifest.unverifiedPaths.length === 0)
    lines.push(
      "  - None. Every reading this export used has been checked against a real file.",
    );
  for (const path of manifest.unverifiedPaths) lines.push(`  - ${path.code}`);

  lines.push("");
  return lines.join("\n");
}
