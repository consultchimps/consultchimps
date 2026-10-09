/**
 * L2: which file a table operation writes, a workbook or CSV (ADR 0007), from
 * the output's name and the format a caller asked for.
 */
import { ConsultChimpsError } from "@consultchimps/core";

import { XLSX_ERRORS } from "../errors.js";
import type { TableOutputFormat } from "../table-output.js";
import { CSV_EXTENSION, CSV_MEDIA_TYPE } from "./writer.js";

/** The formats a table operation can be asked for by name. */
export type TableOutputFormatName = "xlsx" | "csv";

const WORKBOOK_EXTENSION = ".xlsx";
const WORKBOOK_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

export interface ResolvedTableOutput {
  readonly format: TableOutputFormat;
  readonly extension: string;
  readonly mediaType: string;
}

/** The format a name's extension says, or undefined when it says neither. */
function formatOfName(
  name: string | undefined,
): TableOutputFormatName | undefined {
  if (name === undefined) return undefined;
  if (/\.csv$/iu.test(name)) return "csv";
  if (/\.xls[xm]$/iu.test(name)) return "xlsx";
  return undefined;
}

/**
 * How an output named `name` is written. Its extension decides; `requested`
 * decides when the name has neither extension; `fallback` when neither does.
 * A request that contradicts the name is refused before anything is read.
 */
export function resolveTableOutput(
  name: string | undefined,
  requested: TableOutputFormatName | undefined,
  csvBom: boolean | undefined,
  fallback: TableOutputFormatName = "xlsx",
): ResolvedTableOutput {
  if (requested !== undefined && requested !== "xlsx" && requested !== "csv") {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_OUTPUT_FORMAT_INVALID,
      `The output format "${String(requested)}" is not one ConsultChimps writes. Choose xlsx or csv.`,
      { details: { outputFormat: requested } },
    );
  }
  const named = formatOfName(name);
  if (requested !== undefined && named !== undefined && requested !== named) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_OUTPUT_FORMAT_INVALID,
      `The output ${String(name)} is named as ${named === "csv" ? "a CSV file" : "a workbook"}, but ${requested === "csv" ? "CSV" : "a workbook"} was asked for. Name the output with the extension of the format you want.`,
      { details: { output: name, outputFormat: requested } },
    );
  }
  const kind = requested ?? named ?? fallback;
  return kind === "csv"
    ? {
        format: { kind: "csv", bom: csvBom ?? true },
        extension: CSV_EXTENSION,
        mediaType: CSV_MEDIA_TYPE,
      }
    : {
        format: { kind: "xlsx" },
        extension: WORKBOOK_EXTENSION,
        mediaType: WORKBOOK_MEDIA_TYPE,
      };
}

/** Refuse a CSV name for an output that can only be a workbook, such as a merge's. */
export function refuseCsvOutputName(
  name: string | undefined,
  operation: string,
): void {
  if (formatOfName(name) !== "csv") return;
  throw new ConsultChimpsError(
    XLSX_ERRORS.XLSX_OUTPUT_FORMAT_INVALID,
    `A ${operation} writes a workbook of separate tabs, which a CSV file cannot hold, so ${String(name)} cannot be its output. Name the output with .xlsx.`,
    { details: { output: name } },
  );
}
