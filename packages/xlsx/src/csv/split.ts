/**
 * L2: how a split treats a CSV input (ADR 0007). A CSV file has no workbook
 * to keep, so its split always rebuilds each output from the rows.
 */
import { ConsultChimpsError } from "@consultchimps/core";

import { XLSX_ERRORS } from "../errors.js";
import { settleCsvOptions, type CsvReadOptions } from "./options.js";
import { resolveTableOutput, type TableOutputFormatName } from "./output.js";
import { isCsvName } from "./reader.js";

/**
 * The split options for an input named `name`, checked: CSV options are
 * settled whatever the input, and a CSV input splits compactly, refusing a
 * request to keep a workbook it does not have.
 */
export function splitOptionsFor<
  T extends {
    preserveWorkbook?: boolean | undefined;
    csv?: CsvReadOptions | undefined;
    outputFormat?: TableOutputFormatName | undefined;
    table?: string | undefined;
    range?: string | undefined;
    sheet?: string | undefined;
  },
>(
  name: string,
  options: T,
): T & { outputFormat: TableOutputFormatName; tolerantMatching?: boolean } {
  settleCsvOptions(options.csv);
  const csvInput = isCsvName(name);
  // A CSV input splits into CSV files unless a workbook is asked for.
  const outputFormat = options.outputFormat ?? (csvInput ? "csv" : "xlsx");
  resolveTableOutput(undefined, outputFormat, undefined);
  if (!csvInput && outputFormat !== "csv") return { ...options, outputFormat };
  if (options.preserveWorkbook === true) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_CSV_PRESERVE,
      csvInput
        ? `${name} is a CSV file, so there is no workbook to keep. Leave out the request to keep the workbook, and each output holds its group's rows.`
        : `A CSV output holds a table's rows only, so the workbook cannot be kept. Leave out the request to keep the workbook, or ask for workbooks instead of CSV files.`,
      { details: { source: name, outputFormat } },
    );
  }
  // A split that names no source is the default split, so it matches values
  // the way the whole-workbook split does, written compactly or not.
  const defaultSplit =
    csvInput ||
    (options.table === undefined &&
      options.range === undefined &&
      options.sheet === undefined);
  return {
    ...options,
    outputFormat,
    preserveWorkbook: false,
    ...(defaultSplit ? { tolerantMatching: true } : {}),
  };
}
