/**
 * L2: how a split treats a CSV input (ADR 0007). A CSV file has no workbook
 * to keep, so its split always rebuilds each output from the rows.
 */
import { ConsultChimpsError } from "@consultchimps/core";

import { XLSX_ERRORS } from "../errors.js";
import { settleCsvOptions, type CsvReadOptions } from "./options.js";
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
  },
>(name: string, options: T): T {
  settleCsvOptions(options.csv);
  if (!isCsvName(name)) return options;
  if (options.preserveWorkbook === true) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_CSV_PRESERVE,
      `${name} is a CSV file, so there is no workbook to keep. Leave out the request to keep the workbook, and each output is a new workbook holding its group's rows.`,
      { details: { source: name } },
    );
  }
  return { ...options, preserveWorkbook: false };
}
