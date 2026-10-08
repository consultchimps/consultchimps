/**
 * L2: how a CSV file is read (ADR 0007), as a caller asks and as a read
 * settles it.
 */
import { ConsultChimpsError } from "@consultchimps/core";

import { XLSX_ERRORS } from "../errors.js";
import type { CsvDateOrder } from "./typing.js";

export type { CsvDateOrder };

/** The encodings a CSV file can be read in. */
export type CsvEncoding = "utf-8" | "utf-16le" | "utf-16be" | "windows-1252";

export const CSV_ENCODINGS: readonly CsvEncoding[] = [
  "utf-8",
  "utf-16le",
  "utf-16be",
  "windows-1252",
];

export const CSV_DATE_ORDERS: readonly CsvDateOrder[] = [
  "iso",
  "dmy",
  "mdy",
  "ymd",
];

/** How to read the CSV files among an operation's inputs. */
export interface CsvReadOptions {
  /** The text encoding; detected from the bytes when absent. */
  encoding?: CsvEncoding | undefined;
  /** The field delimiter, one character; guessed from the file when absent. */
  delimiter?: string | undefined;
  /** Read fields that are only a number as numbers. Off by default. */
  numbers?: boolean | undefined;
  /** The decimal separator numbers use, `.` or `,`. Default `.`. */
  decimalSeparator?: string | undefined;
  /**
   * The separator grouping a number's whole part in threes, or `""` for none.
   * Default `,`, or `.` when the decimal separator is `,`.
   */
  thousandsSeparator?: string | undefined;
  /**
   * Read dates: `iso` reads `yyyy-mm-dd` only; `dmy`, `mdy` and `ymd` also
   * read dates written in that order. Off by default.
   */
  dates?: CsvDateOrder | undefined;
}

/** Options checked and filled in. */
export interface SettledCsvOptions {
  readonly encoding: CsvEncoding | undefined;
  readonly delimiter: string | undefined;
  readonly numbers:
    { readonly decimal: string; readonly thousands: string } | undefined;
  readonly dates: CsvDateOrder | undefined;
}

const DECIMAL_SEPARATORS = new Set([".", ","]);
const THOUSANDS_SEPARATORS = new Set([
  "",
  ",",
  ".",
  " ",
  "'",
  "\u00a0",
  "\u202f",
]);
const FORBIDDEN_DELIMITERS = new Set(['"', "\r", "\n", "\ufeff"]);

function invalid(message: string, option: string): ConsultChimpsError {
  return new ConsultChimpsError(XLSX_ERRORS.XLSX_CSV_INVALID_OPTION, message, {
    details: { option },
  });
}

/** Check the options before any file is read, and fill in the defaults. */
export function settleCsvOptions(
  options: CsvReadOptions | undefined,
): SettledCsvOptions {
  const encoding = options?.encoding;
  if (encoding !== undefined && !CSV_ENCODINGS.includes(encoding)) {
    throw invalid(
      `The CSV encoding "${String(encoding)}" is not one ConsultChimps reads. Choose ${CSV_ENCODINGS.join(", ")}.`,
      "encoding",
    );
  }
  const delimiter = options?.delimiter;
  if (
    delimiter !== undefined &&
    (delimiter.length !== 1 || FORBIDDEN_DELIMITERS.has(delimiter))
  ) {
    throw invalid(
      "A CSV delimiter is one character, and cannot be a quote or a line break.",
      "delimiter",
    );
  }
  const dates = options?.dates;
  if (dates !== undefined && !CSV_DATE_ORDERS.includes(dates)) {
    throw invalid(
      `The CSV date order "${String(dates)}" is not one ConsultChimps reads. Choose ${CSV_DATE_ORDERS.join(", ")}.`,
      "dates",
    );
  }
  const decimal = options?.decimalSeparator ?? ".";
  if (!DECIMAL_SEPARATORS.has(decimal)) {
    throw invalid(
      'The decimal separator for CSV numbers is "." or ",".',
      "decimalSeparator",
    );
  }
  const thousands =
    options?.thousandsSeparator ?? (decimal === "," ? "." : ",");
  if (!THOUSANDS_SEPARATORS.has(thousands)) {
    throw invalid(
      'The thousands separator for CSV numbers is ",", ".", a space, an apostrophe, or none.',
      "thousandsSeparator",
    );
  }
  if (thousands === decimal) {
    throw invalid(
      `The decimal and thousands separators for CSV numbers are both "${decimal}". Choose two different separators.`,
      "thousandsSeparator",
    );
  }
  return {
    encoding,
    delimiter,
    numbers: options?.numbers === true ? { decimal, thousands } : undefined,
    dates,
  };
}
