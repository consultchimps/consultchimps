/**
 * The options every sheets command that reads CSV files takes (ADR 0007), and
 * their translation into the library's `CsvReadOptions`.
 */
import { InvalidArgumentError, type Command } from "commander";
import type { CsvReadOptions } from "@consultchimps/xlsx";

/** The CSV flags as Commander collects them. */
export interface CsvCliOptions {
  csvDates?: string;
  csvDecimal?: string;
  csvDelimiter?: string;
  csvEncoding?: string;
  csvNumbers?: boolean;
  csvThousands?: string;
}

const DELIMITERS: Record<string, string> = {
  comma: ",",
  semicolon: ";",
  tab: "\t",
  pipe: "|",
  // A tab typed as a backslash and a t, as shells pass it.
  "\\t": "\t",
};

const THOUSANDS: Record<string, string> = {
  none: "",
  space: " ",
  apostrophe: "'",
};

const ENCODINGS: Record<string, string> = {
  utf8: "utf-8",
  utf16le: "utf-16le",
  utf16be: "utf-16be",
  cp1252: "windows-1252",
};

/** Add the CSV reading flags to a command. */
export function withCsvOptions(command: Command): Command {
  return command
    .option(
      "--csv-delimiter <delimiter>",
      "CSV field delimiter: comma, semicolon, tab, pipe, or one character (default: guessed from the file)",
    )
    .option(
      "--csv-encoding <encoding>",
      "CSV text encoding: utf-8, utf-16le, utf-16be, or windows-1252 (default: detected from the file; a byte order mark wins)",
    )
    .option(
      "--csv-numbers",
      "read CSV fields that hold only a number as numbers; every field is text otherwise",
    )
    .option(
      "--csv-decimal <separator>",
      'decimal separator --csv-numbers reads: "." or "," (default: ".")',
    )
    .option(
      "--csv-thousands <separator>",
      'thousands separator --csv-numbers reads: ",", ".", space, apostrophe, or none (default: ",", or "." when the decimal is ",")',
    )
    .option(
      "--csv-dates <order>",
      "read CSV dates: iso reads yyyy-mm-dd only; dmy, mdy, or ymd also read dates in that order",
    );
}

/** The library options the flags ask for, or undefined when none were given. */
export function csvReadOptions(
  options: CsvCliOptions,
): CsvReadOptions | undefined {
  // Checked by the library, which refuses separators without numbers with a
  // stable code, as it does for every surface.
  const read: CsvReadOptions = {};
  if (options.csvDelimiter !== undefined) {
    read.delimiter =
      DELIMITERS[options.csvDelimiter.toLowerCase()] ?? options.csvDelimiter;
  }
  if (options.csvEncoding !== undefined) {
    const named = options.csvEncoding.toLowerCase();
    read.encoding = (ENCODINGS[named] ?? named) as CsvReadOptions["encoding"];
  }
  if (options.csvNumbers === true) read.numbers = true;
  if (options.csvDecimal !== undefined) {
    read.decimalSeparator = options.csvDecimal;
  }
  if (options.csvThousands !== undefined) {
    read.thousandsSeparator =
      THOUSANDS[options.csvThousands.toLowerCase()] ?? options.csvThousands;
  }
  if (options.csvDates !== undefined) {
    read.dates = options.csvDates.toLowerCase() as CsvReadOptions["dates"];
  }
  return Object.keys(read).length === 0 ? undefined : read;
}

/** The output flags as Commander collects them. */
export interface CsvOutputCliOptions {
  outputFormat?: string;
  csvBom?: string;
}

/** Add the flags that choose CSV output to a command. */
export function withCsvOutputOptions(command: Command): Command {
  return command
    .option(
      "--output-format <format>",
      "write xlsx or csv; an output name ending in .csv or .xlsx decides it too, and the two must agree",
    )
    .option(
      "--csv-bom <true|false>",
      "start a CSV output with a UTF-8 byte order mark so Excel opens it as UTF-8 (default: true)",
      (value: string) => {
        const lower = value.toLowerCase();
        if (lower !== "true" && lower !== "false") {
          throw new InvalidArgumentError("Use true or false.");
        }
        return lower;
      },
    );
}

/** The library's output options the flags ask for. */
export function csvOutputOptions(options: CsvOutputCliOptions): {
  outputFormat?: "xlsx" | "csv";
  csvBom?: boolean;
} {
  const chosen: { outputFormat?: "xlsx" | "csv"; csvBom?: boolean } = {};
  if (options.outputFormat !== undefined) {
    chosen.outputFormat = options.outputFormat.toLowerCase() as "xlsx" | "csv";
  }
  if (options.csvBom !== undefined) {
    const value = options.csvBom.toLowerCase();
    if (value !== "true" && value !== "false") {
      throw new Error(
        `--csv-bom takes true or false, not "${options.csvBom}".`,
      );
    }
    chosen.csvBom = value === "true";
  }
  return chosen;
}
