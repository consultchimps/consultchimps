/**
 * The options every sheets command that reads CSV files takes (ADR 0007), and
 * their translation into the library's `CsvReadOptions`.
 */
import type { Command } from "commander";
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
