import { ConsultChimpsError } from "@consultchimps/core";
import {
  uniqueHeaders,
  type Table,
  type TableRow,
} from "@consultchimps/tabular";

/**
 * A comma-separated text reader, written here rather than borrowed.
 *
 * `@consultchimps/xlsx` reads OOXML packages; nothing in the toolkit read
 * delimited text, and the workbook reader's spreadsheet engine guesses types
 * while it parses, which would decide "007" is the number 7 before this package
 * ever saw it. Import has to make that judgement itself, once, under the rules
 * in `inferColumnTypes`, so the parser here deliberately produces text and
 * nothing else: every field arrives as the characters the file held, and only
 * the inference step decides what they mean.
 *
 * It follows RFC 4180: the first record is the header row, fields may be quoted
 * with double quotes, a doubled quote inside a quoted field is one literal
 * quote, and a quoted field may span line breaks. CRLF, LF, and CR all end a
 * record, and a leading byte-order mark is dropped, because a file exported
 * from Excel on Windows carries both.
 */

/** Options for reading delimited text. */
export interface ParseCsvOptions {
  /** The field separator. Defaults to a comma. */
  readonly delimiter?: string | undefined;
  /** Recorded as the table's source file name, for provenance. */
  readonly file?: string | undefined;
}

const BYTE_ORDER_MARK = "﻿";

/** Whether every field of a record is blank, which is a row with no content. */
function isBlankRecord(fields: readonly string[]): boolean {
  return fields.every((field) => field.trim() === "");
}

/**
 * Split delimited text into records of raw fields. Quoting state decides
 * whether a delimiter or a line break separates anything, which is why this is
 * a character scan rather than a split on newlines.
 */
function splitRecords(text: string, delimiter: string): string[][] {
  const records: string[][] = [];
  let record: string[] = [];
  let field = "";
  let quoted = false;
  let index = 0;

  const endField = (): void => {
    record.push(field);
    field = "";
  };
  const endRecord = (): void => {
    endField();
    records.push(record);
    record = [];
  };

  while (index < text.length) {
    const character = text[index] as string;

    if (quoted) {
      if (character === '"') {
        if (text[index + 1] === '"') {
          // A doubled quote inside a quoted field is one literal quote.
          field += '"';
          index += 2;
          continue;
        }
        quoted = false;
        index += 1;
        continue;
      }
      field += character;
      index += 1;
      continue;
    }

    // A quote opens a quoted field only at the start of a field. Anywhere else
    // it is an ordinary character, so a bare `5" pipe` reads as its own text
    // rather than turning the rest of the file into one field.
    if (character === '"' && field === "") {
      quoted = true;
      index += 1;
      continue;
    }
    if (text.startsWith(delimiter, index)) {
      endField();
      index += delimiter.length;
      continue;
    }
    if (character === "\r") {
      endRecord();
      index += text[index + 1] === "\n" ? 2 : 1;
      continue;
    }
    if (character === "\n") {
      endRecord();
      index += 1;
      continue;
    }
    field += character;
    index += 1;
  }

  if (quoted) {
    // The rest of the file was swallowed into one field. Reporting it is the
    // only honest outcome: the alternative is a table whose last column holds
    // the remainder of the document.
    throw new ConsultChimpsError(
      "DB_CSV_UNCLOSED_QUOTE",
      "The file ends inside a quoted value, so a double quote was opened and never closed. Check the quotes in the file and try again.",
      { details: {} },
    );
  }
  // A file that ends with a line break has already closed its last record, so
  // there is nothing pending; anything else still holds one.
  if (record.length > 0 || field !== "") {
    endRecord();
  }
  return records;
}

/**
 * Read delimited text as a `Table`. The first non-blank record is the header
 * row; blank records anywhere are skipped, the way the workbook reader skips
 * blank worksheet rows. Every value comes back as text, or null where the field
 * was empty, so the type of a column is decided later and in one place.
 */
export function parseCsvTable(
  text: string,
  options: ParseCsvOptions = {},
): Table {
  const delimiter = options.delimiter ?? ",";
  if (delimiter === "") {
    throw new ConsultChimpsError(
      "DB_CSV_INVALID_DELIMITER",
      "The field separator cannot be empty.",
      { details: {} },
    );
  }
  const body = text.startsWith(BYTE_ORDER_MARK) ? text.slice(1) : text;
  const records = splitRecords(body, delimiter);

  // Row numbers below are the file's own, counted from 1, so a message points
  // at the line a person can open and look at.
  const headerIndex = records.findIndex((fields) => !isBlankRecord(fields));
  const header = headerIndex === -1 ? undefined : records[headerIndex];
  if (header === undefined) {
    throw new ConsultChimpsError(
      "DB_CSV_NO_HEADER_ROW",
      "The file holds no rows, so there is no header row to read the column names from.",
      { details: {} },
    );
  }

  // Blank and repeated headers are filled in and numbered rather than refused,
  // the same way the workbook reader treats a worksheet's header row, so a
  // spreadsheet with one unnamed column still imports.
  const columns = uniqueHeaders(
    header.map((value) => (value.trim() === "" ? null : value.trim())),
  );

  const rows: TableRow[] = [];
  for (let index = headerIndex + 1; index < records.length; index += 1) {
    const fields = records[index] as string[];
    if (isBlankRecord(fields)) {
      continue;
    }
    if (fields.length > columns.length) {
      const surplus = fields.slice(columns.length);
      // Trailing empty fields are a stray separator, not data, so they are
      // dropped. A value with no header has nowhere to go, and silently
      // discarding it would lose a column of the file.
      if (!isBlankRecord(surplus)) {
        throw new ConsultChimpsError(
          "DB_CSV_EXTRA_VALUES",
          `Row ${index + 1} holds ${fields.length} values but the header row names ${columns.length} columns, so ${surplus.length} value${surplus.length === 1 ? " has" : "s have"} no column to go in. Add the missing header${surplus.length === 1 ? "" : "s"} and try again.`,
          {
            details: {
              columnCount: columns.length,
              row: index + 1,
              valueCount: fields.length,
            },
          },
        );
      }
    }
    // A prototype-free row, so a header spelled like an Object.prototype member
    // is stored as data rather than reaching the prototype chain.
    const row = Object.create(null) as TableRow;
    columns.forEach((column, columnIndex) => {
      const value = fields[columnIndex];
      row[column] = value === undefined || value === "" ? null : value;
    });
    rows.push(row);
  }

  return {
    columns,
    rows,
    ...(options.file === undefined ? {} : { source: { file: options.file } }),
  };
}
