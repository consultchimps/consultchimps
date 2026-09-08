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
 * ## What it accepts, and what it refuses
 *
 * It follows RFC 4180. Where a file departs from that, this parser either
 * refuses it with a stable error or tolerates it deliberately; there is no
 * third case where something malformed is quietly reshaped into data. The whole
 * contract:
 *
 * Accepted as written:
 *
 * - The first record that is not blank is the header row.
 * - A field may be quoted with double quotes, and a quoted field may hold the
 *   separator, a line break, or a doubled quote standing for one literal quote.
 * - CRLF, LF, and a lone CR each end a record.
 * - A leading byte-order mark is dropped, because a file exported from Excel on
 *   Windows carries one.
 *
 * Refused, with a stable error code:
 *
 * - `DB_CSV_TEXT_AFTER_QUOTE`: anything other than the separator, a line break,
 *   or the end of the file directly after a field's closing quote. The field has
 *   already ended, so `"Acme"x` has no reading that is not a guess.
 * - `DB_CSV_UNCLOSED_QUOTE`: a quoted field the file never closes. The rest of
 *   the document would otherwise become one field.
 * - `DB_CSV_EXTRA_VALUES`: a record with a value that no header names, which
 *   would have to be discarded to import the rest.
 * - `DB_CSV_NO_HEADER_ROW`: a file with no record to read column names from.
 * - `DB_CSV_INVALID_DELIMITER`: a separator that is empty, a double quote, or a
 *   line break, none of which the scan could tell apart from its own syntax.
 *
 * Tolerated deliberately, and covered by tests so the tolerance stays a choice:
 *
 * - A double quote that is not the first character of a field is an ordinary
 *   character, so `5" pipe` reads as itself. RFC 4180 leaves this undefined; a
 *   quote only opens a quoted field where a field begins.
 * - A record with fewer values than the header has empty cells at the end,
 *   which is unambiguous.
 * - A trailing separator whose surplus values are all blank is a stray
 *   separator, not data, and is dropped.
 * - A blank record anywhere is skipped, as the workbook reader skips a blank
 *   worksheet row.
 * - An empty field becomes null rather than an empty string, so a missing value
 *   and an empty one are one thing rather than two.
 * - A blank or repeated header is filled in and numbered rather than refused,
 *   as the workbook reader does for a worksheet's header row.
 */

/** Options for reading delimited text. */
export interface ParseCsvOptions {
  /** The field separator. Defaults to a comma. */
  readonly delimiter?: string | undefined;
  /** Recorded as the table's source file name, for provenance. */
  readonly file?: string | undefined;
}

const BYTE_ORDER_MARK = "﻿";

/** One record, with the file line it began on so an error can point at it. */
interface ParsedRecord {
  readonly fields: string[];
  /** The one-based line the record started on, line breaks inside a quoted
   *  field included in the count, so it is the line a person can open. */
  readonly line: number;
}

/** Whether every field of a record is blank, which is a row with no content. */
function isBlankRecord(fields: readonly string[]): boolean {
  return fields.every((field) => field.trim() === "");
}

/**
 * Split delimited text into records of raw fields. Quoting state decides
 * whether a separator or a line break separates anything, which is why this is
 * a character scan rather than a split on newlines.
 *
 * The scan has three states, not two. `quoted` is inside a quoted field;
 * `closed` is the moment just after a field's closing quote, where the only
 * legal characters left are the separator, a line break, or the end of the
 * file. Without that third state the closing quote is indistinguishable from
 * the end of quoting mid-field, and `"Acme"x` silently becomes `Acmex`.
 */
function splitRecords(text: string, delimiter: string): ParsedRecord[] {
  const records: ParsedRecord[] = [];
  let record: string[] = [];
  let field = "";
  let quoted = false;
  let closed = false;
  let index = 0;
  let line = 1;
  let recordLine = 1;
  // Where the quote of an unclosed field was opened, for the error that names
  // it. Read only while `quoted` is true.
  let quoteLine = 1;
  let quoteColumn = 1;

  const endField = (): void => {
    record.push(field);
    field = "";
    closed = false;
  };
  const endRecord = (): void => {
    endField();
    records.push({ fields: record, line: recordLine });
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
        closed = true;
        index += 1;
        continue;
      }
      if (character === "\n" || character === "\r") {
        // A line break inside a quoted field is content, and it still advances
        // the file's line count so a later error points at the right line. CRLF
        // is one break and is consumed together, so it counts once.
        if (character === "\r" && text[index + 1] === "\n") {
          field += "\r\n";
          index += 2;
        } else {
          field += character;
          index += 1;
        }
        line += 1;
        continue;
      }
      field += character;
      index += 1;
      continue;
    }

    const isDelimiter = text.startsWith(delimiter, index);
    const isLineBreak = character === "\r" || character === "\n";

    if (closed && !isDelimiter && !isLineBreak) {
      throw new ConsultChimpsError(
        "DB_CSV_TEXT_AFTER_QUOTE",
        `Row ${line}, column ${record.length + 1} has text after the closing quote of a quoted value, so where the value ends is unclear. Put the whole value inside the quotes, or double any quote that is part of it.`,
        { details: { column: record.length + 1, row: line } },
      );
    }

    // A quote opens a quoted field only at the start of a field. Anywhere else
    // it is an ordinary character, so a bare `5" pipe` reads as its own text
    // rather than turning the rest of the file into one field.
    if (character === '"' && field === "" && !closed) {
      quoted = true;
      quoteLine = line;
      quoteColumn = record.length + 1;
      index += 1;
      continue;
    }
    if (isDelimiter) {
      endField();
      index += delimiter.length;
      continue;
    }
    if (character === "\r") {
      endRecord();
      index += text[index + 1] === "\n" ? 2 : 1;
      line += 1;
      recordLine = line;
      continue;
    }
    if (character === "\n") {
      endRecord();
      index += 1;
      line += 1;
      recordLine = line;
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
      `The quoted value that starts at row ${quoteLine}, column ${quoteColumn} is never closed, so the file ends inside it. Add the closing double quote, or double any quote that is part of the value.`,
      { details: { column: quoteColumn, row: quoteLine } },
    );
  }
  // A file that ends with a line break has already closed its last record, so
  // there is nothing pending; anything else still holds one.
  if (record.length > 0 || field !== "" || closed) {
    endRecord();
  }
  return records;
}

/**
 * Read delimited text as a `Table`, under the contract this module's
 * documentation states. Every value comes back as text, or null where the field
 * was empty, so the type of a column is decided later and in one place.
 */
export function parseCsvTable(
  text: string,
  options: ParseCsvOptions = {},
): Table {
  const delimiter = options.delimiter ?? ",";
  // A separator that is also quoting or a line break has no coherent reading:
  // the scan below would take it as both at once. Refusing it keeps the parser
  // from producing something plausible out of a request that means nothing.
  if (delimiter === "" || /["\r\n]/u.test(delimiter)) {
    throw new ConsultChimpsError(
      "DB_CSV_INVALID_DELIMITER",
      "The field separator cannot be empty, a double quote, or a line break.",
      { details: {} },
    );
  }
  const body = text.startsWith(BYTE_ORDER_MARK) ? text.slice(1) : text;
  const records = splitRecords(body, delimiter);

  const headerIndex = records.findIndex(
    (record) => !isBlankRecord(record.fields),
  );
  const header = headerIndex === -1 ? undefined : records[headerIndex];
  if (header === undefined) {
    throw new ConsultChimpsError(
      "DB_CSV_NO_HEADER_ROW",
      "The file holds no rows, so there is no header row to read the column names from.",
      { details: {} },
    );
  }

  const columns = uniqueHeaders(
    header.fields.map((value) => (value.trim() === "" ? null : value.trim())),
  );

  const rows: TableRow[] = [];
  for (const record of records.slice(headerIndex + 1)) {
    const { fields, line } = record;
    if (isBlankRecord(fields)) {
      continue;
    }
    if (fields.length > columns.length) {
      const surplus = fields.slice(columns.length);
      if (!isBlankRecord(surplus)) {
        throw new ConsultChimpsError(
          "DB_CSV_EXTRA_VALUES",
          `Row ${line} holds ${fields.length} values but the header row names ${columns.length} columns, so ${surplus.length} value${surplus.length === 1 ? " has" : "s have"} no column to go in. Add the missing header${surplus.length === 1 ? "" : "s"} and try again.`,
          {
            details: {
              columnCount: columns.length,
              row: line,
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
