/**
 * L3: a CSV file read as a workbook of one worksheet (ADR 0007), so every
 * operation that reads streamed worksheets reads it the same way.
 *
 * Papa Parse splits the text into rows and fields; this module decides how the
 * bytes become text and what each field holds. The bytes are read in pieces
 * from a random-access source, a file handle or `Blob.slice`, and decoded with
 * `TextDecoder` in streaming mode, so a character split between pieces is
 * decoded whole and the file is never held whole.
 *
 * Opening reads the file once to settle the encoding and to take a fingerprint;
 * each worksheet read then reads it again from the start. Rows are split on LF,
 * and a carriage return ending a row's last field is part of its line ending,
 * so files mixing CRLF and LF read correctly; a file with no LF at all is split
 * on CR.
 */
import Papa from "papaparse";
import {
  ConsultChimpsError,
  type RandomAccessSource,
} from "@consultchimps/core";

import { XLSX_ERRORS } from "../errors.js";
import type { ExcelTableDefinition } from "../excel-tables.js";
import type { CellRectangle } from "../model/references.js";
import type {
  StreamedCell,
  StreamedDefinedName,
  StreamedSheet,
  StreamedValue,
  StreamedWorkbookContext,
  WorksheetConsumer,
  WorksheetRead,
} from "../operations/consolidate/reader.js";
import { readFailure } from "../operations/read-model.js";
import {
  settleCsvOptions,
  type CsvEncoding,
  type CsvReadOptions,
  type SettledCsvOptions,
} from "./options.js";
import { CsvNumberReader, csvDate } from "./typing.js";

/** How many bytes each read takes. The same everywhere, so reads are identical. */
export const CSV_PIECE_BYTES: number = 64 * 1024;

/**
 * The most text a row may run to before it is taken as a quote that opens a
 * field and is never closed, which would otherwise carry the rest of the file.
 */
export const CSV_MAX_OPEN_ROW_CHARS: number = 16 * 1024 * 1024;

/** The delimiters a read guesses between, in Papa's order of preference. */
const GUESSED_DELIMITERS = [",", "\t", "|", ";"];

/** The part name a CSV's one worksheet carries; it names no zip entry. */
const CSV_PART = "csv";

/** Whether a file name names a CSV file. */
export function isCsvName(name: string): boolean {
  return /\.csv$/iu.test(name);
}

/** The worksheet a CSV file reads as is named after the file, without `.csv`. */
export function csvSheetName(file: string): string {
  const base = file.split(/[\\/]/u).pop() ?? file;
  const stem = base.replace(/\.csv$/iu, "");
  return stem === "" ? "Sheet1" : stem;
}

/** How the encoding a read uses was settled. */
export type CsvEncodingSource =
  "chosen" | "byte-order-mark" | "valid-utf-8" | "fallback";

function byteOrderMark(head: Uint8Array): CsvEncoding | undefined {
  if (head[0] === 0xef && head[1] === 0xbb && head[2] === 0xbf) return "utf-8";
  if (head[0] === 0xff && head[1] === 0xfe) return "utf-16le";
  if (head[0] === 0xfe && head[1] === 0xff) return "utf-16be";
  return undefined;
}

const ENCODING_NAMES: Record<CsvEncoding, string> = {
  "utf-8": "UTF-8",
  "utf-16le": "UTF-16 LE",
  "utf-16be": "UTF-16 BE",
  "windows-1252": "Windows-1252",
};

/** The name an encoding goes by in messages. */
export function csvEncodingName(encoding: CsvEncoding): string {
  return ENCODING_NAMES[encoding];
}

const DELIMITER_NAMES: Record<string, string> = {
  ",": "comma",
  ";": "semicolon",
  "\t": "tab",
  "|": "pipe",
};

/** The name a delimiter goes by in messages. */
export function csvDelimiterName(delimiter: string): string {
  return DELIMITER_NAMES[delimiter] ?? `"${delimiter}"`;
}

/** The minimal readable stream Papa Parse reads text from. */
class TextFeed {
  readonly readable = true;
  readonly #listeners = new Map<string, (value?: string) => void>();

  read(): null {
    return null;
  }

  on(event: string, listener: (value?: string) => void): this {
    this.#listeners.set(event, listener);
    return this;
  }

  removeListener(event: string): this {
    this.#listeners.delete(event);
    return this;
  }

  // Papa pauses and resumes its input only when its caller pauses it, which
  // this module never does: each piece is parsed before the next is read.
  pause(): void {}

  resume(): void {}

  data(text: string): void {
    this.#listeners.get("data")?.(text);
  }

  end(): void {
    this.#listeners.get("end")?.();
  }
}

/** A CSV file opened for reading as one worksheet. */
export class CsvWorkbook {
  readonly sheets: readonly StreamedSheet[];
  readonly tables: readonly ExcelTableDefinition[] = [];
  readonly names: readonly StreamedDefinedName[] = [];
  readonly encoding: CsvEncoding;
  readonly encodingSource: CsvEncodingSource;
  readonly delimiter: string;
  /** What the read noticed and the caller should report, such as a fallback encoding. */
  readonly warnings: readonly string[];
  /** Changes whenever the file's bytes change. */
  readonly fingerprint: number;
  readonly #source: RandomAccessSource;
  readonly #context: StreamedWorkbookContext;
  readonly #options: SettledCsvOptions;
  readonly #newline: "\n" | "\r";

  private constructor(
    source: RandomAccessSource,
    context: StreamedWorkbookContext,
    options: SettledCsvOptions,
    settled: {
      encoding: CsvEncoding;
      encodingSource: CsvEncodingSource;
      delimiter: string;
      newline: "\n" | "\r";
      fingerprint: number;
    },
  ) {
    this.#source = source;
    this.#context = context;
    this.#options = options;
    this.encoding = settled.encoding;
    this.encodingSource = settled.encodingSource;
    this.delimiter = settled.delimiter;
    this.#newline = settled.newline;
    this.fingerprint = settled.fingerprint;
    this.sheets = [
      {
        name: csvSheetName(context.file),
        visible: true,
        visibility: "visible",
        part: CSV_PART,
      },
    ];
    this.warnings =
      settled.encodingSource === "fallback"
        ? [
            `${context.file} is not valid UTF-8 and has no byte order mark, so it was read as Windows-1252. If its accented letters look wrong, choose its encoding and run again.`,
          ]
        : [];
  }

  /**
   * Open a CSV file: read it once to settle its encoding, then guess its
   * delimiter from its first rows. `options` are checked first, so unusable
   * options cost no read.
   */
  static async open(
    source: RandomAccessSource,
    context: StreamedWorkbookContext,
    options?: CsvReadOptions,
  ): Promise<CsvWorkbook> {
    const settled = settleCsvOptions(options);
    try {
      return await CsvWorkbook.#open(source, context, settled);
    } catch (error) {
      if (error instanceof ConsultChimpsError) throw error;
      throw readFailure(context, undefined, error);
    }
  }

  static async #open(
    source: RandomAccessSource,
    context: StreamedWorkbookContext,
    options: SettledCsvOptions,
  ): Promise<CsvWorkbook> {
    const size = source.size;
    const head = await source.readAt(0, Math.min(4, size));
    const marked = byteOrderMark(head);
    const validator =
      options.encoding === undefined && marked === undefined
        ? new TextDecoder("utf-8", { fatal: true })
        : undefined;
    let valid = true;
    let zero = false;
    // FNV-1a over every byte: cheap, and any change to the bytes changes it.
    let fingerprint = 0x811c9dc5;
    for (let offset = 0; offset < size; offset += CSV_PIECE_BYTES) {
      const bytes = await source.readAt(
        offset,
        Math.min(CSV_PIECE_BYTES, size - offset),
      );
      for (const byte of bytes) {
        fingerprint = Math.imul(fingerprint ^ byte, 0x01000193) >>> 0;
        if (byte === 0) zero = true;
      }
      if (validator !== undefined && valid) {
        try {
          validator.decode(bytes, { stream: true });
        } catch {
          valid = false;
        }
      }
    }
    if (validator !== undefined && valid) {
      try {
        validator.decode();
      } catch {
        valid = false;
      }
    }

    let encoding: CsvEncoding;
    let encodingSource: CsvEncodingSource;
    if (options.encoding !== undefined) {
      encoding = options.encoding;
      encodingSource = "chosen";
    } else if (marked !== undefined) {
      encoding = marked;
      encodingSource = "byte-order-mark";
    } else if (zero) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_CSV_ENCODING_UNKNOWN,
        `${context.source} holds zero bytes and has no byte order mark, so its encoding is unknown; it is most likely UTF-16. Choose its encoding and run again.`,
        { details: context.details },
      );
    } else if (valid) {
      encoding = "utf-8";
      encodingSource = "valid-utf-8";
    } else {
      encoding = "windows-1252";
      encodingSource = "fallback";
    }

    // The first piece decides the line ending and, unless one was chosen, the
    // delimiter, from its first ten rows that are not blank.
    const sample = new TextDecoder(encoding).decode(
      await source.readAt(0, Math.min(CSV_PIECE_BYTES, size)),
      { stream: true },
    );
    const newline =
      sample.includes("\n") || !sample.includes("\r") ? "\n" : "\r";
    let delimiter = options.delimiter;
    if (delimiter === undefined) {
      const end = sample.lastIndexOf(newline);
      const guessed = Papa.parse<string[]>(
        end > 0 ? sample.slice(0, end + 1) : sample,
        {
          delimitersToGuess: GUESSED_DELIMITERS,
          newline,
          skipEmptyLines: "greedy",
          preview: 10,
        },
      ).meta.delimiter;
      delimiter = GUESSED_DELIMITERS.includes(guessed) ? guessed : ",";
    }
    return new CsvWorkbook(source, context, options, {
      encoding,
      encodingSource,
      delimiter,
      newline,
      fingerprint,
    });
  }

  /** A CSV file has no shared strings; present so it reads like a workbook. */
  loadStrings(): Promise<void> {
    return Promise.resolve();
  }

  releaseStrings(): void {}

  /**
   * Read the one worksheet into `consumer`, rows in file order. `text` gives
   * each cell the field as written; `occupancy` passes each row's columns and
   * gives a converted date the field it was read from as `stored`. The other
   * options a workbook read takes change nothing here: a CSV has no declared
   * range to clip to, and its rows always arrive in order.
   */
  async readWorksheet(
    sheet: StreamedSheet,
    consumer: WorksheetConsumer,
    options: {
      gather?: boolean;
      clip?: boolean;
      text?: boolean;
      occupancy?: boolean;
      between?: () => Promise<void>;
    } = {},
  ): Promise<WorksheetRead> {
    try {
      return await this.#read(consumer, options);
    } catch (error) {
      if (error instanceof ConsultChimpsError) throw error;
      throw readFailure(this.#context, sheet.name, error);
    }
  }

  async #read(
    consumer: WorksheetConsumer,
    options: {
      text?: boolean;
      occupancy?: boolean;
      between?: () => Promise<void>;
    },
  ): Promise<WorksheetRead> {
    const withText = options.text === true;
    const occupancy = options.occupancy === true;
    const numbers =
      this.#options.numbers === undefined
        ? undefined
        : new CsvNumberReader(
            this.#options.numbers.decimal,
            this.#options.numbers.thousands,
          );
    const dates = this.#options.dates;
    const stripCarriageReturn = this.#newline === "\n";
    let range: CellRectangle | undefined;
    let row = 0;
    let failure: unknown;

    const cellOf = (column: number, field: string): StreamedCell => {
      let value: StreamedValue = field;
      let stored: StreamedValue | undefined;
      const number = numbers?.read(field);
      if (number !== undefined) {
        value = number;
      } else if (dates !== undefined) {
        const date = csvDate(field, dates);
        if (date !== undefined) {
          value = date;
          stored = field;
        }
      }
      if (withText && occupancy && stored !== undefined) {
        return { column, value, text: field, stored };
      }
      if (withText) return { column, value, text: field };
      if (occupancy && stored !== undefined) return { column, value, stored };
      return { column, value };
    };

    const deliver = (fields: string[]): void => {
      const index = row;
      row += 1;
      const last = fields.length - 1;
      const ending = fields[last];
      if (stripCarriageReturn && ending?.endsWith("\r") === true) {
        fields[last] = ending.slice(0, -1);
      }
      let cells: StreamedCell[] | undefined;
      for (let column = 0; column < fields.length; column += 1) {
        const field = fields[column]!;
        if (field === "") continue;
        (cells ??= []).push(cellOf(column, field));
      }
      if (cells === undefined) return;
      const first = cells[0]!.column;
      const end = cells[cells.length - 1]!.column;
      range =
        range === undefined
          ? {
              startRow: index,
              endRow: index,
              startColumn: first,
              endColumn: end,
            }
          : {
              startRow: range.startRow,
              endRow: index,
              startColumn: Math.min(range.startColumn, first),
              endColumn: Math.max(range.endColumn, end),
            };
      if (occupancy) {
        consumer.row(
          index,
          cells,
          cells.map((cell) => cell.column),
        );
      } else {
        consumer.row(index, cells);
      }
    };

    const malformed = (rowNumber: number, why: string): ConsultChimpsError =>
      new ConsultChimpsError(
        XLSX_ERRORS.XLSX_CSV_MALFORMED,
        `${this.#context.source} could not be read as CSV: row ${String(rowNumber)} ${why}. Open the file, correct the quotes on that row, save it, and run again.`,
        { details: { ...this.#context.details, row: rowNumber } },
      );

    consumer.begin();
    const feed = new TextFeed();
    Papa.parse<string[]>(feed as unknown as NodeJS.ReadableStream, {
      delimiter: this.delimiter,
      newline: this.#newline,
      quoteChar: '"',
      escapeChar: '"',
      header: false,
      dynamicTyping: false,
      skipEmptyLines: false,
      chunk: (results, parser) => {
        if (failure !== undefined) return;
        try {
          const error = results.errors[0];
          if (error !== undefined) {
            throw malformed(
              row + (error.row ?? 0) + 1,
              error.code === "MissingQuotes"
                ? "opens a quoted field that is never closed"
                : "has a quoted field with text after its closing quote",
            );
          }
          for (const fields of results.data) deliver(fields);
        } catch (error) {
          failure = error;
          parser.abort();
        }
      },
      // Called once Papa has parsed the final piece; the rows are all delivered.
      complete: () => undefined,
      error: (error) => {
        failure ??= error;
      },
    });

    let open = 0;
    const take = (text: string): void => {
      const before = row;
      feed.data(text);
      if (row !== before) {
        open = 0;
        return;
      }
      open += text.length;
      if (open > CSV_MAX_OPEN_ROW_CHARS && failure === undefined) {
        failure = malformed(
          row + 1,
          "runs past 16 MB of text without ending, so it most likely opens a quoted field that is never closed",
        );
      }
    };

    const decoder = new TextDecoder(this.encoding);
    const size = this.#source.size;
    for (let offset = 0; offset < size; offset += CSV_PIECE_BYTES) {
      const bytes = await this.#source.readAt(
        offset,
        Math.min(CSV_PIECE_BYTES, size - offset),
      );
      const text = decoder.decode(bytes, { stream: true });
      if (text !== "") take(text);
      if (failure !== undefined) throw failure;
      await options.between?.();
    }
    const tail = decoder.decode();
    if (tail !== "") take(tail);
    if (failure === undefined) feed.end();
    if (failure !== undefined) throw failure;

    return {
      range,
      merges: [],
      gathered: false,
      uncachedFormulas: [],
      ...(occupancy ? { occupied: range !== undefined } : {}),
    };
  }
}
