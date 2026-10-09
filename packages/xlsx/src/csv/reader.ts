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
 * so files mixing CRLF and LF read correctly; a file whose rows end in CR alone
 * is split on CR.
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
  SheetBook,
  StreamedCell,
  StreamedDefinedName,
  StreamedSheet,
  StreamedValue,
  StreamedWorkbookContext,
  WorksheetConsumer,
  WorksheetRead,
  WorksheetReadOptions,
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

/** A carriage return that does not start a CRLF pair. */
const LONE_CARRIAGE_RETURN = /\r(?!\n)/u;

/** The part name a CSV's one worksheet carries; it names no zip entry. */
const CSV_PART = "csv";

/** Whether a file name names a CSV file. */
export function isCsvName(name: string): boolean {
  return /\.csv$/iu.test(name);
}

/** Characters a worksheet name cannot hold. */
// eslint-disable-next-line no-control-regex -- a sheet name cannot hold control characters; the pattern finds them so they can be replaced.
const SHEET_NAME_FORBIDDEN = /[\\/?*[\]:\x00-\x1f\x7f]/gu;
const SHEET_NAME_LIMIT = 31;

/** `text` without the apostrophes at either end, which a sheet name cannot hold. */
function withoutOuterApostrophes(text: string): string {
  return text.replace(/^'+|'+$/gu, "");
}

/**
 * The worksheet a CSV file reads as: the file's name without `.csv`, made a
 * legal sheet name the same way everywhere. A character Excel forbids becomes
 * `_`, apostrophes at either end are dropped, the name is cut to 31
 * characters, `History`, which Excel reserves, becomes `History_`, and an empty
 * name becomes `Sheet1`.
 */
export function csvSheetName(file: string): string {
  const base = file.split(/[\\/]/u).pop() ?? file;
  let name = withoutOuterApostrophes(
    base.replace(/\.csv$/iu, "").replace(SHEET_NAME_FORBIDDEN, "_"),
  );
  if (name.length > SHEET_NAME_LIMIT) {
    // Never between the two halves of a character outside the basic plane.
    const high = name.charCodeAt(SHEET_NAME_LIMIT - 1);
    const cut =
      high >= 0xd800 && high <= 0xdbff
        ? SHEET_NAME_LIMIT - 1
        : SHEET_NAME_LIMIT;
    name = withoutOuterApostrophes(name.slice(0, cut));
  }
  if (name.toLowerCase() === "history") return `${name}_`;
  return name === "" ? "Sheet1" : name;
}

/** How the encoding a read uses was settled. */
export type CsvEncodingSource =
  "chosen" | "byte-order-mark" | "valid-utf-8" | "fallback";

function byteOrderMark(head: Uint8Array): CsvEncoding | "utf-32" | undefined {
  if (head[0] === 0xef && head[1] === 0xbb && head[2] === 0xbf) return "utf-8";
  // UTF-32 LE's mark begins with UTF-16 LE's, so it is told apart first.
  if (
    (head[0] === 0xff && head[1] === 0xfe && head[2] === 0 && head[3] === 0) ||
    (head[0] === 0 && head[1] === 0 && head[2] === 0xfe && head[3] === 0xff)
  ) {
    return "utf-32";
  }
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

/**
 * How a file's rows end: on LF, unless the first line ending in the file is a
 * CR that no LF follows. The file is read as far as that first line ending, so
 * a first row longer than a piece, or a CRLF split between two pieces, cannot
 * mislead it. A CR file split on LF by mistake is refused, never read as one
 * row, because its lone carriage returns are refused.
 */
async function firstLineEnding(
  source: RandomAccessSource,
  encoding: CsvEncoding,
): Promise<"\n" | "\r"> {
  const decoder = new TextDecoder(encoding);
  let afterCarriageReturn = false;
  for (let offset = 0; offset < source.size; offset += CSV_PIECE_BYTES) {
    const text = decoder.decode(
      await source.readAt(
        offset,
        Math.min(CSV_PIECE_BYTES, source.size - offset),
      ),
      { stream: true },
    );
    for (let index = 0; index < text.length; index += 1) {
      const code = text.charCodeAt(index);
      if (afterCarriageReturn) return code === 10 ? "\n" : "\r";
      if (code === 10) return "\n";
      if (code === 13) afterCarriageReturn = true;
    }
  }
  return "\n";
}

/**
 * Papa's guess between the candidate delimiters over the first ten rows that
 * are not blank, or, when those hold none of the candidates, as title lines
 * above a table may not, over the first ten rows that hold one.
 */
function guessDelimiter(sample: string, newline: "\n" | "\r"): string {
  const guess = (text: string): string | undefined => {
    const result = Papa.parse<string[]>(text, {
      delimitersToGuess: GUESSED_DELIMITERS,
      newline,
      skipEmptyLines: "greedy",
      preview: 10,
    });
    return result.errors.some((error) => error.code === "UndetectableDelimiter")
      ? undefined
      : result.meta.delimiter;
  };
  const first = guess(sample);
  if (first !== undefined && GUESSED_DELIMITERS.includes(first)) return first;
  const rows = sample
    .split(newline)
    .filter((row) =>
      GUESSED_DELIMITERS.some((candidate) => row.includes(candidate)),
    );
  const second = rows.length === 0 ? undefined : guess(rows.join(newline));
  return second !== undefined && GUESSED_DELIMITERS.includes(second)
    ? second
    : ",";
}

/** A CSV file opened for reading as one worksheet. */
export class CsvWorkbook implements SheetBook {
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
      overruled: CsvEncoding | undefined;
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
    const warnings: string[] = [];
    if (settled.encodingSource === "fallback") {
      warnings.push(
        `${context.file} is not valid UTF-8 and has no byte order mark, so it was read as Windows-1252. If its accented letters look wrong, choose its encoding and run again.`,
      );
    }
    if (settled.overruled !== undefined) {
      const marked = csvEncodingName(settled.encoding);
      warnings.push(
        `${context.file} starts with a ${marked} byte order mark, so it was read as ${marked} rather than the chosen ${csvEncodingName(settled.overruled)}.`,
      );
    }
    this.warnings = warnings;
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
    const mark = byteOrderMark(head);
    if (mark === "utf-32") {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_CSV_ENCODING_UNKNOWN,
        `${context.source} starts with a UTF-32 byte order mark, and UTF-32 text is not read. Save it as UTF-8 and run again.`,
        { details: context.details },
      );
    }
    const marked = mark;
    // Every byte is checked against the encoding the file will be read in, so
    // a malformed sequence is refused rather than read as a replacement
    // character: the mark's, else a chosen Unicode encoding, else UTF-8 to
    // detect it. Windows-1252 gives every byte a character.
    const declared =
      marked ??
      (options.encoding === "windows-1252" ? undefined : options.encoding);
    const validator =
      declared !== undefined
        ? new TextDecoder(declared, { fatal: true })
        : options.encoding === undefined
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

    if (declared !== undefined && !valid) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_CSV_ENCODING_UNKNOWN,
        `${context.source} ${marked === undefined ? "was to be read" : "starts with a byte order mark saying it is"} ${csvEncodingName(declared)}, but holds bytes that are not ${csvEncodingName(declared)}, so its text cannot be read without changing it. ${marked === undefined ? "Choose its encoding" : "Save it again as UTF-8"} and run again.`,
        { details: context.details },
      );
    }
    let encoding: CsvEncoding;
    let encodingSource: CsvEncodingSource;
    // A byte order mark says what the file is, so it wins over a choice that
    // contradicts it, which would otherwise read the mark into the first cell.
    let overruled: CsvEncoding | undefined;
    if (marked !== undefined) {
      encoding = marked;
      encodingSource = "byte-order-mark";
      if (options.encoding !== undefined && options.encoding !== marked) {
        overruled = options.encoding;
      }
    } else if (options.encoding !== undefined) {
      encoding = options.encoding;
      encodingSource = "chosen";
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

    const newline = await firstLineEnding(source, encoding);
    let delimiter = options.delimiter;
    if (delimiter === undefined) {
      // The first piece's first ten rows that are not blank decide it.
      const sample = new TextDecoder(encoding).decode(
        await source.readAt(0, Math.min(CSV_PIECE_BYTES, size)),
        { stream: true },
      );
      const end = sample.lastIndexOf(newline);
      delimiter = guessDelimiter(
        end > 0 ? sample.slice(0, end + 1) : sample,
        newline,
      );
    }
    return new CsvWorkbook(source, context, options, {
      encoding,
      encodingSource,
      delimiter,
      newline,
      fingerprint,
      overruled,
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
    options: WorksheetReadOptions = {},
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
        if (stripCarriageReturn && LONE_CARRIAGE_RETURN.test(field)) {
          // In a file split on LF, a CR on its own is either a line ending
          // Papa did not split on or text inside quotes, and nothing tells the
          // two apart; reading it either way could join two rows.
          throw malformed(
            index + 1,
            "holds a carriage return on its own, which cannot be told apart from a line ending",
            "Save the file again with one kind of line ending, and run again.",
          );
        }
        if (!stripCarriageReturn && field.includes("\n")) {
          // In a file split on CR, an LF is the same question the other way
          // round: a CRLF row ending, or text inside quotes.
          throw malformed(
            index + 1,
            "holds a line feed in a file whose rows end in carriage returns, which cannot be told apart from a line ending",
            "Save the file again with one kind of line ending, and run again.",
          );
        }
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

    const malformed = (
      rowNumber: number,
      why: string,
      fix = "Open the file, correct the quotes on that row, save it, and run again.",
    ): ConsultChimpsError =>
      new ConsultChimpsError(
        XLSX_ERRORS.XLSX_CSV_MALFORMED,
        `${this.#context.source} could not be read as CSV: row ${String(rowNumber)} ${why}. ${fix}`,
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
          // An error in the row a piece ends inside is Papa reading half a row,
          // such as a closing quote whose line ending is in the next piece.
          // That row is parsed again whole with the next piece, so only errors
          // in the rows delivered here count.
          const error = results.errors.find(
            (candidate) => (candidate.row ?? 0) < results.data.length,
          );
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
          "is longer than 16 MB, the most one row may hold here; a row that long is most often a quote that opens a field and is never closed",
        );
      }
    };

    const decoder = new TextDecoder(this.encoding);
    const size = this.#source.size;
    // The bytes are fingerprinted again as they are read, so a file changed
    // since it was opened fails the read rather than giving other rows.
    let fingerprint = 0x811c9dc5;
    for (let offset = 0; offset < size; offset += CSV_PIECE_BYTES) {
      const bytes = await this.#source.readAt(
        offset,
        Math.min(CSV_PIECE_BYTES, size - offset),
      );
      for (const byte of bytes) {
        fingerprint = Math.imul(fingerprint ^ byte, 0x01000193) >>> 0;
      }
      const text = decoder.decode(bytes, { stream: true });
      if (text !== "") take(text);
      if (failure !== undefined) throw failure;
      await options.between?.();
    }
    const tail = decoder.decode();
    if (tail !== "") take(tail);
    if (failure === undefined) feed.end();
    if (failure !== undefined) throw failure;
    if (fingerprint !== this.fingerprint) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_READ_FAILED,
        `${this.#context.source} changed while it was being read, so nothing read from it was used. Run again once the file is no longer being changed.`,
        { details: this.#context.details },
      );
    }

    return {
      range,
      merges: [],
      gathered: false,
      uncachedFormulas: [],
      ...(occupancy ? { occupied: range !== undefined } : {}),
    };
  }
}
