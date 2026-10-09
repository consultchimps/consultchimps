/**
 * L2: a table written as CSV (ADR 0007), row by row, so the output is never
 * held whole. Papa Parse quotes each batch of rows; the values are formatted
 * here first, and the batches joined here.
 *
 * - RFC 4180 quoting, comma delimiters, CRLF after every row, the last too.
 * - UTF-8, with a byte order mark unless `bom` is false, so Excel opens it as
 *   UTF-8.
 * - Text a spreadsheet would run as a formula gets a leading apostrophe.
 * - Dates as ISO text, numbers in JavaScript's shortest round-trip form,
 *   booleans as TRUE and FALSE, error cells as their text.
 */
import Papa from "papaparse";

import { calendarIsoParts } from "../model/calendar.js";
import { CellDate } from "../package/cell-date.js";
import { CellError } from "../package/cell-error.js";
import type { WritableCellValue } from "../package/table-writer.js";

export const CSV_MEDIA_TYPE = "text/csv";
export const CSV_EXTENSION = ".csv";

/** The rows, or characters, formatted before Papa quotes them as one batch. */
const BATCH_ROWS = 1000;
const BATCH_CHARACTERS = 1024 * 1024;

/**
 * Text a spreadsheet would read as the start of a formula: `=`, `+`, `-`, `@`,
 * a tab or a carriage return. Text that is only a signed number, or only
 * dashes, cannot run anything and is left as it is, so CSV to CSV keeps it.
 */
const FORMULA_START = /^\s*[=+\-@]|^[\t\r]/u;
const HARMLESS = /^(?:[+-]?[0-9][0-9.,]*|-+)$/u;

/**
 * A text value as CSV output writes it, with the formula guard applied. The
 * guard looks past leading spaces, which some spreadsheets trim first.
 */
export function guardedCsvText(text: string): string {
  if (!FORMULA_START.test(text)) return text;
  const trimmed = text.trim();
  return HARMLESS.test(trimmed) && unchangedAsNumber(trimmed)
    ? text
    : `'${text}`;
}

/**
 * Whether a spreadsheet reading signed numeric text as a number keeps every
 * digit: no leading zero before another digit, and at most 15 significant
 * digits. Otherwise the apostrophe keeps it as written.
 */
function unchangedAsNumber(text: string): boolean {
  if (/^-+$/u.test(text)) return true;
  const digits = text.replace(/[^0-9]/gu, "");
  if (/^[+-]?0[0-9]/u.test(text)) return false;
  return digits.replace(/^0+/u, "").length <= 15;
}

/**
 * A date in the workbook date spelling, yyyy-mm-ddThh:mm:ss.sssZ, as ISO
 * text: the day alone at midnight, else with its time.
 */
function isoDate(text: string): string {
  const day = text.slice(0, 10);
  const time = text.slice(11, 19);
  const milliseconds = text.slice(20, 23);
  if (time === "00:00:00" && milliseconds === "000") return day;
  return milliseconds === "000"
    ? `${day}T${time}`
    : `${day}T${time}.${milliseconds}`;
}

/** One value as CSV text. */
export function csvCellText(value: WritableCellValue): string {
  if (value === null) return "";
  if (value instanceof CellDate) return isoDate(value.text);
  if (value instanceof CellError) return value.text;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") {
    return Number.isFinite(value) ? String(value) : "";
  }
  // A date a workbook cannot count, before 1900, stays in the date spelling.
  if (calendarIsoParts(value) !== undefined) return isoDate(value);
  return guardedCsvText(value);
}

export interface CsvTableWriterOptions {
  /** The header row, written first unless `header` is false. */
  columns: readonly string[];
  /** Write `columns` as the first row. On by default. */
  header?: boolean | undefined;
  /** Start with a UTF-8 byte order mark. On by default. */
  bom?: boolean | undefined;
  /**
   * The data rows that will be written, not counting the header; checked as
   * the workbook writer checks it, so a source that changed between reads
   * fails rather than giving a short file. Unchecked when absent.
   */
  rowCount?: number | undefined;
  /** Receives the file's bytes in order as they are produced. */
  onChunk: (chunk: Uint8Array) => void;
}

/** Writes a table as CSV, a batch of rows at a time. */
export class CsvTableWriter {
  readonly #columns: number;
  readonly #onChunk: (chunk: Uint8Array) => void;
  readonly #encoder = new TextEncoder();
  readonly #rowCount: number | undefined;
  #batch: string[][] = [];
  #batchCharacters = 0;
  #written = 0;
  #finished = false;

  constructor(options: CsvTableWriterOptions) {
    this.#columns = options.columns.length;
    this.#onChunk = options.onChunk;
    this.#rowCount = options.rowCount;
    if (options.bom ?? true) {
      this.#onChunk(Uint8Array.from([0xef, 0xbb, 0xbf]));
    }
    if (options.header ?? true) {
      this.#batch.push(options.columns.map((name) => guardedCsvText(name)));
    }
  }

  writeRow(values: readonly WritableCellValue[]): void {
    if (this.#finished) throw new Error("The CSV file is already finished.");
    if (this.#rowCount !== undefined && this.#written >= this.#rowCount) {
      throw new Error(
        `More rows were written than the ${this.#rowCount} declared.`,
      );
    }
    this.#written += 1;
    const row: string[] = [];
    for (let index = 0; index < this.#columns; index += 1) {
      const text = csvCellText(values[index] ?? null);
      this.#batchCharacters += text.length;
      row.push(text);
    }
    this.#batch.push(row);
    if (
      this.#batch.length >= BATCH_ROWS ||
      this.#batchCharacters >= BATCH_CHARACTERS
    ) {
      this.#flush();
    }
  }

  finish(): void {
    if (this.#rowCount !== undefined && this.#written !== this.#rowCount) {
      throw new Error(
        `${this.#written} rows were written but ${this.#rowCount} were declared.`,
      );
    }
    this.#flush();
    this.#finished = true;
  }

  #flush(): void {
    if (this.#batch.length === 0) return;
    const text = Papa.unparse(this.#batch, {
      delimiter: ",",
      newline: "\r\n",
      quoteChar: '"',
      escapeChar: '"',
      header: false,
      // A lone empty field is quoted, so a one-column row is not a blank
      // line, which many readers skip.
      quotes: this.#columns === 1 ? (value: unknown) => value === "" : false,
      // The guard is applied above, so Papa's own, which would also quote
      // signed numbers, stays off.
      escapeFormulae: false,
    });
    this.#batch = [];
    this.#batchCharacters = 0;
    this.#onChunk(this.#encoder.encode(`${text}\r\n`));
  }
}
