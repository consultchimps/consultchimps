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

import { CellDate } from "../package/cell-date.js";
import { CellError } from "../package/cell-error.js";
import type { WritableCellValue } from "../package/table-writer.js";

export const CSV_MEDIA_TYPE = "text/csv";
export const CSV_EXTENSION = ".csv";

/** The rows formatted before Papa quotes them as one batch. */
const BATCH_ROWS = 1000;

/**
 * Text a spreadsheet would read as the start of a formula: `=`, `+`, `-`, `@`,
 * a tab or a carriage return. Text that is only a signed number, or only
 * dashes, cannot run anything and is left as it is, so CSV to CSV keeps it.
 */
const FORMULA_START = /^[=+\-@\t\r]/u;
const HARMLESS = /^(?:[+-]?[0-9][0-9.,]*|-+)$/u;

/** A text value as CSV output writes it, with the formula guard applied. */
export function guardedCsvText(text: string): string {
  return FORMULA_START.test(text) && !HARMLESS.test(text) ? `'${text}` : text;
}

/** A date as ISO text: the day alone at midnight, else with its time. */
function isoDate(date: CellDate): string {
  // The table text is the workbook date spelling, yyyy-mm-ddThh:mm:ss.sssZ.
  const day = date.text.slice(0, 10);
  if (!date.time) return day;
  const time = date.text.slice(11, 19);
  const milliseconds = date.text.slice(20, 23);
  return milliseconds === "000"
    ? `${day}T${time}`
    : `${day}T${time}.${milliseconds}`;
}

/** One value as CSV text. */
export function csvCellText(value: WritableCellValue): string {
  if (value === null) return "";
  if (value instanceof CellDate) return isoDate(value);
  if (value instanceof CellError) return value.text;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") {
    return Number.isFinite(value) ? String(value) : "";
  }
  return guardedCsvText(value);
}

export interface CsvTableWriterOptions {
  /** The header row, written first unless `header` is false. */
  columns: readonly string[];
  /** Write `columns` as the first row. On by default. */
  header?: boolean | undefined;
  /** Start with a UTF-8 byte order mark. On by default. */
  bom?: boolean | undefined;
  /** Receives the file's bytes in order as they are produced. */
  onChunk: (chunk: Uint8Array) => void;
}

/** Writes a table as CSV, a batch of rows at a time. */
export class CsvTableWriter {
  readonly #columns: number;
  readonly #onChunk: (chunk: Uint8Array) => void;
  readonly #encoder = new TextEncoder();
  #batch: string[][] = [];
  #finished = false;

  constructor(options: CsvTableWriterOptions) {
    this.#columns = options.columns.length;
    this.#onChunk = options.onChunk;
    if (options.bom ?? true) {
      this.#onChunk(Uint8Array.from([0xef, 0xbb, 0xbf]));
    }
    if (options.header ?? true) {
      this.#batch.push(options.columns.map((name) => guardedCsvText(name)));
    }
  }

  writeRow(values: readonly WritableCellValue[]): void {
    if (this.#finished) throw new Error("The CSV file is already finished.");
    const row: string[] = [];
    for (let index = 0; index < this.#columns; index += 1) {
      row.push(csvCellText(values[index] ?? null));
    }
    this.#batch.push(row);
    if (this.#batch.length >= BATCH_ROWS) this.#flush();
  }

  finish(): void {
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
      quotes: false,
      // The guard is applied above, so Papa's own, which would also quote
      // signed numbers, stays off.
      escapeFormulae: false,
    });
    this.#batch = [];
    this.#onChunk(this.#encoder.encode(`${text}\r\n`));
  }
}
