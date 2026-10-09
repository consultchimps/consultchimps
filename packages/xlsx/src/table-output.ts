/**
 * L2: where a table operation writes its table, as a workbook or as CSV (ADR
 * 0007). Both writers take the same rows, so an operation picks one and writes.
 */
import { CsvTableWriter } from "./csv/writer.js";
import {
  TableWorkbookWriter,
  type TableWorkbookWriterOptions,
  type WritableCellValue,
} from "./package/table-writer.js";

/** The file a table operation writes. */
export type TableOutputFormat =
  { readonly kind: "xlsx" } | { readonly kind: "csv"; readonly bom: boolean };

export const XLSX_TABLE_OUTPUT: TableOutputFormat = { kind: "xlsx" };

/** What a table operation writes its rows through. */
export interface TableRowWriter {
  writeRow(values: readonly WritableCellValue[]): void;
  finish(): void;
}

/**
 * A writer for `format`. A CSV file has no worksheet, widths or declared row
 * count, so it uses only the columns and the chunk receiver.
 */
export function openTableWriter(
  format: TableOutputFormat,
  options: TableWorkbookWriterOptions,
): TableRowWriter {
  if (format.kind === "xlsx") return new TableWorkbookWriter(options);
  return new CsvTableWriter({
    columns: options.columns,
    header: options.header ?? true,
    bom: format.bom,
    rowCount: options.rowCount,
    onChunk: options.onChunk,
  });
}
