/**
 * L3: opening an input for the operations that read streamed worksheets, a
 * workbook or a CSV file alike (ADR 0007).
 */
import type { RandomAccessSource } from "@consultchimps/core";

import { CsvWorkbook, isCsvName } from "../csv/reader.js";
import type { CsvReadOptions } from "../csv/options.js";
import {
  StreamedWorkbook,
  type SheetBook,
  type StreamedWorkbookContext,
} from "./consolidate/reader.js";

/**
 * Open `source` by its name: a `.csv` file as a workbook of one worksheet, read
 * with `csv`, and anything else as a workbook.
 */
export function openSheetBook(
  source: RandomAccessSource,
  context: StreamedWorkbookContext,
  csv?: CsvReadOptions,
): Promise<SheetBook> {
  return isCsvName(context.file)
    ? CsvWorkbook.open(source, context, csv)
    : StreamedWorkbook.open(source, context);
}
