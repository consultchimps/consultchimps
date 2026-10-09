/**
 * L3: a CSV file written out as a workbook of one worksheet (ADR 0007), for
 * the operations that copy whole worksheets rather than read their rows, such
 * as a merge. The grid is copied as it is: every row from row 1, no header
 * row, no AutoFilter, values typed as the CSV options ask and dates written as
 * Excel dates.
 */
import {
  ConsultChimpsError,
  type RandomAccessSource,
} from "@consultchimps/core";

import type { ByteSink } from "../bytes.js";
import { XLSX_ERRORS } from "../errors.js";
import { writableCellValue } from "../model/date-cells.js";
import {
  cellWidthLength,
  tableColumnWidth,
  TableWorkbookWriter,
  type WritableCellValue,
} from "../package/table-writer.js";
import type { CsvReadOptions } from "./options.js";
import { CsvWorkbook } from "./reader.js";

/** Where a converted workbook is kept until the operation has read it. */
export interface CsvScratchFile {
  /** Receives the workbook's bytes in order. */
  readonly sink: ByteSink;
  /** Called once every byte is written; gives the workbook back to read. */
  finish(): Promise<RandomAccessSource>;
}

/** Make a place for one converted workbook, named `name`. */
export type CsvScratch = (name: string) => Promise<CsvScratchFile>;

/** Keep converted workbooks in memory, for a caller with nowhere else. */
export const memoryCsvScratch: CsvScratch = (name) => {
  const chunks: Uint8Array[] = [];
  return Promise.resolve({
    sink: {
      write: (chunk) => {
        chunks.push(chunk);
      },
      flush: () => Promise.resolve(),
      abort: () => {
        chunks.length = 0;
        return Promise.resolve();
      },
    },
    finish: () => {
      let size = 0;
      for (const chunk of chunks) size += chunk.length;
      const bytes = new Uint8Array(size);
      let offset = 0;
      for (const chunk of chunks) {
        bytes.set(chunk, offset);
        offset += chunk.length;
      }
      chunks.length = 0;
      return Promise.resolve({
        name,
        size,
        readAt: (start: number, length: number) =>
          Promise.resolve(bytes.subarray(start, start + length)),
      });
    },
  });
};

const WORKSHEET_ROWS = 1_048_576;
const WORKSHEET_COLUMNS = 16_384;

export interface CsvAsWorkbook {
  /** The workbook, named as the CSV file is so the operation reports it so. */
  readonly workbook: RandomAccessSource;
  /** What opening the CSV file noticed, such as an encoding fallback. */
  readonly warnings: readonly string[];
}

/**
 * Write `source`, a CSV file, as a one-worksheet workbook into `scratch`, and
 * hand it back to be read. The file is read twice, in pieces: once to size
 * the grid, once to write it.
 */
export async function csvAsWorkbook(
  source: RandomAccessSource,
  csv: CsvReadOptions | undefined,
  scratch: CsvScratch,
  between?: () => Promise<void>,
): Promise<CsvAsWorkbook> {
  const book = await CsvWorkbook.open(
    source,
    {
      file: source.name,
      source: source.name,
      details: { source: source.name },
    },
    csv,
  );
  const sheet = book.sheets[0]!;

  let rowCount = 0;
  const lengths: number[] = [];
  await book.readWorksheet(
    sheet,
    {
      begin: () => {
        rowCount = 0;
        lengths.length = 0;
      },
      row: (row, cells) => {
        rowCount = row + 1;
        for (const cell of cells) {
          const length = cellWidthLength(
            writableCellValue(cell.value as WritableCellValue),
          );
          lengths[cell.column] = Math.max(lengths[cell.column] ?? 0, length);
        }
      },
    },
    between === undefined ? {} : { between },
  );
  if (rowCount > WORKSHEET_ROWS || lengths.length > WORKSHEET_COLUMNS) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_OUTPUT_TOO_LARGE,
      `${source.name} holds ${rowCount.toLocaleString("en-US")} rows and ${lengths.length.toLocaleString("en-US")} columns, more than an Excel worksheet holds (1,048,576 rows and 16,384 columns), so it cannot become a worksheet.`,
      {
        details: {
          source: source.name,
          rows: rowCount,
          columns: lengths.length,
        },
      },
    );
  }

  const columns = Math.max(lengths.length, 1);
  const widths: number[] = [];
  for (let column = 0; column < columns; column += 1) {
    widths.push(tableColumnWidth(lengths[column] ?? 0));
  }
  const file = await scratch(source.name);
  try {
    const writer = new TableWorkbookWriter({
      sheetName: sheet.name,
      columns: new Array<string>(columns).fill(""),
      widths,
      rowCount,
      header: false,
      onChunk: (chunk) => {
        file.sink.write(chunk);
      },
    });
    let written = 0;
    const empty: WritableCellValue[] = [];
    await book.readWorksheet(
      sheet,
      {
        begin: () => {
          if (written > 0) {
            throw new Error("The CSV file was read again from its start.");
          }
        },
        row: (row, cells) => {
          // Blank lines are empty rows, so a row keeps its number.
          while (written < row) {
            writer.writeRow(empty);
            written += 1;
          }
          const values = new Array<WritableCellValue>(columns).fill(null);
          for (const cell of cells) {
            values[cell.column] = writableCellValue(
              cell.value as WritableCellValue,
            );
          }
          writer.writeRow(values);
          written += 1;
        },
      },
      {
        between: async () => {
          await file.sink.flush();
          await between?.();
        },
      },
    );
    if (written !== rowCount) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_READ_FAILED,
        `${source.name} changed while it was being read, so it was not merged. Run again once the file is no longer being changed.`,
        { details: { source: source.name } },
      );
    }
    writer.finish();
    await file.sink.flush();
    return { workbook: await file.finish(), warnings: book.warnings };
  } catch (error) {
    await file.sink.abort().catch(() => undefined);
    throw error;
  }
}
