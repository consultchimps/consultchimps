/**
 * L3: consolidation in two passes over streamed workbooks (ADR 0006, #225).
 *
 * The first pass reads every selected worksheet without keeping its rows: it
 * counts what each row holds, finds the header row and the spacer columns by
 * the rules in `src/region/header-detection.ts`, names the columns, and keeps
 * per column only how wide its values run and, when a mapping coerces dates,
 * the first value that is not text. From that alone it settles the union of
 * columns, the mapping, its warnings and refusals, the drafted mapping, the
 * column widths and the row count, so a run that is going to be refused is
 * refused before anything is written. A mapping that coerces values is the
 * one exception: its coercions are checked by reading the affected worksheets
 * once more before writing, since a coerced value's width is only known once
 * it is coerced.
 *
 * The second pass reads each worksheet's rows again and writes them through
 * the streaming table writer as they arrive. Memory is bounded by the largest
 * worksheet and its shared strings, not by the sum of the inputs: a worksheet
 * whose rows arrive in order is never held, and one whose rows do not is held
 * only while it is read.
 */
import {
  ConsultChimpsError,
  throwIfAborted,
  type AbortOutputContext,
  type ProgressReporter,
  type RandomAccessSource,
} from "@consultchimps/core";
import {
  applyColumnMapping,
  normalizedColumnKey,
  planTableUnion,
  uniqueHeaders,
  type CellValue,
  type ColumnMapping,
  type ColumnMappingSuggestion,
  type Table,
  type TableRow,
} from "@consultchimps/tabular";

import { XLSX_ERRORS } from "../../errors.js";
import {
  cellWidthLength,
  tableColumnWidth,
  TableWorkbookWriter,
  type WritableCellValue,
} from "../../package/table-writer.js";
import {
  cellKey,
  countTitleRows,
  detectHeaderRow,
  isBlankValue,
  settleHeaderCandidate,
  type RowValueCount,
} from "../../region/header-detection.js";
import {
  CONSOLIDATE_OPERATION,
  declaredHeaderRowIndex,
  refuseNonTextDateColumns,
  suggestMappingForTables,
  WORKBOOK_DATE_TEXT,
  yieldToEventLoop,
} from "../../shared.js";
import { writableCellValue } from "../../model/date-cells.js";
import type { CellRectangle } from "../../model/references.js";
import {
  uncachedFormulaHint,
  uncachedLocationsWithin,
} from "../../uncached-formulas.js";
import {
  StreamedWorkbook,
  type StreamedCell,
  type StreamedSheet,
  type StreamedValue,
  type WorksheetConsumer,
  type WorksheetRead,
} from "./reader.js";

/** One workbook to consolidate, opened afresh for each pass. */
export interface ConsolidationSource {
  /** The name a row records in `_source_file`. */
  readonly file: string;
  /** The workbook's name in read errors: a path, or an input name. */
  readonly source: string;
  /** The details each surface identifies its input by. */
  readonly details: Record<string, unknown>;
  /** Open the workbook's bytes; the caller closes what it opens. */
  open(): Promise<OpenedSource>;
}

export interface OpenedSource extends RandomAccessSource {
  close(): Promise<void>;
  /**
   * Refuse when the source changed while it was open, which a fingerprint
   * taken when it was opened cannot see. A file source checks its identity.
   */
  verifyUnchanged?(): Promise<void>;
}

export interface ConsolidationSettings {
  readonly headerRow?: number | undefined;
  readonly includeHiddenSheets?: boolean | undefined;
  readonly sheets?: string[] | undefined;
  readonly addSourceColumns?: boolean | undefined;
  readonly normalizeHeaders?: boolean | undefined;
  /** A mapping already validated by `validateColumnMapping`. */
  readonly mapping?: ColumnMapping | undefined;
  readonly suggestMapping?: boolean | undefined;
  readonly signal?: AbortSignal | undefined;
  readonly onProgress?: ProgressReporter | undefined;
  /** What a cancellation leaves behind: files on disk, or nothing in memory. */
  readonly outputContext: AbortOutputContext;
  /**
   * Hand the event loop back between units of work, so a Web Worker can
   * collect a cancellation (see `yieldToEventLoop`).
   */
  readonly yieldControl: boolean;
}

/** A worksheet that yields a table, as the first pass settled it. */
interface SheetTable {
  readonly input: number;
  readonly sheetIndex: number;
  readonly sheet: string;
  readonly gathered: boolean;
  /** Zero-based header row. */
  readonly headerRow: number;
  /** The kept columns, in the worksheet's zero-based numbering. */
  readonly columns: readonly number[];
  /** The kept columns' names. */
  readonly names: readonly string[];
  readonly rowCount: number;
  /** One-based number of the last data row. */
  readonly lastRowNumber: number;
  /** Per kept column, the longest value's length. */
  readonly widths: readonly number[];
  /** Per kept column, the first value a date coercion cannot read. */
  readonly dateOffences: ReadonlyArray<DateOffence | undefined>;
}

interface DateOffence {
  readonly row: number;
  readonly value: StreamedValue;
}

/** How one table's rows reach the output. */
interface TableOutput {
  readonly table: SheetTable;
  /** The table's columns after the mapping. */
  readonly mappedColumns: readonly string[];
  /** Constant values appended after the mapped source columns. */
  readonly constants: readonly CellValue[];
  /** Whether a coercion applies, so rows go through the mapping. */
  readonly coerced: boolean;
  /** Per output column, the position in the mapped columns, or -1. */
  readonly projection: readonly number[];
}

/** Everything the first pass settles, enough to write the output. */
export interface ConsolidationPlan {
  readonly outputs: readonly TableOutput[];
  readonly columns: readonly string[];
  readonly widths: readonly number[];
  readonly rowCount: number;
  readonly addSourceColumns: boolean;
  readonly unmappedColumns: readonly string[];
  readonly suggestion: ColumnMappingSuggestion | undefined;
  readonly inputTables: number;
  readonly skippedTitleRows: number;
  readonly skippedSpacerColumns: number;
  /**
   * `Sheet!B4` of every formula cell with no cached value in a table the
   * output takes rows from, or anywhere on a selected worksheet that yielded
   * no table. Such a cell comes out blank.
   */
  readonly uncachedFormulas: readonly string[];
  /**
   * Each input's size and zip directory fingerprint in the first pass, which
   * every later read must match.
   */
  readonly inputVersions: readonly string[];
}

/** Where the output's bytes go, in order. */
export interface ConsolidationSink {
  write(chunk: Uint8Array): void;
  /** Called between reads; a file sink writes what it has collected. */
  flush(): Promise<void>;
}

/** Rows a coercing mapping is applied to at a time. */
const MAPPING_BATCH_ROWS = 1000;

/** A cell value as a mapping or a table row carries it. */
function asCellValue(value: StreamedValue | null): CellValue {
  return value as CellValue;
}

/** The name a header cell gives its column, or null for a blank cell. */
function headerCellName(value: StreamedValue | undefined): string | null {
  return value === undefined || isBlankValue(value) ? null : String(value);
}

/** What the first pass learns about one worksheet. */
interface SheetOutcome {
  readonly table:
    Omit<SheetTable, "input" | "sheetIndex" | "sheet" | "gathered"> | undefined;
  readonly skippedTitleRows: number;
  readonly skippedSpacerColumns: number;
  /** The rectangle the table's cells were read from, header row included. */
  readonly region?: CellRectangle | undefined;
}

interface ColumnStats {
  width: number;
  dateOffence: DateOffence | undefined;
}

const NO_TABLE: SheetOutcome = {
  table: undefined,
  skippedTitleRows: 0,
  skippedSpacerColumns: 0,
};

/**
 * The first pass over one worksheet. Every populated row is counted; the rows
 * that could still turn out to be the header, or to sit above it, are kept
 * until the header rule can no longer change its answer, which on an ordinary
 * worksheet is a dozen rows down. Below that, rows only feed per-column
 * statistics and are dropped.
 */
class SheetProfile implements WorksheetConsumer {
  readonly #declared: number | undefined;
  readonly #trackDates: boolean;
  #counts: RowValueCount[] = [];
  #lastValueRow = new Map<number, number>();
  #held: Array<{ row: number; cells: readonly StreamedCell[]; count: number }> =
    [];
  #settled = false;
  #candidate = 0;
  #declaredCells: readonly StreamedCell[] | undefined;
  #stats = new Map<number, ColumnStats>();

  /**
   * `headerRow` is the one-based row a caller declared, if any. It is judged
   * in `finish`, as the table reader judges it, only once the worksheet turns
   * out to have a used range.
   */
  constructor(headerRow: number | undefined, trackDates: boolean) {
    this.#declared = headerRow === undefined ? undefined : headerRow - 1;
    this.#trackDates = trackDates;
  }

  begin(): void {
    this.#counts = [];
    this.#lastValueRow = new Map();
    this.#held = [];
    this.#settled = false;
    this.#candidate = 0;
    this.#declaredCells = undefined;
    this.#stats = new Map();
  }

  row(row: number, cells: readonly StreamedCell[]): void {
    let values = 0;
    for (const cell of cells) {
      if (!isBlankValue(cell.value)) {
        values += 1;
        this.#lastValueRow.set(cell.column, row);
      }
    }
    if (values === 0) return;
    this.#counts.push({ row, values, bannerValues: 0 });

    if (this.#declared !== undefined) {
      if (row === this.#declared) this.#declaredCells = cells;
      else if (row > this.#declared) this.#feed(row, cells);
      return;
    }
    if (this.#settled) {
      this.#feed(row, cells);
      return;
    }
    this.#held.push({ row, cells, count: this.#counts.length - 1 });
    // A candidate's answer is final once all the rows it is measured against
    // have arrived; the first candidate that qualifies is the one the rule
    // picks, so nothing past it can change which rows are kept below.
    const { index, settled } = settleHeaderCandidate(
      this.#counts,
      this.#candidate,
    );
    this.#candidate = index;
    this.#settled = settled;
  }

  #feed(row: number, cells: readonly StreamedCell[]): void {
    for (const cell of cells) {
      let stats = this.#stats.get(cell.column);
      if (stats === undefined) {
        stats = { width: 0, dateOffence: undefined };
        this.#stats.set(cell.column, stats);
      }
      stats.width = Math.max(stats.width, cellWidthLength(cell.value));
      if (
        this.#trackDates &&
        !(
          typeof cell.value === "string" && !WORKBOOK_DATE_TEXT.test(cell.value)
        ) &&
        (stats.dateOffence === undefined || row < stats.dateOffence.row)
      ) {
        stats.dateOffence = { row, value: cell.value };
      }
    }
  }

  finish(read: WorksheetRead, headerRow: number | undefined): SheetOutcome {
    const range = read.range;
    if (range === undefined) return NO_TABLE;
    const declared = declaredHeaderRowIndex(headerRow);
    if (
      declared !== undefined &&
      (declared < range.startRow || declared > range.endRow)
    ) {
      return NO_TABLE;
    }

    // Banners are declared after the cells, so the rows kept for the header
    // rule are only now measured for them; the rule reads no other row's.
    const banners = new Set<string>();
    for (const merge of read.merges) {
      if (merge.endColumn > merge.startColumn) {
        banners.add(cellKey(merge.startRow, merge.startColumn));
      }
    }
    if (banners.size > 0) {
      for (const held of this.#held) {
        let bannerValues = 0;
        for (const cell of held.cells) {
          if (
            !isBlankValue(cell.value) &&
            banners.has(cellKey(held.row, cell.column))
          ) {
            bannerValues += 1;
          }
        }
        const count = this.#counts[held.count]!;
        this.#counts[held.count] = { ...count, bannerValues };
      }
    }

    const header = declared ?? detectHeaderRow(this.#counts);
    if (header === undefined) return NO_TABLE;
    let headerCells = this.#declaredCells;
    for (const held of this.#held) {
      if (held.row > header) this.#feed(held.row, held.cells);
      else if (held.row === header) headerCells = held.cells;
    }

    let rowCount = 0;
    let lastRow = header;
    for (const count of this.#counts) {
      if (count.row > header) {
        rowCount += 1;
        lastRow = count.row;
      }
    }
    if (rowCount === 0) return NO_TABLE;

    const columns: number[] = [];
    for (
      let column = range.startColumn;
      column <= range.endColumn;
      column += 1
    ) {
      const last = this.#lastValueRow.get(column);
      if (last !== undefined && last >= header) columns.push(column);
    }
    const headerValues = new Map<number, StreamedValue>();
    for (const cell of headerCells ?? [])
      headerValues.set(cell.column, cell.value);
    return {
      table: {
        headerRow: header,
        columns,
        names: uniqueHeaders(
          columns.map((column) => headerCellName(headerValues.get(column))),
        ),
        rowCount,
        lastRowNumber: lastRow + 1,
        widths: columns.map((column) => this.#stats.get(column)?.width ?? 0),
        dateOffences: columns.map(
          (column) => this.#stats.get(column)?.dateOffence,
        ),
      },
      skippedTitleRows: countTitleRows(this.#counts, header),
      skippedSpacerColumns:
        range.endColumn - range.startColumn + 1 - columns.length,
      // Every column of the range: a column holding only formulas with no
      // cached value reads as a spacer, and is exactly what must be reported.
      // To the range's last row, as the table reader's region runs: a
      // trailing row of only such formulas is dropped, and must be reported.
      region: {
        startRow: header,
        endRow: range.endRow,
        startColumn: range.startColumn,
        endColumn: range.endColumn,
      },
    };
  }
}

/** The canonical date columns a mapping declares, by normalized key. */
function dateColumnKeys(mapping: ColumnMapping | undefined): Set<string> {
  const keys = new Set<string>();
  for (const column of mapping?.columns ?? []) {
    if (column.coercion?.type !== "date") continue;
    for (const spelling of [column.name, ...column.aliases]) {
      keys.add(normalizedColumnKey(spelling));
    }
  }
  return keys;
}

function coercedColumnKeys(mapping: ColumnMapping | undefined): Set<string> {
  const keys = new Set<string>();
  for (const column of mapping?.columns ?? []) {
    if (column.coercion === undefined) continue;
    for (const spelling of [column.name, ...column.aliases]) {
      keys.add(normalizedColumnKey(spelling));
    }
  }
  return keys;
}

function changedWhileRead(source: ConsolidationSource): ConsultChimpsError {
  return new ConsultChimpsError(
    XLSX_ERRORS.XLSX_READ_FAILED,
    `Workbook ${source.source} changed while it was being consolidated, so the consolidated workbook was not written. Run the consolidation again once the workbook is no longer being changed.`,
    { details: source.details },
  );
}

async function withWorkbook<T>(
  source: ConsolidationSource,
  expectedVersion: string | undefined,
  use: (workbook: StreamedWorkbook, version: string) => Promise<T>,
): Promise<T> {
  const opened = await source.open();
  try {
    const workbook = await StreamedWorkbook.open(opened, {
      file: source.file,
      source: source.source,
      details: source.details,
    });
    // The directory lists every entry's size and CRC, so any change to the
    // workbook's contents between two reads changes this.
    const version = `${String(opened.size)}:${String(workbook.fingerprint)}`;
    if (expectedVersion !== undefined && version !== expectedVersion) {
      throw changedWhileRead(source);
    }
    let result: T;
    try {
      result = await use(workbook, version);
    } finally {
      workbook.releaseStrings();
    }
    try {
      await opened.verifyUnchanged?.();
    } catch {
      throw changedWhileRead(source);
    }
    return result;
  } finally {
    await opened.close();
  }
}

function selectedSheets(
  sheets: readonly StreamedSheet[],
  settings: ConsolidationSettings,
): Array<{ sheet: StreamedSheet; index: number }> {
  const selected = settings.sheets
    ? new Set(settings.sheets.map((sheet) => sheet.toLocaleLowerCase()))
    : undefined;
  return sheets.flatMap((sheet, index) =>
    (!settings.includeHiddenSheets && !sheet.visible) ||
    (selected !== undefined && !selected.has(sheet.name.toLocaleLowerCase()))
      ? []
      : [{ sheet, index }],
  );
}

/** Pass a table's mapped rows to `emit`, batching through a coercing mapping. */
class TableRows implements WorksheetConsumer {
  readonly #output: TableOutput;
  readonly #source: ConsolidationSource;
  readonly #mapping: ColumnMapping | undefined;
  readonly #emit: (
    values: ReadonlyArray<WritableCellValue>,
    rowNumber: number,
  ) => void;
  readonly #position = new Map<number, number>();
  #begun = false;
  #rows = 0;
  #batch: TableRow[] = [];
  #batchRows: number[] = [];

  constructor(
    output: TableOutput,
    source: ConsolidationSource,
    mapping: ColumnMapping | undefined,
    emit: (values: ReadonlyArray<WritableCellValue>, rowNumber: number) => void,
  ) {
    this.#output = output;
    this.#source = source;
    this.#mapping = mapping;
    this.#emit = emit;
    output.table.columns.forEach((column, position) => {
      this.#position.set(column, position);
    });
  }

  get rows(): number {
    return this.#rows;
  }

  begin(): void {
    // Rows already passed on cannot be taken back, so a read that has to start
    // over means the worksheet is no longer the one the first pass read.
    if (this.#begun) throw changedWhileRead(this.#source);
    this.#begun = true;
  }

  row(row: number, cells: readonly StreamedCell[]): void {
    const table = this.#output.table;
    if (row <= table.headerRow) return;
    const values: Array<StreamedValue | null> = new Array<StreamedValue | null>(
      table.columns.length,
    ).fill(null);
    let populated = false;
    for (const cell of cells) {
      const position = this.#position.get(cell.column);
      if (position === undefined) continue;
      values[position] = cell.value;
      if (!isBlankValue(cell.value)) populated = true;
    }
    if (!populated) return;
    this.#rows += 1;
    if (!this.#output.coerced) {
      this.#emit([...values, ...this.#output.constants], row + 1);
      return;
    }
    const record = Object.create(null) as TableRow;
    table.names.forEach((name, position) => {
      record[name] = asCellValue(values[position] ?? null);
    });
    this.#batch.push(record);
    this.#batchRows.push(row + 1);
    if (this.#batch.length >= MAPPING_BATCH_ROWS) this.flush();
  }

  /** Apply the mapping to the rows collected so far and pass them on. */
  flush(): void {
    if (this.#batch.length === 0 || this.#mapping === undefined) return;
    const table = this.#output.table;
    const mapped = applyColumnMapping(
      {
        columns: [...table.names],
        rows: this.#batch,
        sourceRows: this.#batchRows,
        source: { file: this.#source.file, sheet: table.sheet },
      },
      this.#mapping,
    ).table;
    mapped.rows.forEach((record, index) => {
      this.#emit(
        mapped.columns.map((column) =>
          Object.hasOwn(record, column) ? (record[column] ?? null) : null,
        ),
        this.#batchRows[index]!,
      );
    });
    this.#batch = [];
    this.#batchRows = [];
  }
}

/**
 * The first pass, and every refusal a consolidation can make: unreadable
 * inputs, no tables, the mapping's refusals in the order the mapping applies
 * them, and a source column colliding with a data column.
 */
export async function planConsolidation(
  sources: readonly ConsolidationSource[],
  settings: ConsolidationSettings,
): Promise<ConsolidationPlan> {
  const { signal, outputContext } = settings;
  const mapping = settings.mapping;
  const dateKeys = dateColumnKeys(mapping);
  const tables: SheetTable[] = [];
  const inputVersions: string[] = [];
  const checkpoint = async (): Promise<void> => {
    if (settings.yieldControl) await yieldToEventLoop();
    throwIfAborted(signal, CONSOLIDATE_OPERATION, outputContext);
  };
  let skippedTitleRows = 0;
  let skippedSpacerColumns = 0;
  const uncachedFormulas: string[] = [];

  for (const [input, source] of sources.entries()) {
    throwIfAborted(signal, CONSOLIDATE_OPERATION, outputContext);
    const before = tables.length;
    await withWorkbook(source, undefined, async (workbook, version) => {
      inputVersions.push(version);
      for (const { sheet, index } of selectedSheets(
        workbook.sheets,
        settings,
      )) {
        const profile = new SheetProfile(settings.headerRow, dateKeys.size > 0);
        const read = await workbook.readWorksheet(sheet, profile);
        const outcome = profile.finish(read, settings.headerRow);
        skippedTitleRows += outcome.skippedTitleRows;
        skippedSpacerColumns += outcome.skippedSpacerColumns;
        // A worksheet that yielded no table is counted whole: its formulas
        // may be why it looked empty.
        uncachedFormulas.push(
          ...uncachedLocationsWithin(
            // Excel's own spelling, `[file.xlsx]Sheet!B4`, once two inputs
            // can share a sheet name.
            sources.length > 1 ? `[${source.file}]${sheet.name}` : sheet.name,
            read.uncachedFormulas,
            outcome.table === undefined ? undefined : outcome.region,
          ),
        );
        if (outcome.table !== undefined) {
          tables.push({
            ...outcome.table,
            input,
            sheetIndex: index,
            sheet: sheet.name,
            gathered: read.gathered,
          });
        }
      }
    });
    const read = tables.slice(before);
    settings.onProgress?.({
      operation: CONSOLIDATE_OPERATION,
      stage: "reading-workbooks",
      completed: input + 1,
      total: sources.length,
      detail: source.file,
      measures: {
        tables: read.length,
        rows: read.reduce((sum, table) => sum + table.rowCount, 0),
        columns: read.reduce(
          (widest, table) => Math.max(widest, table.columns.length),
          0,
        ),
      },
    });
    if (settings.yieldControl) await yieldToEventLoop();
  }

  throwIfAborted(signal, CONSOLIDATE_OPERATION, outputContext);
  if (tables.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_TABLES,
      `No visible, non-empty worksheets were found in the input workbooks.${uncachedFormulaHint(uncachedFormulas)}`,
      { details: { uncachedFormulas } },
    );
  }

  const headerTable = (table: SheetTable): Table => ({
    columns: [...table.names],
    rows: [],
    source: {
      file: sources[table.input]!.file,
      firstDataRow: table.headerRow + 2,
      sheet: table.sheet,
    },
  });

  if (mapping !== undefined && dateKeys.size > 0) {
    // The mapping's first refusal: a declared date coercion over a value that
    // is not text, in the first table holding one, at its first such row.
    for (const table of tables) {
      let first: { offence: DateOffence; position: number } | undefined;
      table.names.forEach((name, position) => {
        const offence = table.dateOffences[position];
        if (
          offence !== undefined &&
          dateKeys.has(normalizedColumnKey(name)) &&
          (first === undefined || offence.row < first.offence.row)
        ) {
          first = { offence, position };
        }
      });
      if (first !== undefined) {
        const record: TableRow = {};
        record[table.names[first.position]!] = asCellValue(first.offence.value);
        refuseNonTextDateColumns(
          [
            {
              ...headerTable(table),
              rows: [record],
              sourceRows: [first.offence.row + 1],
            },
          ],
          mapping,
        );
      }
    }
  }

  const coercedKeys = coercedColumnKeys(mapping);
  const unmappedColumns: string[] = [];
  const seenUnmapped = new Set<string>();
  const mappedTables: Array<{
    table: SheetTable;
    mappedColumns: string[];
    constants: CellValue[];
    coerced: boolean;
    widths: number[];
  }> = [];
  for (const table of tables) {
    if (mapping === undefined) {
      mappedTables.push({
        table,
        mappedColumns: [...table.names],
        constants: [],
        coerced: false,
        widths: [...table.widths],
      });
      continue;
    }
    // Header-only, the mapping makes every refusal it makes per table (two
    // columns folding into one, a constant colliding with a column) and
    // reports the columns it did not claim.
    const result = applyColumnMapping(headerTable(table), mapping);
    for (const column of result.unmappedColumns) {
      if (!seenUnmapped.has(column)) {
        seenUnmapped.add(column);
        unmappedColumns.push(column);
      }
    }
    const constants = Object.values(mapping.constants ?? {});
    const coerced = table.names.some((name) =>
      coercedKeys.has(normalizedColumnKey(name)),
    );
    const entry = {
      table,
      mappedColumns: result.table.columns,
      constants,
      coerced,
      widths: [
        ...table.widths,
        ...constants.map((value) => cellWidthLength(value)),
      ],
    };
    if (coerced) {
      // A coerced value's width is the coerced text's, and a value the
      // coercion cannot read is refused here, before anything is written.
      const widths = new Array<number>(entry.mappedColumns.length).fill(0);
      const output: TableOutput = {
        table,
        mappedColumns: entry.mappedColumns,
        constants,
        coerced,
        projection: [],
      };
      const source = sources[table.input]!;
      await withWorkbook(
        source,
        inputVersions[table.input],
        async (workbook) => {
          const rows = new TableRows(output, source, mapping, (values) => {
            values.forEach((value, position) => {
              widths[position] = Math.max(
                widths[position]!,
                cellWidthLength(value),
              );
            });
          });
          await workbook.readWorksheet(
            workbook.sheets[table.sheetIndex]!,
            rows,
            {
              gather: table.gathered,
              between: checkpoint,
            },
          );
          rows.flush();
          if (rows.rows !== table.rowCount) throw changedWhileRead(source);
        },
      );
      entry.widths = widths;
    }
    mappedTables.push(entry);
  }

  const union = planTableUnion(
    mappedTables.map((entry) => entry.mappedColumns),
    {
      addSourceColumns: settings.addSourceColumns,
      normalizeHeaders: settings.normalizeHeaders,
    },
  );
  const sourceColumnCount = union.sourceColumns === undefined ? 0 : 3;
  const widthLengths = union.columns.map((column) => column.length);
  const outputs: TableOutput[] = mappedTables.map((entry, index) => {
    const positions = new Map(
      entry.mappedColumns.map(
        (column, position) => [column, position] as const,
      ),
    );
    const projection = union.inputColumns[index]!.map((column) =>
      column ? (positions.get(column) ?? -1) : -1,
    );
    projection.forEach((position, outputIndex) => {
      if (position >= 0) {
        widthLengths[outputIndex] = Math.max(
          widthLengths[outputIndex]!,
          entry.widths[position]!,
        );
      }
    });
    if (sourceColumnCount > 0) {
      const base = union.columns.length - sourceColumnCount;
      const source = sources[entry.table.input]!;
      widthLengths[base] = Math.max(widthLengths[base]!, source.file.length);
      widthLengths[base + 1] = Math.max(
        widthLengths[base + 1]!,
        entry.table.sheet.length,
      );
      widthLengths[base + 2] = Math.max(
        widthLengths[base + 2]!,
        String(entry.table.lastRowNumber).length,
      );
    }
    return {
      table: entry.table,
      mappedColumns: entry.mappedColumns,
      constants: entry.constants,
      coerced: entry.coerced,
      projection,
    };
  });

  return {
    outputs,
    columns: union.columns,
    widths: widthLengths.map((length) => tableColumnWidth(length)),
    rowCount: tables.reduce((sum, table) => sum + table.rowCount, 0),
    addSourceColumns: sourceColumnCount > 0,
    unmappedColumns,
    suggestion:
      settings.suggestMapping === true
        ? suggestMappingForTables(tables.map(headerTable))
        : undefined,
    inputTables: tables.length,
    skippedTitleRows,
    skippedSpacerColumns,
    uncachedFormulas,
    inputVersions,
  };
}

/**
 * The second pass: read every table's rows again and write them, in input and
 * worksheet order, through the streaming writer into `sink`. A sink that
 * throws fails the run with its own error, not as a failure to read the input
 * being copied when it happened, and not as whatever the zip stream it broke
 * reports next.
 */
export async function writeConsolidation(
  sources: readonly ConsolidationSource[],
  plan: ConsolidationPlan,
  settings: ConsolidationSettings,
  sheetName: string,
  sink: ConsolidationSink,
): Promise<void> {
  let failure: { readonly error: unknown } | undefined;
  const guarded: ConsolidationSink = {
    write: (chunk) => {
      try {
        sink.write(chunk);
      } catch (error) {
        failure ??= { error };
        throw error;
      }
    },
    flush: async () => {
      try {
        await sink.flush();
      } catch (error) {
        failure ??= { error };
        throw error;
      }
    },
  };
  try {
    await copyTables(sources, plan, settings, sheetName, guarded);
  } catch (error) {
    throw failure === undefined ? error : failure.error;
  }
}

async function copyTables(
  sources: readonly ConsolidationSource[],
  plan: ConsolidationPlan,
  settings: ConsolidationSettings,
  sheetName: string,
  sink: ConsolidationSink,
): Promise<void> {
  const { signal, outputContext } = settings;
  const writer = new TableWorkbookWriter({
    sheetName,
    columns: plan.columns,
    widths: plan.widths,
    rowCount: plan.rowCount,
    onChunk: (chunk) => {
      sink.write(chunk);
    },
  });
  const between = async (): Promise<void> => {
    await sink.flush();
    if (settings.yieldControl) await yieldToEventLoop();
    throwIfAborted(signal, CONSOLIDATE_OPERATION, outputContext);
  };
  const width = plan.columns.length;
  const sourceBase = plan.addSourceColumns ? width - 3 : width;

  for (const [input, source] of sources.entries()) {
    const outputs = plan.outputs.filter(
      (output) => output.table.input === input,
    );
    if (outputs.length === 0) continue;
    throwIfAborted(signal, CONSOLIDATE_OPERATION, outputContext);
    await withWorkbook(source, plan.inputVersions[input], async (workbook) => {
      for (const output of outputs) {
        const table = output.table;
        const rows = new TableRows(
          output,
          source,
          settings.mapping,
          (values, rowNumber) => {
            const row = new Array<WritableCellValue>(width).fill(null);
            output.projection.forEach((position, outputIndex) => {
              if (position >= 0) row[outputIndex] = values[position] ?? null;
            });
            if (plan.addSourceColumns) {
              row[sourceBase] = source.file;
              row[sourceBase + 1] = table.sheet;
              row[sourceBase + 2] = rowNumber;
            }
            writer.writeRow(row.map((value) => writableCellValue(value)));
          },
        );
        await workbook.readWorksheet(workbook.sheets[table.sheetIndex]!, rows, {
          gather: table.gathered,
          between,
        });
        rows.flush();
        if (rows.rows !== table.rowCount) throw changedWhileRead(source);
        await between();
      }
    });
  }
  writer.finish();
  await sink.flush();
}
