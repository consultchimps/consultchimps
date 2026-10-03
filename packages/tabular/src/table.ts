import { ConsultChimpsError } from "@consultchimps/core";

export type CellValue = string | number | boolean | null;
export type TableRow = Record<string, CellValue>;

export interface TableSource {
  file?: string;
  sheet?: string;
  firstDataRow?: number;
}

export interface Table {
  columns: string[];
  rows: TableRow[];
  source?: TableSource;
  sourceRows?: number[];
}

export interface UnionTablesOptions {
  addSourceColumns?: boolean | undefined;
  /**
   * Treat headers that differ only in case, spacing, or punctuation - such as
   * "Failed Checks", "Failed_Checks", and "Failed  Checks " - as the same
   * column. The first spelling seen names the output column.
   */
  normalizeHeaders?: boolean | undefined;
  sourceColumnNames?:
    | {
        file: string;
        sheet: string;
        row: string;
      }
    | undefined;
}

export interface GroupTableByColumnOptions {
  includeBlank?: boolean | undefined;
}

export interface TableGroup {
  table: Table;
  value: CellValue;
}

export interface GroupTableByColumnResult {
  column: string;
  groups: TableGroup[];
  skippedRows: number;
}

const DEFAULT_SOURCE_COLUMNS = {
  file: "_source_file",
  sheet: "_source_sheet",
  row: "_source_row",
} as const;

/**
 * Case-folded matching key. The fold is `toLowerCase`, never
 * `toLocaleLowerCase`: the locale-aware fold reads the host's default locale,
 * so the same headers would match on one machine and not on another - under a
 * Turkish locale "ID" folds to "ıd" rather than "id" - and identical inputs
 * must produce identical columns everywhere.
 */
export function columnKey(column: string): string {
  return column.trim().toLowerCase();
}

/**
 * Matching key that also ignores spacing and punctuation differences, so
 * "Failed Checks", "Failed_Checks", and "Reviewer: Lead Contact" versus
 * "Reviewer_Lead_Contact" resolve to the same column. Letters and digits in
 * any script are kept; every other run of characters becomes one underscore.
 * The case fold is locale-independent for the reason `columnKey` explains.
 */
export function normalizedColumnKey(column: string): string {
  // The first replace collapses every separator run - underscores included -
  // to one "_", so the edge trims below never face repeated underscores and
  // stay linear on any input.
  const normalized = column
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, "_")
    .replace(/^_/, "")
    .replace(/_$/, "");
  return normalized === "" ? columnKey(column) : normalized;
}

/**
 * Name every column of a header row so that no two names collide, filling a
 * blank header with its position.
 *
 * Uniqueness is decided against the whole header row, not against the names
 * seen so far. Counting occurrences left to right is not enough: for A, A, A_2
 * it renames the second A to A_2 and then meets the real A_2, and two columns
 * end up with one name. Where those names become object keys, as they do in
 * every `TableRow`, the second column's values are simply overwritten by the
 * third's, and nothing says so.
 *
 * So every original spelling is reserved first, and a generated name has to
 * clear both the originals and the names already generated. The suffix is the
 * lowest number from 2 that does, which keeps the result deterministic and
 * keeps the common case (A, A) reading as A, A_2.
 *
 * Names are compared case-insensitively, through `columnKey`, because that is
 * how every other column lookup here matches them: "Region" and "region" are
 * one column, so they cannot both be a column name.
 */
export function uniqueHeaders(values: Array<string | null>): string[] {
  const bases = values.map(
    (value, index) => value?.trim() || `column_${index + 1}`,
  );
  // Every original spelling, so a generated name never lands on one that the
  // header row already carries further along.
  const reserved = new Set(bases.map(columnKey));
  const taken = new Set<string>();
  // The next suffix to try for each folded base, so a run of repeated headers
  // does not restart at 2 and re-probe every number already handed out. A wide
  // file of identical headers then costs one probe per column rather than one
  // per column per earlier duplicate, which matters for delimited text, whose
  // width nothing bounds. The names produced are the same either way.
  const nextSuffix = new Map<string, number>();

  return bases.map((base) => {
    const key = columnKey(base);
    if (!taken.has(key)) {
      taken.add(key);
      return base;
    }
    let suffix = nextSuffix.get(key) ?? 2;
    let candidate = `${base}_${suffix}`;
    while (
      taken.has(columnKey(candidate)) ||
      reserved.has(columnKey(candidate))
    ) {
      suffix += 1;
      candidate = `${base}_${suffix}`;
    }
    nextSuffix.set(key, suffix + 1);
    taken.add(columnKey(candidate));
    return candidate;
  });
}

function isBlankValue(value: CellValue): boolean {
  return value === null || (typeof value === "string" && value.trim() === "");
}

function groupKey(value: CellValue): string {
  if (value === null) {
    return "null";
  }

  return `${typeof value}:${String(value)}`;
}

export function groupTableByColumn(
  table: Table,
  column: string,
  options: GroupTableByColumnOptions = {},
): GroupTableByColumnResult {
  const requestedColumnKey = columnKey(column);
  const matchedColumn = table.columns.find(
    (candidate) => columnKey(candidate) === requestedColumnKey,
  );

  if (!matchedColumn) {
    throw new ConsultChimpsError(
      "TABLE_COLUMN_NOT_FOUND",
      `Column "${column}" was not found in the table.`,
      {
        details: {
          availableColumns: table.columns,
          column,
        },
      },
    );
  }

  const includeBlank = options.includeBlank ?? true;
  const groups = new Map<
    string,
    {
      rowIndexes: number[];
      value: CellValue;
    }
  >();
  let skippedRows = 0;

  table.rows.forEach((row, rowIndex) => {
    const rawValue = row[matchedColumn] ?? null;
    const value = isBlankValue(rawValue) ? null : rawValue;

    if (value === null && !includeBlank) {
      skippedRows += 1;
      return;
    }

    const key = groupKey(value);
    const existing = groups.get(key);
    if (existing) {
      existing.rowIndexes.push(rowIndex);
      return;
    }

    groups.set(key, {
      rowIndexes: [rowIndex],
      value,
    });
  });

  return {
    column: matchedColumn,
    groups: [...groups.values()].map(({ rowIndexes, value }) => {
      const groupedTable: Table = {
        columns: [...table.columns],
        rows: rowIndexes.map((rowIndex) => ({
          ...table.rows[rowIndex],
        })),
      };

      if (table.source) {
        groupedTable.source = { ...table.source };
      }
      if (table.sourceRows) {
        groupedTable.sourceRows = rowIndexes.map(
          (rowIndex) => table.sourceRows?.[rowIndex] ?? rowIndex + 2,
        );
      }

      return {
        table: groupedTable,
        value,
      };
    }),
    skippedRows,
  };
}

/** The source columns a union appends, by role. */
export interface UnionSourceColumns {
  file: string;
  sheet: string;
  row: string;
}

/**
 * How tables stack into one union, decided from their headers alone: the
 * output columns, and for each input, which of its columns fills each output
 * column. A caller that streams rows applies it row by row; `unionTables` is
 * this plan applied to rows held in memory.
 */
export interface TableUnionPlan {
  columns: string[];
  /**
   * Per input, per output column, the input column that fills it, or
   * undefined when the input has no such column or the column is a source
   * column.
   */
  inputColumns: Array<Array<string | undefined>>;
  /** The appended source columns, or undefined when none are added. */
  sourceColumns: UnionSourceColumns | undefined;
}

/**
 * Plan a union of tables with these column lists. The first spelling seen
 * names each output column; the source columns, when added, come last and
 * may not collide with an input column.
 */
export function planTableUnion(
  columnLists: ReadonlyArray<readonly string[]>,
  options: UnionTablesOptions = {},
): TableUnionPlan {
  if (columnLists.length === 0) {
    throw new ConsultChimpsError(
      "TABLES_EMPTY",
      "At least one table is required for a union.",
    );
  }

  const addSourceColumns = options.addSourceColumns ?? true;
  const sourceColumns = options.sourceColumnNames ?? DEFAULT_SOURCE_COLUMNS;
  const keyOf =
    options.normalizeHeaders === true ? normalizedColumnKey : columnKey;
  const outputColumnByKey = new Map<string, string>();

  for (const columns of columnLists) {
    for (const column of columns) {
      const key = keyOf(column);
      if (!outputColumnByKey.has(key)) {
        outputColumnByKey.set(key, column);
      }
    }
  }

  const sourceKeys = new Set<string>();
  if (addSourceColumns) {
    for (const column of Object.values(sourceColumns)) {
      const key = keyOf(column);
      if (outputColumnByKey.has(key)) {
        throw new ConsultChimpsError(
          "TABLE_SOURCE_COLUMN_COLLISION",
          `Source column "${column}" already exists in the input data.`,
          { details: { column } },
        );
      }
      outputColumnByKey.set(key, column);
      sourceKeys.add(key);
    }
  }

  const inputColumns = columnLists.map((columns) => {
    const inputColumnByKey = new Map<string, string>();
    for (const column of columns) {
      const key = keyOf(column);
      if (!inputColumnByKey.has(key)) {
        inputColumnByKey.set(key, column);
      }
    }
    return [...outputColumnByKey.keys()].map((key) =>
      sourceKeys.has(key) ? undefined : inputColumnByKey.get(key),
    );
  });

  return {
    columns: [...outputColumnByKey.values()],
    inputColumns,
    sourceColumns: addSourceColumns
      ? {
          file: sourceColumns.file,
          sheet: sourceColumns.sheet,
          row: sourceColumns.row,
        }
      : undefined,
  };
}

export function unionTables(
  tables: Table[],
  options: UnionTablesOptions = {},
): Table {
  const plan = planTableUnion(
    tables.map((table) => table.columns),
    options,
  );
  const rows: TableRow[] = [];

  tables.forEach((table, tableIndex) => {
    const inputColumns = plan.inputColumns[tableIndex]!;
    table.rows.forEach((inputRow, index) => {
      const outputRow: TableRow = {};

      plan.columns.forEach((outputColumn, position) => {
        if (
          plan.sourceColumns !== undefined &&
          position >= plan.columns.length - 3
        ) {
          return;
        }
        const inputColumn = inputColumns[position];
        outputRow[outputColumn] = inputColumn
          ? (inputRow[inputColumn] ?? null)
          : null;
      });

      if (plan.sourceColumns !== undefined) {
        outputRow[plan.sourceColumns.file] = table.source?.file ?? null;
        outputRow[plan.sourceColumns.sheet] = table.source?.sheet ?? null;
        outputRow[plan.sourceColumns.row] =
          table.sourceRows?.[index] ??
          (table.source?.firstDataRow ?? 2) + index;
      }

      rows.push(outputRow);
    });
  });

  return { columns: plan.columns, rows };
}
