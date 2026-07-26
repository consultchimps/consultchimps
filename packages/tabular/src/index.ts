import { ChimpconsError } from "@chimpcons/core";

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
  sourceColumnNames?:
    | {
        file: string;
        sheet: string;
        row: string;
      }
    | undefined;
}

const DEFAULT_SOURCE_COLUMNS = {
  file: "_source_file",
  sheet: "_source_sheet",
  row: "_source_row",
} as const;

export function columnKey(column: string): string {
  return column.trim().toLocaleLowerCase();
}

export function uniqueHeaders(values: Array<string | null>): string[] {
  const occurrences = new Map<string, number>();

  return values.map((value, index) => {
    const base = value?.trim() || `column_${index + 1}`;
    const key = columnKey(base);
    const occurrence = (occurrences.get(key) ?? 0) + 1;
    occurrences.set(key, occurrence);
    return occurrence === 1 ? base : `${base}_${occurrence}`;
  });
}

export function unionTables(
  tables: Table[],
  options: UnionTablesOptions = {},
): Table {
  if (tables.length === 0) {
    throw new ChimpconsError(
      "TABLES_EMPTY",
      "At least one table is required for a union.",
    );
  }

  const addSourceColumns = options.addSourceColumns ?? true;
  const sourceColumns = options.sourceColumnNames ?? DEFAULT_SOURCE_COLUMNS;
  const outputColumnByKey = new Map<string, string>();

  for (const table of tables) {
    for (const column of table.columns) {
      const key = columnKey(column);
      if (!outputColumnByKey.has(key)) {
        outputColumnByKey.set(key, column);
      }
    }
  }

  if (addSourceColumns) {
    for (const column of Object.values(sourceColumns)) {
      const key = columnKey(column);
      if (outputColumnByKey.has(key)) {
        throw new ChimpconsError(
          "TABLE_SOURCE_COLUMN_COLLISION",
          `Source column "${column}" already exists in the input data.`,
          { details: { column } },
        );
      }
      outputColumnByKey.set(key, column);
    }
  }

  const columns = [...outputColumnByKey.values()];
  const rows: TableRow[] = [];

  for (const table of tables) {
    const inputColumnByKey = new Map(
      table.columns.map((column) => [columnKey(column), column] as const),
    );

    table.rows.forEach((inputRow, index) => {
      const outputRow: TableRow = {};

      for (const [key, outputColumn] of outputColumnByKey) {
        if (
          addSourceColumns &&
          [
            columnKey(sourceColumns.file),
            columnKey(sourceColumns.sheet),
            columnKey(sourceColumns.row),
          ].includes(key)
        ) {
          continue;
        }

        const inputColumn = inputColumnByKey.get(key);
        outputRow[outputColumn] = inputColumn
          ? (inputRow[inputColumn] ?? null)
          : null;
      }

      if (addSourceColumns) {
        outputRow[sourceColumns.file] = table.source?.file ?? null;
        outputRow[sourceColumns.sheet] = table.source?.sheet ?? null;
        outputRow[sourceColumns.row] =
          table.sourceRows?.[index] ??
          (table.source?.firstDataRow ?? 2) + index;
      }

      rows.push(outputRow);
    });
  }

  return { columns, rows };
}
