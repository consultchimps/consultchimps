import { ConsultChimpsError } from "@consultchimps/core";
import type { CellValue, Table } from "@consultchimps/tabular";

import type { Database, InsertedRecord } from "./database.js";
import { RECORD_ID_COLUMN } from "./schema.js";

/**
 * The bridge between the `@consultchimps/tabular` `Table` model and a database
 * table, in both directions. It lets a database feed the existing table
 * operations (PowerPoint populate, split) and be fed rows in return, with the
 * `Table` as the exchange format either way.
 */

/**
 * Read a database table out as a `Table`. Columns come back Record ID first,
 * then the declared columns; rows are in insertion order with values coerced to
 * their declared types.
 */
export function databaseTableToTable(
  database: Database,
  tableName: string,
): Table {
  return {
    columns: database.columnNames(tableName),
    rows: database.readRecords(tableName),
  };
}

/**
 * Insert the rows of a `Table` into an existing database table, generating a
 * fresh Record ID for each. A Record ID column in the input is ignored, since
 * identifiers are always generated; any other column the input carries must be
 * declared on the target table. Returns the inserted records in input order.
 */
export function addRecordsFromTable(
  database: Database,
  tableName: string,
  table: Table,
): InsertedRecord[] {
  const schema = database.getTableSchema(tableName);
  const known = new Set(schema.columns.map((column) => column.name));
  const mappedColumns = table.columns.filter(
    (column) => column !== RECORD_ID_COLUMN,
  );

  for (const column of mappedColumns) {
    if (!known.has(column)) {
      throw new ConsultChimpsError(
        "DB_UNKNOWN_COLUMN",
        `The table "${tableName}" has no column "${column}" to receive the bridged data.`,
        { details: { table: tableName, column } },
      );
    }
  }

  const inserted: InsertedRecord[] = [];
  for (const row of table.rows) {
    const values: Record<string, CellValue> = {};
    for (const column of mappedColumns) {
      values[column] = row[column] ?? null;
    }
    inserted.push(database.insertRecord(tableName, values));
  }
  return inserted;
}
