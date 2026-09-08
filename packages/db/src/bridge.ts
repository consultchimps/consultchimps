import { ConsultChimpsError } from "@consultchimps/core";
import type { CellValue, Table } from "@consultchimps/tabular";

import type { Database, InsertedRecord } from "./database.js";
import { identifierKey, sameIdentifier, RECORD_ID_COLUMN } from "./schema.js";

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
  // The whole batch is one transaction: if a later row fails a constraint,
  // earlier inserts and their consumed id counters roll back, so retrying the
  // same call does not duplicate a partial prefix.
  return database.sql.transaction(() =>
    insertRecordsFromTable(database, tableName, table),
  );
}

/**
 * The insert itself, without a transaction of its own, so a caller that is
 * already inside one (the import, which creates a table and fills it as a
 * single unit) can reuse it. SQLite has no nested transactions, so this split
 * is what keeps `addRecordsFromTable` all-or-nothing on its own while the
 * import stays all-or-nothing across every table it creates.
 */
export function insertRecordsFromTable(
  database: Database,
  tableName: string,
  table: Table,
): InsertedRecord[] {
  const schema = database.getTableSchema(tableName);
  // Column names are matched case-insensitively, like SQLite identifiers, so a
  // Table header of any casing lines up with its declared column and a
  // Record ID column in any casing is ignored.
  const known = new Set(
    schema.columns.map((column) => identifierKey(column.name)),
  );
  const mappedColumns = table.columns.filter(
    (column) => !sameIdentifier(column, RECORD_ID_COLUMN),
  );

  for (const column of mappedColumns) {
    if (!known.has(identifierKey(column))) {
      throw new ConsultChimpsError(
        "DB_UNKNOWN_COLUMN",
        `The table "${tableName}" has no column "${column}" to receive the bridged data.`,
        { details: { table: tableName, column } },
      );
    }
  }

  const inserted: InsertedRecord[] = [];
  for (const row of table.rows) {
    // A prototype-free destination, and own-property reads, so a column named
    // like an Object.prototype member ("constructor") reads the row's own
    // value or null, never an inherited function.
    const values: Record<string, CellValue> = Object.create(null);
    for (const column of mappedColumns) {
      values[column] = Object.prototype.hasOwnProperty.call(row, column)
        ? (row[column] ?? null)
        : null;
    }
    inserted.push(database.insertRecord(tableName, values));
  }
  return inserted;
}
