import { ConsultChimpsError } from "@consultchimps/core";
import type { CellValue, TableRow } from "@consultchimps/tabular";

import {
  loadSqlDatabase,
  type SqlDatabase,
  type SqlEngineConfig,
  type SqlValueType,
} from "./engine.js";
import {
  assertSafeIdentifier,
  cellFromSqlValue,
  formatRecordId,
  quoteIdentifier,
  sqlStorageClass,
  sqlValueFromCell,
  type ColumnDefinition,
  type RecordIdConfig,
  type TableSchema,
  METADATA_TABLE,
  RECORD_ID_COLUMN,
  SCHEMA_FORMAT_VERSION,
  TABLE_REGISTRY_TABLE,
} from "./schema.js";

/** The outcome of inserting a record: its generated Record ID and internal rowid. */
export interface InsertedRecord {
  recordId: string;
  rowId: number;
}

/** The stored shape of a table definition in the registry. */
interface StoredDefinition {
  columns: ColumnDefinition[];
  foreignKeys: TableSchema["foreignKeys"];
  recordId: RecordIdConfig;
}

/**
 * A local relational database held in memory. It owns a schema, generates a
 * stable Record ID for every inserted record, and persists both the schema and
 * the per-table id counter inside the SQLite file, so serializing and reloading
 * the bytes restores the schema and the next-id state. It never touches sql.js
 * directly; all engine access goes through the `SqlDatabase` wrapper.
 */
export class Database {
  readonly #sql: SqlDatabase;

  private constructor(sql: SqlDatabase) {
    this.#sql = sql;
  }

  /** Create an empty database with the metadata tables initialized. */
  static async create(config?: SqlEngineConfig): Promise<Database> {
    const sql = await loadSqlDatabase(undefined, config);
    const database = new Database(sql);
    database.#initializeMetadata();
    return database;
  }

  /** Open a database from serialized bytes, reading its stored schema. */
  static async open(
    bytes: Uint8Array,
    config?: SqlEngineConfig,
  ): Promise<Database> {
    const sql = await loadSqlDatabase(bytes, config);
    const database = new Database(sql);
    sql.run("PRAGMA foreign_keys = ON;");
    database.#assertMetadataPresent();
    database.#assertSupportedSchemaVersion();
    return database;
  }

  #initializeMetadata(): void {
    this.#sql.run("PRAGMA foreign_keys = ON;");
    this.#sql.run(
      `CREATE TABLE ${quoteIdentifier(METADATA_TABLE)} (key TEXT PRIMARY KEY, value TEXT NOT NULL);`,
    );
    this.#sql.run(
      `CREATE TABLE ${quoteIdentifier(TABLE_REGISTRY_TABLE)} (name TEXT PRIMARY KEY, definition TEXT NOT NULL, next_counter INTEGER NOT NULL);`,
    );
    this.#sql.run(
      `INSERT INTO ${quoteIdentifier(METADATA_TABLE)} (key, value) VALUES (?, ?);`,
      ["schema_format_version", String(SCHEMA_FORMAT_VERSION)],
    );
  }

  #assertMetadataPresent(): void {
    const found = this.#sql.selectValue(
      "SELECT count(*) FROM sqlite_schema WHERE type = 'table' AND name IN (?, ?);",
      [METADATA_TABLE, TABLE_REGISTRY_TABLE],
    );
    if (found !== 2) {
      throw new ConsultChimpsError(
        "DB_NOT_A_WORKSPACE",
        "This database file was not created by ConsultChimps: its schema metadata tables are missing.",
        { details: { expectedTables: [METADATA_TABLE, TABLE_REGISTRY_TABLE] } },
      );
    }
  }

  #assertSupportedSchemaVersion(): void {
    const version = this.schemaFormatVersion();
    if (version > SCHEMA_FORMAT_VERSION) {
      throw new ConsultChimpsError(
        "DB_UNSUPPORTED_SCHEMA_VERSION",
        `This database was written with schema format version ${version}, which is newer than this version supports (${SCHEMA_FORMAT_VERSION}). Update ConsultChimps to open it.`,
        {
          details: {
            fileVersion: version,
            supportedVersion: SCHEMA_FORMAT_VERSION,
          },
        },
      );
    }
  }

  /** The metadata-format version stored in this database file. */
  schemaFormatVersion(): number {
    const value = this.#sql.selectValue(
      `SELECT value FROM ${quoteIdentifier(METADATA_TABLE)} WHERE key = ?;`,
      ["schema_format_version"],
    );
    return typeof value === "string" ? Number(value) : SCHEMA_FORMAT_VERSION;
  }

  /**
   * Create a table with a reserved Record ID column and the given user columns,
   * foreign keys, and Record ID configuration. The Record ID column is a UNIQUE
   * text column, immutable once a row is written (enforced by a trigger), and is
   * the only column a foreign key may reference.
   */
  createTable(schema: TableSchema): void {
    assertSafeIdentifier(schema.name, "table");
    this.#assertTableAbsent(schema.name);
    this.#validateRecordIdConfig(schema.name, schema.recordId);

    const seenColumns = new Set<string>();
    for (const column of schema.columns) {
      assertSafeIdentifier(column.name, "column");
      if (column.name === RECORD_ID_COLUMN) {
        throw new ConsultChimpsError(
          "DB_RESERVED_COLUMN",
          `The column name "${RECORD_ID_COLUMN}" is reserved for the generated Record ID.`,
          { details: { table: schema.name } },
        );
      }
      const key = column.name.toLowerCase();
      if (seenColumns.has(key)) {
        throw new ConsultChimpsError(
          "DB_DUPLICATE_COLUMN",
          `The table "${schema.name}" declares the column "${column.name}" more than once.`,
          { details: { table: schema.name, column: column.name } },
        );
      }
      seenColumns.add(key);
    }

    const definitions = schema.columns.map((column) => {
      const nullClause = column.nullable === false ? " NOT NULL" : "";
      return `${quoteIdentifier(column.name)} ${sqlStorageClass(column.type)}${nullClause}`;
    });

    const columnByName = new Map(
      schema.columns.map((column) => [column.name, column]),
    );
    const foreignKeyClauses = schema.foreignKeys.map((foreignKey) => {
      const column = columnByName.get(foreignKey.column);
      if (column === undefined) {
        throw new ConsultChimpsError(
          "DB_FOREIGN_KEY_COLUMN_MISSING",
          `The foreign key on "${schema.name}" references its own column "${foreignKey.column}", which is not declared.`,
          { details: { table: schema.name, column: foreignKey.column } },
        );
      }
      if (column.type !== "text") {
        throw new ConsultChimpsError(
          "DB_FOREIGN_KEY_COLUMN_TYPE",
          `The foreign key column "${foreignKey.column}" on "${schema.name}" must be a text column, because it holds a Record ID.`,
          { details: { table: schema.name, column: foreignKey.column } },
        );
      }
      if (
        foreignKey.referencesTable !== schema.name &&
        !this.#tableExists(foreignKey.referencesTable)
      ) {
        throw new ConsultChimpsError(
          "DB_FOREIGN_KEY_TABLE_MISSING",
          `The foreign key on "${schema.name}" references table "${foreignKey.referencesTable}", which does not exist.`,
          {
            details: {
              table: schema.name,
              referencesTable: foreignKey.referencesTable,
            },
          },
        );
      }
      assertSafeIdentifier(foreignKey.referencesTable, "table");
      return `FOREIGN KEY (${quoteIdentifier(foreignKey.column)}) REFERENCES ${quoteIdentifier(foreignKey.referencesTable)}(${quoteIdentifier(RECORD_ID_COLUMN)})`;
    });

    const parts = [
      `${quoteIdentifier(RECORD_ID_COLUMN)} TEXT NOT NULL UNIQUE`,
      ...definitions,
      ...foreignKeyClauses,
    ];
    this.#sql.run(
      `CREATE TABLE ${quoteIdentifier(schema.name)} (${parts.join(", ")});`,
    );

    // Enforce Record ID immutability at the storage layer: any statement that
    // would change a record_id aborts.
    this.#sql.run(
      `CREATE TRIGGER ${quoteIdentifier(`${schema.name}:record_id_immutable`)} ` +
        `BEFORE UPDATE OF ${quoteIdentifier(RECORD_ID_COLUMN)} ON ${quoteIdentifier(schema.name)} ` +
        `FOR EACH ROW WHEN OLD.${quoteIdentifier(RECORD_ID_COLUMN)} <> NEW.${quoteIdentifier(RECORD_ID_COLUMN)} ` +
        `BEGIN SELECT RAISE(ABORT, 'Record ID is immutable and cannot be changed'); END;`,
    );

    const definition: StoredDefinition = {
      columns: schema.columns,
      foreignKeys: schema.foreignKeys,
      recordId: schema.recordId,
    };
    this.#sql.run(
      `INSERT INTO ${quoteIdentifier(TABLE_REGISTRY_TABLE)} (name, definition, next_counter) VALUES (?, ?, ?);`,
      [schema.name, JSON.stringify(definition), 1],
    );
  }

  #validateRecordIdConfig(table: string, config: RecordIdConfig): void {
    if (config.prefix.trim() === "") {
      throw new ConsultChimpsError(
        "DB_INVALID_RECORD_ID_CONFIG",
        `The Record ID prefix for "${table}" cannot be empty.`,
        { details: { table } },
      );
    }
    if (!Number.isInteger(config.padding) || config.padding < 0) {
      throw new ConsultChimpsError(
        "DB_INVALID_RECORD_ID_CONFIG",
        `The Record ID padding for "${table}" must be a non-negative whole number.`,
        { details: { table, padding: config.padding } },
      );
    }
  }

  /**
   * Insert a record, generating and assigning the next Record ID in the table's
   * sequence. The provided values are keyed by column name; the Record ID is
   * generated, so it cannot be supplied.
   */
  insertRecord(
    tableName: string,
    values: Readonly<Record<string, CellValue>>,
  ): InsertedRecord {
    const definition = this.#requireDefinition(tableName);
    const columnByName = new Map(
      definition.columns.map((column) => [column.name, column]),
    );

    if (Object.prototype.hasOwnProperty.call(values, RECORD_ID_COLUMN)) {
      throw new ConsultChimpsError(
        "DB_RECORD_ID_IS_GENERATED",
        `The Record ID for "${tableName}" is generated and cannot be supplied when inserting.`,
        { details: { table: tableName } },
      );
    }

    const insertColumns: string[] = [RECORD_ID_COLUMN];
    const insertValues: SqlValueType[] = [];

    const counter = this.#nextCounter(tableName);
    const recordId = formatRecordId(definition.recordId, counter);
    insertValues.push(recordId);

    for (const [name, value] of Object.entries(values)) {
      const column = columnByName.get(name);
      if (column === undefined) {
        throw new ConsultChimpsError(
          "DB_UNKNOWN_COLUMN",
          `The table "${tableName}" has no column "${name}".`,
          { details: { table: tableName, column: name } },
        );
      }
      insertColumns.push(name);
      insertValues.push(sqlValueFromCell(column.type, value));
    }

    const placeholders = insertColumns.map(() => "?").join(", ");
    const quotedColumns = insertColumns.map(quoteIdentifier).join(", ");
    this.#sql.run(
      `INSERT INTO ${quoteIdentifier(tableName)} (${quotedColumns}) VALUES (${placeholders});`,
      insertValues,
    );

    this.#sql.run(
      `UPDATE ${quoteIdentifier(TABLE_REGISTRY_TABLE)} SET next_counter = ? WHERE name = ?;`,
      [counter + 1, tableName],
    );

    const rowId = this.#sql.selectValue("SELECT last_insert_rowid();");
    return {
      recordId,
      rowId: typeof rowId === "number" ? rowId : Number(rowId),
    };
  }

  #nextCounter(tableName: string): number {
    const value = this.#sql.selectValue(
      `SELECT next_counter FROM ${quoteIdentifier(TABLE_REGISTRY_TABLE)} WHERE name = ?;`,
      [tableName],
    );
    return typeof value === "number" ? value : Number(value);
  }

  /** The schema of a single table, read from the stored metadata. */
  getTableSchema(tableName: string): TableSchema {
    const definition = this.#requireDefinition(tableName);
    return {
      name: tableName,
      columns: definition.columns,
      foreignKeys: definition.foreignKeys,
      recordId: definition.recordId,
    };
  }

  /** The schema of every table, read from the stored metadata, ordered by name. */
  getSchema(): TableSchema[] {
    const rows = this.#sql.select(
      `SELECT name, definition FROM ${quoteIdentifier(TABLE_REGISTRY_TABLE)} ORDER BY name;`,
    );
    return rows.map((row) => {
      const definition = JSON.parse(
        String(row["definition"]),
      ) as StoredDefinition;
      return {
        name: String(row["name"]),
        columns: definition.columns,
        foreignKeys: definition.foreignKeys,
        recordId: definition.recordId,
      };
    });
  }

  /**
   * Read every record in a table as tabular rows, including the Record ID
   * column, with values coerced to their declared types. Rows come back ordered
   * by internal rowid, which is insertion order.
   */
  readRecords(tableName: string): TableRow[] {
    const definition = this.#requireDefinition(tableName);
    const columnNames = [
      RECORD_ID_COLUMN,
      ...definition.columns.map((column) => column.name),
    ];
    const selectList = columnNames.map(quoteIdentifier).join(", ");
    const rows = this.#sql.select(
      `SELECT ${selectList} FROM ${quoteIdentifier(tableName)} ORDER BY rowid;`,
    );
    return rows.map((row) => {
      const output: TableRow = {
        [RECORD_ID_COLUMN]: String(row[RECORD_ID_COLUMN]),
      };
      for (const column of definition.columns) {
        output[column.name] = cellFromSqlValue(
          column.type,
          (row[column.name] ?? null) as SqlValueType,
        );
      }
      return output;
    });
  }

  /** The ordered column names of a table, Record ID first. */
  columnNames(tableName: string): string[] {
    const definition = this.#requireDefinition(tableName);
    return [
      RECORD_ID_COLUMN,
      ...definition.columns.map((column) => column.name),
    ];
  }

  /** Serialize the whole database back to bytes for saving. */
  serialize(): Uint8Array {
    return this.#sql.serialize();
  }

  /** Release the database and its memory. */
  close(): void {
    this.#sql.close();
  }

  /** Direct engine access for advanced callers and the table bridge. */
  get sql(): SqlDatabase {
    return this.#sql;
  }

  #tableExists(tableName: string): boolean {
    const found = this.#sql.selectValue(
      `SELECT count(*) FROM ${quoteIdentifier(TABLE_REGISTRY_TABLE)} WHERE name = ?;`,
      [tableName],
    );
    return found === 1;
  }

  #assertTableAbsent(tableName: string): void {
    if (this.#tableExists(tableName)) {
      throw new ConsultChimpsError(
        "DB_TABLE_EXISTS",
        `The table "${tableName}" already exists.`,
        { details: { table: tableName } },
      );
    }
  }

  #requireDefinition(tableName: string): StoredDefinition {
    const value = this.#sql.selectValue(
      `SELECT definition FROM ${quoteIdentifier(TABLE_REGISTRY_TABLE)} WHERE name = ?;`,
      [tableName],
    );
    if (typeof value !== "string") {
      throw new ConsultChimpsError(
        "DB_TABLE_NOT_FOUND",
        `The table "${tableName}" does not exist.`,
        { details: { table: tableName } },
      );
    }
    return JSON.parse(value) as StoredDefinition;
  }
}
