import { ConsultChimpsError, isConsultChimpsError } from "@consultchimps/core";
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
  identifierKey,
  quoteIdentifier,
  sameIdentifier,
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

// Parse one registry definition. A damaged or externally edited file can hold
// malformed JSON here, so a parse failure becomes a stable corruption error
// rather than a raw SyntaxError leaking from JSON.parse.
function parseDefinition(json: string, tableName: string): StoredDefinition {
  try {
    return JSON.parse(json) as StoredDefinition;
  } catch (error) {
    throw new ConsultChimpsError(
      "DB_CORRUPT_WORKSPACE",
      `The stored definition for "${tableName}" is not valid JSON, so the database may be damaged.`,
      { cause: error, details: { table: tableName } },
    );
  }
}

// Turn sql.js's generic constraint exception into a stable ConsultChimpsError so
// callers and future adapters can identify an expected input failure by code.
// The engine message carries only identifiers (table and column names), never
// the attempted values, so nothing confidential is surfaced.
function translateConstraintError(error: unknown, tableName: string): unknown {
  const message = error instanceof Error ? error.message : String(error);
  const target = message.includes(":")
    ? message.slice(message.indexOf(":") + 1).trim()
    : undefined;

  if (/FOREIGN KEY constraint failed/i.test(message)) {
    return new ConsultChimpsError(
      "DB_FOREIGN_KEY_VIOLATION",
      `A record in "${tableName}" references a Record ID that does not exist in the related table.`,
      { cause: error, details: { table: tableName } },
    );
  }
  if (/NOT NULL constraint failed/i.test(message)) {
    return new ConsultChimpsError(
      "DB_NOT_NULL_VIOLATION",
      `A required column on "${tableName}" was left empty${target ? ` (${target})` : ""}.`,
      { cause: error, details: { table: tableName, constraint: target } },
    );
  }
  if (/UNIQUE constraint failed/i.test(message)) {
    return new ConsultChimpsError(
      "DB_UNIQUE_VIOLATION",
      `A value on "${tableName}" must be unique${target ? ` (${target})` : ""}.`,
      { cause: error, details: { table: tableName, constraint: target } },
    );
  }
  return error;
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
    try {
      sql.run("PRAGMA foreign_keys = ON;");
      database.#assertMetadataPresent();
      database.#assertSupportedSchemaVersion();
      // Force every stored definition to parse now, so a damaged registry is
      // reported as a corruption error while the handle can still be closed,
      // rather than throwing later from getSchema or a record operation.
      database.getSchema();
    } catch (error) {
      // A rejected open must not leak the sql.js allocation the load created,
      // since the caller never receives a handle to close.
      sql.close();
      throw error;
    }
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
    // A missing row, malformed text, or a non-positive value all mean this is
    // not a schema state this version can safely assume, so reject them rather
    // than let a NaN comparison quietly fall through to version-1 handling.
    if (!Number.isInteger(version) || version < 1) {
      throw new ConsultChimpsError(
        "DB_UNSUPPORTED_SCHEMA_VERSION",
        "This database is missing a recognizable schema format version, so it may be corrupt or was not created by ConsultChimps.",
        {
          details: {
            fileVersion: version,
            supportedVersion: SCHEMA_FORMAT_VERSION,
          },
        },
      );
    }
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

  /**
   * The metadata-format version stored in this database file. Returns `NaN`
   * when the value is absent or malformed; a database opened through `open`
   * has already been rejected in that case.
   */
  schemaFormatVersion(): number {
    const value = this.#sql.selectValue(
      `SELECT value FROM ${quoteIdentifier(METADATA_TABLE)} WHERE key = ?;`,
      ["schema_format_version"],
    );
    return typeof value === "string" ? Number(value) : NaN;
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
      const key = identifierKey(column.name);
      // SQLite column names are case-insensitive, so any casing of the reserved
      // Record ID column would collide with the generated one; reject it here
      // with the stable error rather than let CREATE TABLE fail generically.
      if (key === identifierKey(RECORD_ID_COLUMN)) {
        throw new ConsultChimpsError(
          "DB_RESERVED_COLUMN",
          `The column name "${column.name}" is reserved for the generated Record ID.`,
          { details: { table: schema.name, column: column.name } },
        );
      }
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

    // Columns are matched case-insensitively, like SQLite identifiers, and are
    // unique case-insensitively (enforced above), so the lowercased name is an
    // unambiguous key.
    const columnByName = new Map(
      schema.columns.map((column) => [identifierKey(column.name), column]),
    );
    const foreignKeyClauses = schema.foreignKeys.map((foreignKey) => {
      const column = columnByName.get(identifierKey(foreignKey.column));
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
      // A self-reference is matched case-insensitively, like every other table
      // name, so "Customer" referencing "customer" is recognized as itself
      // even though the table is not yet in the registry.
      const isSelfReference = sameIdentifier(
        foreignKey.referencesTable,
        schema.name,
      );
      if (!isSelfReference && !this.#tableExists(foreignKey.referencesTable)) {
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
    const { definition } = this.#requireDefinition(tableName);
    // Columns are matched case-insensitively, like SQLite identifiers.
    const columnByName = new Map(
      definition.columns.map((column) => [identifierKey(column.name), column]),
    );

    const insertColumns: string[] = [RECORD_ID_COLUMN];
    const insertValues: SqlValueType[] = [];

    const counter = this.#nextCounter(tableName);
    const recordId = formatRecordId(definition.recordId, counter);
    insertValues.push(recordId);

    const providedColumns = new Set<string>();
    for (const [name, value] of Object.entries(values)) {
      if (sameIdentifier(name, RECORD_ID_COLUMN)) {
        throw new ConsultChimpsError(
          "DB_RECORD_ID_IS_GENERATED",
          `The Record ID for "${tableName}" is generated and cannot be supplied when inserting.`,
          { details: { table: tableName } },
        );
      }
      const column = columnByName.get(identifierKey(name));
      if (column === undefined) {
        throw new ConsultChimpsError(
          "DB_UNKNOWN_COLUMN",
          `The table "${tableName}" has no column "${name}".`,
          { details: { table: tableName, column: name } },
        );
      }
      // Two keys that resolve to the same column (differing only by case) would
      // both land in the INSERT, where SQLite keeps only the first and silently
      // drops the rest; reject the ambiguity instead.
      const columnKey = identifierKey(column.name);
      if (providedColumns.has(columnKey)) {
        throw new ConsultChimpsError(
          "DB_DUPLICATE_INSERT_COLUMN",
          `The insert for "${tableName}" gives the column "${column.name}" more than once.`,
          { details: { table: tableName, column: column.name } },
        );
      }
      providedColumns.add(columnKey);
      // Store under the schema's declared column name, not the caller's casing.
      insertColumns.push(column.name);
      insertValues.push(
        this.#convertCell(column.type, value, tableName, column.name),
      );
    }

    const placeholders = insertColumns.map(() => "?").join(", ");
    const quotedColumns = insertColumns.map(quoteIdentifier).join(", ");
    try {
      this.#sql.run(
        `INSERT INTO ${quoteIdentifier(tableName)} (${quotedColumns}) VALUES (${placeholders});`,
        insertValues,
      );
    } catch (error) {
      throw translateConstraintError(error, tableName);
    }

    this.#sql.run(
      `UPDATE ${quoteIdentifier(TABLE_REGISTRY_TABLE)} SET next_counter = ? WHERE name = ? COLLATE NOCASE;`,
      [counter + 1, tableName],
    );

    const rowId = this.#sql.selectValue("SELECT last_insert_rowid();");
    return {
      recordId,
      rowId: typeof rowId === "number" ? rowId : Number(rowId),
    };
  }

  // Convert one cell for storage, adding table and column context to a
  // conversion error while keeping the offending value (imported cell content)
  // out of it.
  #convertCell(
    type: ColumnDefinition["type"],
    value: CellValue,
    tableName: string,
    columnName: string,
  ): SqlValueType {
    try {
      return sqlValueFromCell(type, value);
    } catch (error) {
      if (
        isConsultChimpsError(error) &&
        (error.code === "DB_INVALID_BOOLEAN" ||
          error.code === "DB_INVALID_NUMBER")
      ) {
        throw new ConsultChimpsError(
          error.code,
          `${error.message} (table "${tableName}", column "${columnName}")`,
          {
            cause: error,
            details: { ...error.details, table: tableName, column: columnName },
          },
        );
      }
      throw error;
    }
  }

  #nextCounter(tableName: string): number {
    const value = this.#sql.selectValue(
      `SELECT next_counter FROM ${quoteIdentifier(TABLE_REGISTRY_TABLE)} WHERE name = ? COLLATE NOCASE;`,
      [tableName],
    );
    const counter = typeof value === "number" ? value : Number(value);
    // The counter is persisted state; a damaged or externally edited file could
    // hold a zero, negative, fractional, or out-of-range value that would form
    // an invalid Record ID. Reject it rather than generate one.
    if (!Number.isSafeInteger(counter) || counter < 1) {
      throw new ConsultChimpsError(
        "DB_CORRUPT_RECORD_COUNTER",
        `The Record ID counter for "${tableName}" is not a positive whole number, so the database may be damaged.`,
        { details: { table: tableName } },
      );
    }
    return counter;
  }

  /** The schema of a single table, read from the stored metadata. */
  getTableSchema(tableName: string): TableSchema {
    const { name, definition } = this.#requireDefinition(tableName);
    return {
      // The declared spelling, not the caller's casing.
      name,
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
      const name = String(row["name"]);
      const definition = parseDefinition(String(row["definition"]), name);
      return {
        name,
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
    const { definition } = this.#requireDefinition(tableName);
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
    const { definition } = this.#requireDefinition(tableName);
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

  // SQLite table names are case-insensitive, so the registry is queried the
  // same way: creating "customer" when "Customer" exists is a duplicate, and a
  // lookup finds a table whatever casing the caller passes.
  #tableExists(tableName: string): boolean {
    const found = this.#sql.selectValue(
      `SELECT count(*) FROM ${quoteIdentifier(TABLE_REGISTRY_TABLE)} WHERE name = ? COLLATE NOCASE;`,
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

  // Returns the stored declaration name (its original case) alongside the parsed
  // definition, so callers can report the declared spelling even when looked up
  // through the case-insensitive API.
  #requireDefinition(tableName: string): {
    name: string;
    definition: StoredDefinition;
  } {
    const rows = this.#sql.select(
      `SELECT name, definition FROM ${quoteIdentifier(TABLE_REGISTRY_TABLE)} WHERE name = ? COLLATE NOCASE;`,
      [tableName],
    );
    const row = rows[0];
    if (row === undefined) {
      throw new ConsultChimpsError(
        "DB_TABLE_NOT_FOUND",
        `The table "${tableName}" does not exist.`,
        { details: { table: tableName } },
      );
    }
    const storedName = String(row["name"]);
    return {
      name: storedName,
      definition: parseDefinition(String(row["definition"]), storedName),
    };
  }
}
