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

// A generous but bounded cap on Record ID zero-padding, so a mistaken
// configuration cannot drive String.padStart into an enormous allocation.
const MAX_RECORD_ID_PADDING = 64;

/** The stored shape of a table definition in the registry. */
interface StoredDefinition {
  columns: ColumnDefinition[];
  foreignKeys: TableSchema["foreignKeys"];
  recordId: RecordIdConfig;
}

const VALID_COLUMN_TYPES: ReadonlySet<string> = new Set([
  "text",
  "integer",
  "real",
  "boolean",
  "date",
]);

// Parse and validate one registry definition. A damaged or externally edited
// file can hold malformed JSON or valid JSON of the wrong shape, so both become
// a stable corruption error here rather than a raw SyntaxError or a later
// TypeError from, say, definition.columns.map.
function parseDefinition(json: string, tableName: string): StoredDefinition {
  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch (error) {
    throw new ConsultChimpsError(
      "DB_CORRUPT_WORKSPACE",
      `The stored definition for "${tableName}" is not valid JSON, so the database may be damaged.`,
      { cause: error, details: { table: tableName } },
    );
  }

  const fail = (detail: string): never => {
    throw new ConsultChimpsError(
      "DB_CORRUPT_WORKSPACE",
      `The stored definition for "${tableName}" is not valid (${detail}), so the database may be damaged.`,
      { details: { table: tableName } },
    );
  };
  const asObject = (candidate: unknown): Record<string, unknown> => {
    if (typeof candidate !== "object" || candidate === null) {
      fail("it is not an object");
    }
    return candidate as Record<string, unknown>;
  };

  const definition = asObject(parsed);
  if (!Array.isArray(definition["columns"])) {
    fail("its columns are missing");
  }
  for (const rawColumn of definition["columns"] as unknown[]) {
    const column = asObject(rawColumn);
    if (typeof column["name"] !== "string" || column["name"].trim() === "") {
      fail("a column name is missing");
    }
    const columnType = column["type"];
    if (typeof columnType !== "string" || !VALID_COLUMN_TYPES.has(columnType)) {
      fail("a column type is unknown");
    }
    if (
      column["nullable"] !== undefined &&
      typeof column["nullable"] !== "boolean"
    ) {
      fail("a column nullable flag is not a boolean");
    }
  }
  if (!Array.isArray(definition["foreignKeys"])) {
    fail("its foreign keys are missing");
  }
  for (const rawForeignKey of definition["foreignKeys"] as unknown[]) {
    const foreignKey = asObject(rawForeignKey);
    if (
      typeof foreignKey["column"] !== "string" ||
      typeof foreignKey["referencesTable"] !== "string"
    ) {
      fail("a foreign key is malformed");
    }
  }
  const recordId = asObject(definition["recordId"]);
  if (typeof recordId["prefix"] !== "string") {
    fail("its Record ID prefix is missing");
  }
  const padding = recordId["padding"];
  if (
    typeof padding !== "number" ||
    !Number.isInteger(padding) ||
    padding < 0 ||
    padding > MAX_RECORD_ID_PADDING
  ) {
    fail("its Record ID padding is out of range");
  }
  if (
    recordId["separator"] !== undefined &&
    typeof recordId["separator"] !== "string"
  ) {
    fail("its Record ID separator is not text");
  }

  // Reapply the identifier rules that table creation enforces, so an externally
  // edited definition with an unsafe, reserved, or duplicate name is reported as
  // corruption here rather than surfacing a raw SQLite syntax error when the
  // name later reaches quoteIdentifier.
  const checkIdentifier = (name: string, role: "table" | "column"): void => {
    try {
      assertSafeIdentifier(name, role);
    } catch {
      fail(`its ${role} name "${name}" is not a valid identifier`);
    }
  };
  checkIdentifier(tableName, "table");
  const seenColumns = new Set<string>();
  for (const rawColumn of definition["columns"] as Array<
    Record<string, unknown>
  >) {
    const columnName = rawColumn["name"] as string;
    checkIdentifier(columnName, "column");
    if (identifierKey(columnName) === identifierKey(RECORD_ID_COLUMN)) {
      fail("a column uses the reserved Record ID name");
    }
    const key = identifierKey(columnName);
    if (seenColumns.has(key)) {
      fail("a column name is duplicated");
    }
    seenColumns.add(key);
  }
  for (const rawForeignKey of definition["foreignKeys"] as Array<
    Record<string, unknown>
  >) {
    checkIdentifier(rawForeignKey["column"] as string, "column");
    checkIdentifier(rawForeignKey["referencesTable"] as string, "table");
  }

  return parsed as StoredDefinition;
}

// Translate a failure while opening bytes: a ConsultChimpsError (this package's
// own validation) passes through, an sql.js "not a database / malformed" error
// (a wrong file or a damaged save, an expected failure) becomes a stable error,
// and anything else (a wasm-loading or programming error) is left as itself so
// it stays distinguishable.
function translateOpenError(error: unknown): unknown {
  if (isConsultChimpsError(error)) {
    return error;
  }
  const message = error instanceof Error ? error.message : String(error);
  if (/not a database|malformed|encrypted|file is not/i.test(message)) {
    return new ConsultChimpsError(
      "DB_INVALID_DATABASE_FILE",
      "These bytes are not a readable database file; the file may be the wrong one or damaged.",
      { cause: error },
    );
  }
  return error;
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
    let sql: SqlDatabase;
    try {
      sql = await loadSqlDatabase(bytes, config);
    } catch (error) {
      // Invalid bytes can fail either here or at the first query below.
      throw translateOpenError(error);
    }
    const database = new Database(sql);
    try {
      sql.run("PRAGMA foreign_keys = ON;");
      database.#assertMetadataPresent();
      database.#assertSupportedSchemaVersion();
      // Reject a registry that holds two table names differing only by case,
      // which a binary-collated primary key permits in an externally edited file
      // but this package's case-insensitive lookups cannot resolve.
      database.#assertRegistryDistinct();
      // Force every stored definition to parse now, so a damaged registry is
      // reported as a corruption error while the handle can still be closed,
      // rather than throwing later from getSchema or a record operation.
      database.getSchema();
    } catch (error) {
      // A rejected open must not leak the sql.js allocation the load created,
      // since the caller never receives a handle to close.
      sql.close();
      throw translateOpenError(error);
    }
    return database;
  }

  #assertRegistryDistinct(): void {
    const rows = this.#sql.select(
      `SELECT name FROM ${quoteIdentifier(TABLE_REGISTRY_TABLE)};`,
    );
    const seen = new Set<string>();
    for (const row of rows) {
      const key = identifierKey(String(row["name"]));
      if (seen.has(key)) {
        throw new ConsultChimpsError(
          "DB_CORRUPT_WORKSPACE",
          "The database holds two tables whose names differ only by case, so it may be damaged.",
          { details: {} },
        );
      }
      seen.add(key);
    }
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
    if (
      !Number.isInteger(config.padding) ||
      config.padding < 0 ||
      config.padding > MAX_RECORD_ID_PADDING
    ) {
      throw new ConsultChimpsError(
        "DB_INVALID_RECORD_ID_CONFIG",
        `The Record ID padding for "${table}" must be a whole number from 0 to ${MAX_RECORD_ID_PADDING}.`,
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
        // Own-property read so a column named like an Object.prototype member
        // never picks up an inherited value; "__proto__" is refused as a column
        // name, so assigning declared names to this plain object is safe.
        const stored = Object.prototype.hasOwnProperty.call(row, column.name)
          ? (row[column.name] as SqlValueType)
          : null;
        output[column.name] = cellFromSqlValue(column.type, stored);
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
    const bytes = this.#sql.serialize();
    // sql.js export() closes and reopens the underlying connection, which resets
    // connection-local PRAGMAs, so re-enable foreign key enforcement for editing
    // that continues after a save or autosave.
    this.#sql.run("PRAGMA foreign_keys = ON;");
    return bytes;
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
