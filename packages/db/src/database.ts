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

/** The outcome of updating a record: which record was written, and what it now holds. */
export interface UpdatedRecord {
  /** The Record ID of the updated record, as it is stored. */
  recordId: string;
  /**
   * The values now stored for the columns that were written, read back and
   * coerced to their declared types. A caller that displays the record can show
   * what the database kept rather than what it sent, which differ whenever a
   * value was converted (the text "yes" into a boolean column, say).
   */
  values: Record<string, CellValue>;
}

/** Narrow what `Database.readRecords` reads. Omit either field for all. */
export interface ReadRecordsOptions {
  /**
   * Columns to read besides the Record ID, which is always read first. Names
   * match case-insensitively; an unknown or repeated name is refused.
   */
  readonly columns?: readonly string[];
  /** Read at most this many records, in storage order. Zero reads none. */
  readonly limit?: number;
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
  const prefix = recordId["prefix"];
  if (typeof prefix !== "string" || prefix.trim() === "") {
    fail("its Record ID prefix is missing or empty");
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
      database.#validateLoadedWorkspace();
    } catch (error) {
      // A rejected open must not leak the sql.js allocation the load created,
      // since the caller never receives a handle to close.
      sql.close();
      throw translateOpenError(error);
    }
    return database;
  }

  /**
   * The single open-time validation routine. The create path enforces the
   * schema invariants when a workspace is built; opening a file cannot trust the
   * stored registry or physical schema, so this re-applies the same invariants
   * to the loaded state and cross-checks the stored definitions against the real
   * tables. It deliberately validates schema-level invariants only (shape,
   * identifiers, types, constraints), not a full row scan: the threat model is a
   * trusted single writer and a wrong or damaged file, not adversarial data.
   */
  #validateLoadedWorkspace(): void {
    // 1. The file is a ConsultChimps workspace with the expected metadata tables
    //    and layout, at a schema version this build understands.
    this.#assertMetadataPresent();
    this.#assertSupportedSchemaVersion();
    // 2. The registry has no two table names differing only by case, which a
    //    binary-collated key permits but the case-insensitive lookups cannot
    //    resolve.
    this.#assertRegistryDistinct();
    // 3. Every stored definition parses and satisfies the create-path invariants
    //    (identifiers, Record ID prefix and padding, column types), and matches
    //    its physical table (columns, types, NOT NULL, and declared foreign
    //    keys). getSchema() performs the parse-and-invariant pass; the physical
    //    cross-check follows.
    this.#assertPhysicalSchemaMatches();
  }

  #assertPhysicalSchemaMatches(): void {
    const corrupt = (detail: string): never => {
      throw new ConsultChimpsError(
        "DB_CORRUPT_WORKSPACE",
        `The stored schema does not match the database (${detail}), so it may be damaged.`,
        { details: {} },
      );
    };

    for (const schema of this.getSchema()) {
      const physicalColumns = new Map<
        string,
        { type: string; notNull: boolean }
      >();
      for (const row of this.#sql.select(
        `PRAGMA table_info(${quoteIdentifier(schema.name)});`,
      )) {
        physicalColumns.set(identifierKey(String(row["name"])), {
          type: String(row["type"]).toUpperCase(),
          notNull: Number(row["notnull"]) === 1,
        });
      }
      if (physicalColumns.size === 0) {
        corrupt(`table "${schema.name}" is missing`);
      }
      if (!physicalColumns.has(identifierKey(RECORD_ID_COLUMN))) {
        corrupt(`table "${schema.name}" has no Record ID column`);
      }
      for (const column of schema.columns) {
        const physical = physicalColumns.get(identifierKey(column.name));
        if (physical === undefined) {
          corrupt(`table "${schema.name}" is missing column "${column.name}"`);
        } else {
          if (physical.type !== sqlStorageClass(column.type)) {
            corrupt(
              `column "${column.name}" on "${schema.name}" has the wrong type`,
            );
          }
          // A column declared non-nullable must carry the physical NOT NULL, or
          // an insert that omits it would silently store NULL against the schema.
          if (column.nullable === false && !physical.notNull) {
            corrupt(
              `column "${column.name}" on "${schema.name}" is not marked NOT NULL as declared`,
            );
          }
        }
      }
      const foreignKeyList = this.#sql.select(
        `PRAGMA foreign_key_list(${quoteIdentifier(schema.name)});`,
      );
      for (const foreignKey of schema.foreignKeys) {
        const present = foreignKeyList.some(
          (row) =>
            sameIdentifier(String(row["from"]), foreignKey.column) &&
            sameIdentifier(String(row["table"]), foreignKey.referencesTable) &&
            sameIdentifier(String(row["to"]), RECORD_ID_COLUMN),
        );
        if (!present) {
          corrupt(
            `table "${schema.name}" is missing a declared foreign key on "${foreignKey.column}"`,
          );
        }
      }
    }
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
    const notAWorkspace = (): never => {
      throw new ConsultChimpsError(
        "DB_NOT_A_WORKSPACE",
        "This database file was not created by ConsultChimps: its schema metadata tables are missing or malformed.",
        { details: { expectedTables: [METADATA_TABLE, TABLE_REGISTRY_TABLE] } },
      );
    };
    const found = this.#sql.selectValue(
      "SELECT count(*) FROM sqlite_schema WHERE type = 'table' AND name IN (?, ?);",
      [METADATA_TABLE, TABLE_REGISTRY_TABLE],
    );
    if (found !== 2) {
      notAWorkspace();
    }
    // The tables can exist with the wrong columns (a damaged or unrelated file),
    // which would otherwise surface as a raw "no such column" error from the
    // first metadata query. Confirm the expected layout up front.
    const hasColumns = (table: string, expected: string[]): boolean => {
      const columns = new Set(
        this.#sql
          .select(`PRAGMA table_info(${quoteIdentifier(table)});`)
          .map((row) => identifierKey(String(row["name"]))),
      );
      return expected.every((column) => columns.has(identifierKey(column)));
    };
    if (
      !hasColumns(METADATA_TABLE, ["key", "value"]) ||
      !hasColumns(TABLE_REGISTRY_TABLE, ["name", "definition", "next_counter"])
    ) {
      notAWorkspace();
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

  /**
   * Update an existing record, found by its Record ID. The values are keyed by
   * column name and go through the same conversion as an insert, so a value
   * that does not fit the column's declared type is refused here rather than
   * stored in a shape the read side would later call corrupt. The Record ID is
   * assigned once and cannot be among them.
   *
   * Passing no values checks that the record exists and writes nothing, which
   * is what "update these zero columns" means; it is not an error.
   */
  updateRecord(
    tableName: string,
    recordId: string,
    values: Readonly<Record<string, CellValue>>,
  ): UpdatedRecord {
    const { name, definition } = this.#requireDefinition(tableName);
    // Columns are matched case-insensitively, like SQLite identifiers.
    const columnByName = new Map(
      definition.columns.map((column) => [identifierKey(column.name), column]),
    );

    const updateColumns: ColumnDefinition[] = [];
    const updateValues: SqlValueType[] = [];
    const providedColumns = new Set<string>();
    for (const [key, value] of Object.entries(values)) {
      // The storage trigger refuses this too, but the caller deserves a stable
      // code and a sentence rather than a raw engine abort.
      if (sameIdentifier(key, RECORD_ID_COLUMN)) {
        throw new ConsultChimpsError(
          "DB_RECORD_ID_IMMUTABLE",
          `The Record ID for "${name}" is assigned once and cannot be changed.`,
          { details: { table: name } },
        );
      }
      const column = columnByName.get(identifierKey(key));
      if (column === undefined) {
        throw new ConsultChimpsError(
          "DB_UNKNOWN_COLUMN",
          `The table "${name}" has no column "${key}".`,
          { details: { table: name, column: key } },
        );
      }
      // Two keys that resolve to the same column (differing only by case) would
      // both land in the SET clause, where the last one silently wins; reject
      // the ambiguity instead, exactly as an insert does.
      const columnKey = identifierKey(column.name);
      if (providedColumns.has(columnKey)) {
        throw new ConsultChimpsError(
          "DB_DUPLICATE_UPDATE_COLUMN",
          `The update for "${name}" gives the column "${column.name}" more than once.`,
          { details: { table: name, column: column.name } },
        );
      }
      providedColumns.add(columnKey);
      updateColumns.push(column);
      updateValues.push(
        this.#convertCell(column.type, value, name, column.name),
      );
    }

    // An UPDATE that matches no row is not an error in SQL, so a stale or
    // mistyped Record ID would otherwise report success having changed nothing.
    // The Record ID is data, not an identifier, so it is matched exactly.
    const storedId = this.#sql.selectValue(
      `SELECT ${quoteIdentifier(RECORD_ID_COLUMN)} FROM ${quoteIdentifier(name)} WHERE ${quoteIdentifier(RECORD_ID_COLUMN)} = ?;`,
      [recordId],
    );
    if (typeof storedId !== "string") {
      throw new ConsultChimpsError(
        "DB_RECORD_NOT_FOUND",
        `The table "${name}" has no record with the Record ID "${recordId}".`,
        { details: { table: name, recordId } },
      );
    }

    if (updateColumns.length > 0) {
      const assignments = updateColumns
        .map((column) => `${quoteIdentifier(column.name)} = ?`)
        .join(", ");
      try {
        this.#sql.run(
          `UPDATE ${quoteIdentifier(name)} SET ${assignments} WHERE ${quoteIdentifier(RECORD_ID_COLUMN)} = ?;`,
          [...updateValues, storedId],
        );
      } catch (error) {
        throw translateConstraintError(error, name);
      }
    }

    return {
      recordId: storedId,
      values: this.#readColumns(name, storedId, updateColumns),
    };
  }

  // Read the given columns of one record back through the declared-type
  // conversion, so the caller is told what is stored rather than what it sent.
  /**
   * Resolve the columns a read asked for, in the order asked. Names match
   * case-insensitively like every other identifier here. The Record ID column
   * is always read, so naming it is allowed and changes nothing; an unknown
   * column or one named twice is refused, since either means the caller and
   * the schema disagree.
   */
  #selectColumns(
    tableName: string,
    definition: StoredDefinition,
    requested: readonly string[],
  ): ColumnDefinition[] {
    const byKey = new Map(
      definition.columns.map((column) => [identifierKey(column.name), column]),
    );
    const selected: ColumnDefinition[] = [];
    const seen = new Set<string>();
    for (const key of requested) {
      // Resolve the declared name first, so the Record ID takes part in
      // duplicate tracking like any other column even though it is never
      // added to the projection (it is always read).
      let resolved: string;
      let column: ColumnDefinition | undefined;
      if (sameIdentifier(key, RECORD_ID_COLUMN)) {
        resolved = RECORD_ID_COLUMN;
      } else {
        column = byKey.get(identifierKey(key));
        if (column === undefined) {
          throw new ConsultChimpsError(
            "DB_UNKNOWN_COLUMN",
            `The table "${tableName}" has no column "${key}".`,
            { details: { table: tableName, column: key } },
          );
        }
        resolved = column.name;
      }
      const columnKey = identifierKey(resolved);
      if (seen.has(columnKey)) {
        throw new ConsultChimpsError(
          "DB_DUPLICATE_READ_COLUMN",
          `The read from "${tableName}" names the column "${resolved}" more than once.`,
          { details: { table: tableName, column: resolved } },
        );
      }
      seen.add(columnKey);
      if (column !== undefined) {
        selected.push(column);
      }
    }
    return selected;
  }

  #readColumns(
    tableName: string,
    recordId: string,
    columns: readonly ColumnDefinition[],
  ): Record<string, CellValue> {
    const values: Record<string, CellValue> = {};
    if (columns.length === 0) {
      return values;
    }
    const selectList = columns
      .map((column) => quoteIdentifier(column.name))
      .join(", ");
    const rows = this.#sql.select(
      `SELECT ${selectList} FROM ${quoteIdentifier(tableName)} WHERE ${quoteIdentifier(RECORD_ID_COLUMN)} = ?;`,
      [recordId],
    );
    const row = rows[0];
    if (row === undefined) {
      throw new ConsultChimpsError(
        "DB_RECORD_NOT_FOUND",
        `The table "${tableName}" has no record with the Record ID "${recordId}".`,
        { details: { table: tableName, recordId } },
      );
    }
    for (const column of columns) {
      // Own-property read, for the same reason readRecords takes one: a column
      // named like an Object.prototype member must never pick up an inherited
      // value. "__proto__" is refused as a column name, so assigning declared
      // names to this plain object is safe.
      const stored = Object.prototype.hasOwnProperty.call(row, column.name)
        ? (row[column.name] as SqlValueType)
        : null;
      values[column.name] = cellFromSqlValue(column.type, stored);
    }
    return values;
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
  readRecords(tableName: string, options: ReadRecordsOptions = {}): TableRow[] {
    const { name, definition } = this.#requireDefinition(tableName);
    const columns =
      options.columns === undefined
        ? definition.columns
        : this.#selectColumns(name, definition, options.columns);
    const limit = options.limit;
    // Safe integer, not merely integer: 1e20 passes Number.isInteger but is
    // past what SQLite accepts as a LIMIT, and the raw engine error would
    // otherwise escape in place of the stable one.
    if (limit !== undefined && (!Number.isSafeInteger(limit) || limit < 0)) {
      throw new ConsultChimpsError(
        "DB_INVALID_LIMIT",
        `A record limit for "${name}" must be a whole number from zero up to the safe integer range, not ${String(limit)}.`,
        { details: { table: name, limit } },
      );
    }
    const columnNames = [
      RECORD_ID_COLUMN,
      ...columns.map((column) => column.name),
    ];
    const selectList = columnNames.map(quoteIdentifier).join(", ");
    const rows = this.#sql.select(
      `SELECT ${selectList} FROM ${quoteIdentifier(name)} ORDER BY rowid${limit === undefined ? "" : " LIMIT ?"};`,
      limit === undefined ? [] : [limit],
    );
    return rows.map((row) => {
      // The Record ID is a stored text value like any other, so validate it the
      // same way rather than String()-coercing a stray BLOB or number into a
      // plausible but wrong identifier.
      const storedId = row[RECORD_ID_COLUMN];
      if (typeof storedId !== "string") {
        throw new ConsultChimpsError(
          "DB_CORRUPT_STORED_VALUE",
          `A Record ID in "${tableName}" is not text, so the database may be damaged.`,
          { details: { table: tableName } },
        );
      }
      const output: TableRow = { [RECORD_ID_COLUMN]: storedId };
      for (const column of columns) {
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
