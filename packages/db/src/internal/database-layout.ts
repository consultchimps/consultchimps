import { isConsultChimpsError } from "@consultchimps/core";

import { databaseError } from "../errors.js";
import {
  identifierKey,
  quoteIdentifier,
  type ColumnDefinition,
  type DatabaseFormat,
  type TableSchema,
} from "../schema.js";
import { parseDatabaseSchema } from "../validators.js";
import { COUNTERS_TABLE, TABLE_REGISTRY_TABLE } from "../metadata.js";
import type {
  DatabaseEngine,
  EngineRow,
  EngineTransaction,
  EngineValue,
} from "./engine.js";
import { DATABASE_STORAGE } from "./storage-layouts.js";
import { validateInternalTables } from "./storage-schema.js";

export async function assertNoSqliteTableTriggers(
  transaction: EngineTransaction,
  format: DatabaseFormat,
  tables: readonly string[],
  error: {
    readonly code: string;
    readonly message: string;
  },
): Promise<void> {
  if (format !== "sqlite") return;
  const protectedTables = new Set(tables.map(identifierKey));
  const rows = await transaction.query(
    "SELECT name, tbl_name FROM sqlite_schema WHERE type = 'trigger' UNION ALL SELECT name, tbl_name FROM sqlite_temp_schema WHERE type = 'trigger' ORDER BY name",
  );
  const trigger = rows.find(
    (row) =>
      typeof row["tbl_name"] === "string" &&
      protectedTables.has(identifierKey(row["tbl_name"])),
  );
  if (trigger === undefined) return;
  throw databaseError(error.code, error.message, {
    table:
      typeof trigger["tbl_name"] === "string" ? trigger["tbl_name"] : undefined,
    trigger: typeof trigger["name"] === "string" ? trigger["name"] : undefined,
  });
}

export async function assertNoManagedDatabaseTriggers(
  transaction: EngineTransaction,
  format: DatabaseFormat,
): Promise<void> {
  if (format !== "sqlite") return;
  const registered = await transaction.query(
    `SELECT table_name FROM ${TABLE_REGISTRY_TABLE}`,
  );
  await assertNoSqliteTableTriggers(
    transaction,
    format,
    [
      ...Object.values(DATABASE_STORAGE.tables).map(({ name }) => name),
      ...registered.flatMap((row) =>
        typeof row["table_name"] === "string" ? [row["table_name"]] : [],
      ),
    ],
    {
      code: "DB_SCHEMA_DRIFT",
      message:
        "A trigger was added to a managed database table outside the database operations. Remove the trigger or restore a verified database copy before writing.",
    },
  );
}

function corruptDatabase(details?: Record<string, unknown>, cause?: unknown) {
  return databaseError(
    "DB_CORRUPT_DATABASE",
    "The database is incomplete or damaged. Restore a verified database copy before retrying.",
    details,
    cause,
  );
}

function storedBigInt(value: unknown, field: string, minimum: bigint): bigint {
  const parsed =
    typeof value === "bigint"
      ? value
      : typeof value === "number" && Number.isSafeInteger(value)
        ? BigInt(value)
        : typeof value === "string" && /^-?\d+$/u.test(value)
          ? BigInt(value)
          : undefined;
  if (parsed !== undefined && parsed >= minimum) return parsed;
  throw corruptDatabase({ field });
}

export function parseStoredTableSchema(
  schemaText: string,
  expectedName: string,
): TableSchema {
  try {
    const value: unknown = JSON.parse(schemaText);
    const parsed = parseDatabaseSchema({ version: 1, tables: [value] });
    const stored = parsed.tables[0];
    if (stored === undefined || stored.name !== expectedName)
      throw new Error("stored schema table name mismatch");
    return stored;
  } catch (cause) {
    throw databaseError(
      "DB_CORRUPT_DATABASE",
      `The stored schema for "${expectedName}" is invalid.`,
      { table: expectedName },
      cause,
    );
  }
}

export function storageType(
  format: DatabaseFormat,
  column: ColumnDefinition,
): string {
  switch (column.type) {
    case "text":
      return "VARCHAR";
    case "timestamp":
      return format === "duckdb" ? "TIMESTAMP" : "VARCHAR";
    case "integer":
      return "BIGINT";
    case "real":
      return "DOUBLE";
    case "boolean":
      return format === "duckdb" ? "BOOLEAN" : "BIGINT";
    case "date":
      return format === "duckdb" ? "DATE" : "VARCHAR";
    case "decimal":
      return format === "duckdb"
        ? `DECIMAL(${column.precision}, ${column.scale})`
        : "VARCHAR";
  }
}

export interface RegisteredTable {
  readonly schema: TableSchema;
  readonly schemaVersion: bigint;
}

const REQUIRED_COUNTERS = [
  "source_file",
  "capture",
  "plan",
  "import",
  "delivery",
  "imported_row",
] as const;

export async function readRegisteredTables(
  transaction: EngineTransaction,
): Promise<readonly RegisteredTable[]> {
  const counterRows = await transaction.query(
    `SELECT counter_name, next_value FROM ${COUNTERS_TABLE}`,
  );
  const counters = new Map(
    counterRows.flatMap((row) =>
      typeof row["counter_name"] === "string"
        ? [[row["counter_name"], row["next_value"]] as const]
        : [],
    ),
  );
  for (const counter of REQUIRED_COUNTERS)
    storedBigInt(counters.get(counter), `${counter} counter`, 1n);
  const rows = await transaction.query(
    `SELECT table_name, schema_json, schema_version, next_record_id FROM ${TABLE_REGISTRY_TABLE} ORDER BY table_name`,
  );
  return rows.map((row) => {
    const name = row["table_name"];
    const schemaText = row["schema_json"];
    if (
      typeof name !== "string" ||
      name.trim().length === 0 ||
      typeof schemaText !== "string"
    )
      throw corruptDatabase({ field: "table registry" });
    storedBigInt(row["next_record_id"], "Record ID counter", 1n);
    return {
      schema: parseStoredTableSchema(schemaText, name),
      schemaVersion: storedBigInt(row["schema_version"], "schema version", 1n),
    };
  });
}

export async function assertRegisteredStorage(
  transaction: EngineTransaction,
  format: DatabaseFormat,
  tables: readonly TableSchema[],
  options: { readonly allowExtraColumns?: boolean } = {},
): Promise<void> {
  if (tables.length > 0) {
    const names = tables.map(({ name }) => name);
    const placeholders = names.map(() => "?").join(", ");
    const rows = await transaction.query(
      format === "sqlite"
        ? `SELECT name AS table_name FROM sqlite_master WHERE type = 'table' AND name IN (${placeholders})`
        : `SELECT table_name FROM information_schema.tables WHERE table_schema = 'main' AND table_type = 'BASE TABLE' AND table_name IN (${placeholders})`,
      names,
    );
    const baseTables = new Set(
      rows.flatMap((row) =>
        typeof row["table_name"] === "string" ? [row["table_name"]] : [],
      ),
    );
    const missing = names.find((name) => !baseTables.has(name));
    if (missing !== undefined)
      throw databaseError(
        "DB_SCHEMA_DRIFT",
        "A managed table was changed outside the database schema operations. Restore its declared layout or copy the changed data into a new managed table before importing.",
        { table: missing },
      );
  }
  for (const table of tables) {
    const actual = await transaction.query(
      `PRAGMA ${format === "sqlite" ? "table_xinfo" : "table_info"}(${quoteIdentifier(table.name)})`,
    );
    const actualByName = new Map(
      actual.flatMap((column) =>
        typeof column["name"] === "string"
          ? [[identifierKey(column["name"]), column] as const]
          : [],
      ),
    );
    const expected = new Map([
      ["record_id", "VARCHAR"],
      ["_imported_row_id", "BIGINT"],
      ["_import_id", "VARCHAR"],
      ["_source_file_id", "VARCHAR"],
      ["_source_selection", "VARCHAR"],
      ["_source_row", "BIGINT"],
      ...table.columns.map(
        (column) =>
          [identifierKey(column.name), storageType(format, column)] as const,
      ),
    ]);
    const matches =
      (options.allowExtraColumns === true || actual.length === expected.size) &&
      [...expected].every(([expectedName, expectedType]) => {
        const column = actualByName.get(expectedName);
        if (column === undefined) return false;
        const name = column["name"];
        const type = column["type"];
        if (
          typeof name !== "string" ||
          typeof type !== "string" ||
          expectedType.replaceAll(" ", "") !==
            type.toUpperCase().replaceAll(" ", "")
        )
          return false;
        const definition = table.columns.find(
          (candidate) => identifierKey(candidate.name) === identifierKey(name),
        );
        return (
          (column["dflt_value"] === null ||
            column["dflt_value"] === undefined) &&
          (column["hidden"] === undefined ||
            String(column["hidden"]) === "0") &&
          (identifierKey(name) === "record_id" ||
            definition?.nullable === false) ===
            (column["notnull"] === true ||
              column["notnull"] === 1n ||
              column["notnull"] === 1)
        );
      });
    if (!matches)
      throw databaseError(
        "DB_SCHEMA_DRIFT",
        "A managed table was changed outside the database schema operations. Restore its declared layout or copy the changed data into a new managed table before importing.",
        { table: table.name },
      );
  }
}

export async function validateRegisteredTables(
  transaction: EngineTransaction,
  format: DatabaseFormat,
  options: { readonly allowExtraColumns?: boolean } = {},
): Promise<readonly TableSchema[]> {
  const schemas = (await readRegisteredTables(transaction)).map(
    ({ schema }) => schema,
  );
  await assertRegisteredStorage(transaction, format, schemas, options);
  return schemas;
}

export async function queryDatabaseMetadata(
  engine: DatabaseEngine,
  sql: string,
  values?: readonly EngineValue[],
): Promise<readonly EngineRow[]> {
  try {
    return await engine.query(sql, values);
  } catch (cause) {
    if (isConsultChimpsError(cause)) throw cause;
    if (cause instanceof Error && cause.name === "AbortError") throw cause;
    throw corruptDatabase(undefined, cause);
  }
}

export async function validateDatabaseLayout(
  engine: DatabaseEngine,
): Promise<void> {
  await validateInternalTables({
    query: (sql, values) => queryDatabaseMetadata(engine, sql, values),
    format: engine.format,
    schema: DATABASE_STORAGE,
    invalid: corruptDatabase,
  });
}
