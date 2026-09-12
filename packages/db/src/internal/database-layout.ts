import { isConsultChimpsError } from "@consultchimps/core";

import { databaseError } from "../errors.js";
import {
  APPLICATION_TABLE,
  CAPTURE_ROW_TABLE,
  CAPTURE_TABLE,
  COUNTERS_TABLE,
  DATABASE_METADATA_TABLE,
  DELIVERY_MEMBERSHIP_TABLE,
  DELIVERY_TABLE,
  IMPORT_REQUEST_TABLE,
  PLAN_TABLE,
  SOURCE_CONTENT_TABLE,
  SOURCE_FILE_TABLE,
  SOURCE_NAME_TABLE,
  TABLE_REGISTRY_TABLE,
} from "../metadata.js";
import type { DatabaseEngine, EngineRow, EngineValue } from "./engine.js";

const REQUIRED_DATABASE_SCHEMA = [
  {
    table: DATABASE_METADATA_TABLE,
    columns: ["database_id", "format", "format_version", "revision"],
  },
  {
    table: TABLE_REGISTRY_TABLE,
    columns: ["table_name", "schema_json", "schema_version", "next_record_id"],
  },
  {
    table: COUNTERS_TABLE,
    columns: ["counter_name", "next_value"],
  },
  {
    table: SOURCE_CONTENT_TABLE,
    columns: ["content_hash", "byte_count"],
  },
  {
    table: SOURCE_FILE_TABLE,
    columns: ["source_file_id", "content_hash", "display_name"],
  },
  {
    table: SOURCE_NAME_TABLE,
    columns: ["source_file_id", "display_name"],
  },
  {
    table: CAPTURE_TABLE,
    columns: [
      "capture_id",
      "source_file_id",
      "source_key",
      "selection_key",
      "selection_label",
      "reader_version",
      "state",
      "row_count",
      "columns_json",
    ],
  },
  {
    table: PLAN_TABLE,
    columns: [
      "plan_id",
      "plan_revision",
      "baseline_revision",
      "state",
      "recipe_json",
      "conflicts_json",
      "decisions_json",
      "bindings_json",
    ],
  },
  {
    table: CAPTURE_ROW_TABLE,
    columns: ["capture_id", "source_row", "values_json"],
  },
  {
    table: APPLICATION_TABLE,
    columns: [
      "import_id",
      "application_key",
      "request_id",
      "capture_id",
      "table_name",
      "plan_id",
      "plan_revision",
      "row_count",
    ],
  },
  {
    table: IMPORT_REQUEST_TABLE,
    columns: [
      "request_id",
      "plan_id",
      "plan_revision",
      "import_ids_json",
      "capture_ids_json",
      "row_count",
    ],
  },
  {
    table: DELIVERY_TABLE,
    columns: ["delivery_id", "request_id", "context_json"],
  },
  {
    table: DELIVERY_MEMBERSHIP_TABLE,
    columns: ["delivery_id", "capture_id"],
  },
] as const;

function corruptDatabase(details?: Record<string, unknown>, cause?: unknown) {
  return databaseError(
    "DB_CORRUPT_DATABASE",
    "The database is incomplete or damaged. Restore a verified database copy before retrying.",
    details,
    cause,
  );
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
  const requiredTables = REQUIRED_DATABASE_SCHEMA.map(({ table }) => table);
  const placeholders = requiredTables.map(() => "?").join(", ");
  const tableRows = await queryDatabaseMetadata(
    engine,
    engine.format === "sqlite"
      ? `SELECT name AS table_name FROM sqlite_master WHERE type = 'table' AND name IN (${placeholders})`
      : `SELECT table_name FROM information_schema.tables WHERE table_schema = 'main' AND table_type = 'BASE TABLE' AND table_name IN (${placeholders})`,
    requiredTables,
  );
  const baseTables = new Set(
    tableRows.flatMap((row) =>
      typeof row["table_name"] === "string" ? [row["table_name"]] : [],
    ),
  );
  const missingTables = requiredTables.filter(
    (table) => !baseTables.has(table),
  );
  if (missingTables.length > 0) {
    throw corruptDatabase({ missingTables });
  }
  for (const required of REQUIRED_DATABASE_SCHEMA) {
    const rows = await queryDatabaseMetadata(
      engine,
      engine.format === "sqlite"
        ? "SELECT name AS column_name FROM pragma_table_info(?)"
        : "SELECT column_name FROM information_schema.columns WHERE table_schema = 'main' AND table_name = ? ORDER BY ordinal_position",
      [required.table],
    );
    const columns = new Set(
      rows.flatMap((row) =>
        typeof row["column_name"] === "string" ? [row["column_name"]] : [],
      ),
    );
    const missingColumns = required.columns.filter(
      (column) => !columns.has(column),
    );
    if (missingColumns.length > 0) {
      throw corruptDatabase({ table: required.table, missingColumns });
    }
  }
}
