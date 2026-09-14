import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { NodeDuckDbEngine } from "../src/engines/duckdb/node.js";
import { NodeSqliteEngine } from "../src/engines/sqlite/node.js";
import type { DatabaseEngine, EngineRow } from "../src/internal/engine.js";
import {
  DATABASE_STORAGE,
  PREPARED_STORAGE,
} from "../src/internal/storage-layouts.js";
import {
  createInternalTables,
  insertInternalRow,
  internalCreateTableSql,
  internalCopyTables,
  validateInternalTables,
  type InternalStorageSchema,
} from "../src/internal/storage-schema.js";
import type { DatabaseFormat } from "../src/schema.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

function currentDatabaseDdl(format: DatabaseFormat): readonly string[] {
  return [
    "CREATE TABLE _consultchimps_database (database_id VARCHAR PRIMARY KEY, format VARCHAR NOT NULL, format_version BIGINT NOT NULL, revision BIGINT NOT NULL)",
    "CREATE TABLE _consultchimps_tables (table_name VARCHAR PRIMARY KEY, schema_json VARCHAR NOT NULL, schema_version BIGINT NOT NULL, next_record_id BIGINT NOT NULL)",
    "CREATE TABLE _consultchimps_counters (counter_name VARCHAR PRIMARY KEY, next_value BIGINT NOT NULL)",
    "CREATE TABLE _consultchimps_source_contents (content_hash VARCHAR PRIMARY KEY, byte_count BIGINT NOT NULL)",
    "CREATE TABLE _consultchimps_source_files (source_file_id VARCHAR PRIMARY KEY, content_hash VARCHAR NOT NULL, display_name VARCHAR NOT NULL, UNIQUE(content_hash))",
    "CREATE TABLE _consultchimps_source_names (source_file_id VARCHAR NOT NULL, display_name VARCHAR NOT NULL, PRIMARY KEY(source_file_id, display_name))",
    "CREATE TABLE _consultchimps_captures (capture_id VARCHAR PRIMARY KEY, source_file_id VARCHAR NOT NULL, source_key VARCHAR NOT NULL, selection_key VARCHAR NOT NULL, selection_label VARCHAR NOT NULL, reader_version VARCHAR NOT NULL, state VARCHAR NOT NULL, row_count BIGINT NOT NULL, columns_json VARCHAR NOT NULL, UNIQUE(source_file_id, selection_key, reader_version))",
    "CREATE TABLE _consultchimps_import_plans (plan_id VARCHAR NOT NULL, plan_revision BIGINT NOT NULL, baseline_revision BIGINT NOT NULL, state VARCHAR NOT NULL, recipe_json VARCHAR NOT NULL, conflicts_json VARCHAR NOT NULL, decisions_json VARCHAR NOT NULL, bindings_json VARCHAR NOT NULL, PRIMARY KEY(plan_id, plan_revision))",
    format === "duckdb"
      ? "CREATE TABLE _consultchimps_capture_rows (capture_id VARCHAR NOT NULL, source_row BIGINT NOT NULL, values_json VARCHAR NOT NULL)"
      : "CREATE TABLE _consultchimps_capture_rows (capture_id VARCHAR NOT NULL, source_row BIGINT NOT NULL, values_json VARCHAR NOT NULL, PRIMARY KEY(capture_id, source_row))",
    "CREATE TABLE _consultchimps_import_applications (import_id VARCHAR PRIMARY KEY, application_key VARCHAR NOT NULL UNIQUE, request_id VARCHAR NOT NULL, capture_id VARCHAR NOT NULL, table_name VARCHAR NOT NULL, plan_id VARCHAR NOT NULL, plan_revision BIGINT NOT NULL, row_count BIGINT NOT NULL, UNIQUE(request_id, capture_id, table_name))",
    "CREATE TABLE _consultchimps_import_requests (request_id VARCHAR PRIMARY KEY, plan_id VARCHAR NOT NULL, plan_revision BIGINT NOT NULL, import_ids_json VARCHAR NOT NULL, capture_ids_json VARCHAR NOT NULL, row_count BIGINT NOT NULL)",
    "CREATE TABLE _consultchimps_delivery_events (delivery_id VARCHAR PRIMARY KEY, request_id VARCHAR NOT NULL UNIQUE, context_json VARCHAR NOT NULL)",
    "CREATE TABLE _consultchimps_delivery_memberships (delivery_id VARCHAR NOT NULL, capture_id VARCHAR NOT NULL, PRIMARY KEY(delivery_id, capture_id))",
  ];
}

const currentPreparedDdl = [
  "CREATE TABLE _consultchimps_prepared (format_version BIGINT NOT NULL, plan_id VARCHAR PRIMARY KEY, database_id VARCHAR NOT NULL, baseline_revision BIGINT NOT NULL, schema_fingerprint VARCHAR NOT NULL, plan_revision BIGINT NOT NULL, state VARCHAR NOT NULL, recipe_json VARCHAR NOT NULL, conflicts_json VARCHAR NOT NULL, decisions_json VARCHAR NOT NULL, review_fingerprint VARCHAR NOT NULL)",
  "CREATE TABLE _consultchimps_prepared_captures (capture_id VARCHAR PRIMARY KEY, source_file_id VARCHAR, source_key VARCHAR NOT NULL, display_name VARCHAR NOT NULL, selection_key VARCHAR NOT NULL, selection_label VARCHAR NOT NULL, reader_version VARCHAR NOT NULL, content_hash VARCHAR NOT NULL, byte_count BIGINT NOT NULL, reused BIGINT NOT NULL, row_count BIGINT NOT NULL, columns_json VARCHAR NOT NULL, row_checksum VARCHAR NOT NULL)",
  "CREATE TABLE _consultchimps_prepared_bindings (source_key VARCHAR NOT NULL, selection_key VARCHAR NOT NULL, capture_id VARCHAR NOT NULL, display_name VARCHAR NOT NULL, PRIMARY KEY(source_key, selection_key))",
  "CREATE TABLE _consultchimps_prepared_rows (capture_id VARCHAR NOT NULL, source_row BIGINT NOT NULL, values_json VARCHAR NOT NULL, PRIMARY KEY(capture_id, source_row))",
] as const;

async function createEngine(
  filename: string,
  format: DatabaseFormat,
): Promise<DatabaseEngine> {
  return format === "sqlite"
    ? NodeSqliteEngine.create(filename)
    : NodeDuckDbEngine.create(filename);
}

async function executeDdl(
  engine: DatabaseEngine,
  statements: readonly string[],
): Promise<void> {
  await engine.transaction(async (transaction) => {
    for (const statement of statements) await transaction.execute(statement);
  });
}

function invalidLayout(details: Record<string, unknown>): Error {
  return Object.assign(new Error("invalid storage layout"), { details });
}

async function sqliteKeys(
  engine: DatabaseEngine,
  table: string,
): Promise<readonly string[]> {
  const rows = await engine.query(
    "SELECT indexes.name AS index_name, indexes.origin, columns.seqno, columns.name AS column_name FROM pragma_index_list(?, 'main') AS indexes JOIN pragma_index_xinfo(indexes.name, 'main') AS columns WHERE indexes.\"unique\" = 1 AND columns.key = 1 ORDER BY indexes.name, columns.seqno",
    [table],
  );
  const groups = new Map<string, { origin: string; columns: string[] }>();
  for (const row of rows) {
    const indexName = String(row["index_name"]);
    const group = groups.get(indexName) ?? {
      origin: String(row["origin"]),
      columns: [],
    };
    group.columns.push(String(row["column_name"]));
    groups.set(indexName, group);
  }
  return [...groups.values()]
    .map(({ origin, columns }) => `${origin}:${columns.join(",")}`)
    .sort();
}

async function duckdbKeys(
  engine: DatabaseEngine,
  table: string,
): Promise<readonly string[]> {
  const rows = await engine.query(
    "SELECT constraint_type, array_to_string(constraint_column_names, ',') AS column_names FROM duckdb_constraints() WHERE database_name = current_database() AND schema_name = 'main' AND table_name = ? AND constraint_type <> 'NOT NULL' ORDER BY constraint_type, column_names",
    [table],
  );
  return rows.map(
    (row) => `${String(row["constraint_type"])}:${String(row["column_names"])}`,
  );
}

async function storageSignature(
  engine: DatabaseEngine,
  format: DatabaseFormat,
  schema: InternalStorageSchema,
): Promise<readonly unknown[]> {
  const signature: unknown[] = [];
  for (const table of Object.values(schema.tables)) {
    const columns = await engine.query(
      format === "sqlite"
        ? "SELECT name, type, \"notnull\" AS required, pk FROM pragma_table_xinfo(?, 'main') ORDER BY cid"
        : 'SELECT name, type, "notnull" AS required FROM pragma_table_info(?) ORDER BY cid',
      [table.name],
    );
    signature.push({
      table: table.name,
      columns: columns.map((row: EngineRow) => ({
        name: row["name"],
        type: String(row["type"]).replaceAll(" ", "").toUpperCase(),
        required: String(row["required"]),
        ...(format === "sqlite" ? { primaryKey: String(row["pk"]) } : {}),
      })),
      keys:
        format === "sqlite"
          ? await sqliteKeys(engine, table.name)
          : await duckdbKeys(engine, table.name),
    });
  }
  return signature;
}

test.each(["sqlite", "duckdb"] as const)(
  "%s generated database storage matches the current physical layout",
  async (format) => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-storage-schema-"));
    directories.push(directory);
    const expected = await createEngine(
      path.join(directory, `expected.${format}`),
      format,
    );
    const generated = await createEngine(
      path.join(directory, `generated.${format}`),
      format,
    );
    try {
      await executeDdl(expected, currentDatabaseDdl(format));
      await generated.transaction((transaction) =>
        createInternalTables(transaction, DATABASE_STORAGE, format),
      );
      await validateInternalTables({
        query: expected.query.bind(expected),
        format,
        schema: DATABASE_STORAGE,
        invalid: invalidLayout,
      });
      await validateInternalTables({
        query: generated.query.bind(generated),
        format,
        schema: DATABASE_STORAGE,
        invalid: invalidLayout,
      });
      expect(
        await storageSignature(generated, format, DATABASE_STORAGE),
      ).toEqual(await storageSignature(expected, format, DATABASE_STORAGE));
    } finally {
      await Promise.all([expected.close(), generated.close()]);
    }
  },
);

test("generated prepared storage matches the current format 3 layout", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-prepared-storage-"));
  directories.push(directory);
  const expected = NodeSqliteEngine.create(
    path.join(directory, "expected.sqlite"),
  );
  const generated = NodeSqliteEngine.create(
    path.join(directory, "generated.sqlite"),
  );
  try {
    await executeDdl(expected, currentPreparedDdl);
    await generated.transaction((transaction) =>
      createInternalTables(transaction, PREPARED_STORAGE, "sqlite"),
    );
    await validateInternalTables({
      query: expected.query.bind(expected),
      format: "sqlite",
      schema: PREPARED_STORAGE,
      invalid: invalidLayout,
    });
    await validateInternalTables({
      query: generated.query.bind(generated),
      format: "sqlite",
      schema: PREPARED_STORAGE,
      invalid: invalidLayout,
    });
    expect(
      await storageSignature(generated, "sqlite", PREPARED_STORAGE),
    ).toEqual(await storageSignature(expected, "sqlite", PREPARED_STORAGE));
  } finally {
    await Promise.all([expected.close(), generated.close()]);
  }
});

test("named writes follow descriptor order rather than object property order", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-storage-insert-"));
  directories.push(directory);
  const engine = NodeSqliteEngine.create(path.join(directory, "insert.sqlite"));
  try {
    await engine.transaction(async (transaction) => {
      await createInternalTables(transaction, DATABASE_STORAGE, "sqlite");
      await insertInternalRow(transaction, DATABASE_STORAGE.tables.database, {
        revision: 9n,
        format_version: 1n,
        format: "sqlite",
        database_id: "DB-synthetic",
      });
    });
    await expect(
      engine.query(
        "SELECT database_id, format, format_version, revision FROM _consultchimps_database",
      ),
    ).resolves.toEqual([
      {
        database_id: "DB-synthetic",
        format: "sqlite",
        format_version: 1n,
        revision: 9n,
      },
    ]);
  } finally {
    await engine.close();
  }
});

test("named writes require every declared column and explicit null values", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-storage-row-"));
  directories.push(directory);
  const engine = NodeSqliteEngine.create(path.join(directory, "row.sqlite"));
  const completeRow = {
    capture_id: "CAP-synthetic",
    source_file_id: null,
    source_key: "source",
    display_name: "source.xlsx",
    selection_key: "Sheet1",
    selection_label: "Sheet1",
    reader_version: "reader-1",
    content_hash: "sha256:synthetic",
    byte_count: 1n,
    reused: 0n,
    row_count: 0n,
    columns_json: "[]",
    row_checksum: "checksum",
  } as const;
  try {
    await engine.transaction(async (transaction) => {
      await createInternalTables(transaction, PREPARED_STORAGE, "sqlite");
      await insertInternalRow(
        transaction,
        PREPARED_STORAGE.tables.captures,
        completeRow,
      );
      await expect(
        Reflect.apply(insertInternalRow, undefined, [
          transaction,
          PREPARED_STORAGE.tables.captures,
          {
            capture_id: "CAP-missing-null",
            source_key: "source",
            display_name: "source.xlsx",
            selection_key: "Sheet1",
            selection_label: "Sheet1",
            reader_version: "reader-1",
            content_hash: "sha256:missing",
            byte_count: 1n,
            reused: 0n,
            row_count: 0n,
            columns_json: "[]",
            row_checksum: "checksum",
          },
        ]),
      ).rejects.toThrow(/Missing: source_file_id/u);
      await expect(
        Reflect.apply(insertInternalRow, undefined, [
          transaction,
          PREPARED_STORAGE.tables.captures,
          { ...completeRow, capture_id: "CAP-extra", source_file: null },
        ]),
      ).rejects.toThrow(/Unexpected: source_file/u);
    });
    await expect(
      engine.query(
        "SELECT capture_id, source_file_id FROM _consultchimps_prepared_captures",
      ),
    ).resolves.toEqual([{ capture_id: "CAP-synthetic", source_file_id: null }]);
  } finally {
    await engine.close();
  }
});

test("generated DDL separates identifiers from storage types", () => {
  expect(
    internalCreateTableSql(DATABASE_STORAGE.tables.database, "sqlite"),
  ).toContain('"database_id" VARCHAR PRIMARY KEY');
});

test("database copy metadata retains the current bounded keyset order", () => {
  expect(
    internalCopyTables(DATABASE_STORAGE).map(({ name, columns, key }) => ({
      name,
      columns: columns.map((column) => [column.name, column.storage]),
      key,
    })),
  ).toEqual([
    {
      name: "_consultchimps_counters",
      columns: [
        ["counter_name", "text"],
        ["next_value", "integer"],
      ],
      key: ["counter_name"],
    },
    {
      name: "_consultchimps_source_contents",
      columns: [
        ["content_hash", "text"],
        ["byte_count", "integer"],
      ],
      key: ["content_hash"],
    },
    {
      name: "_consultchimps_source_files",
      columns: [
        ["source_file_id", "text"],
        ["content_hash", "text"],
        ["display_name", "text"],
      ],
      key: ["source_file_id"],
    },
    {
      name: "_consultchimps_source_names",
      columns: [
        ["source_file_id", "text"],
        ["display_name", "text"],
      ],
      key: ["source_file_id", "display_name"],
    },
    {
      name: "_consultchimps_captures",
      columns: [
        ["capture_id", "text"],
        ["source_file_id", "text"],
        ["source_key", "text"],
        ["selection_key", "text"],
        ["selection_label", "text"],
        ["reader_version", "text"],
        ["state", "text"],
        ["row_count", "integer"],
        ["columns_json", "text"],
      ],
      key: ["capture_id"],
    },
    {
      name: "_consultchimps_capture_rows",
      columns: [
        ["capture_id", "text"],
        ["source_row", "integer"],
        ["values_json", "text"],
      ],
      key: ["capture_id", "source_row"],
    },
    {
      name: "_consultchimps_import_plans",
      columns: [
        ["plan_id", "text"],
        ["plan_revision", "integer"],
        ["baseline_revision", "integer"],
        ["state", "text"],
        ["recipe_json", "text"],
        ["conflicts_json", "text"],
        ["decisions_json", "text"],
        ["bindings_json", "text"],
      ],
      key: ["plan_id", "plan_revision"],
    },
    {
      name: "_consultchimps_import_applications",
      columns: [
        ["import_id", "text"],
        ["application_key", "text"],
        ["request_id", "text"],
        ["capture_id", "text"],
        ["table_name", "text"],
        ["plan_id", "text"],
        ["plan_revision", "integer"],
        ["row_count", "integer"],
      ],
      key: ["import_id"],
    },
    {
      name: "_consultchimps_import_requests",
      columns: [
        ["request_id", "text"],
        ["plan_id", "text"],
        ["plan_revision", "integer"],
        ["import_ids_json", "text"],
        ["capture_ids_json", "text"],
        ["row_count", "integer"],
      ],
      key: ["request_id"],
    },
    {
      name: "_consultchimps_delivery_events",
      columns: [
        ["delivery_id", "text"],
        ["request_id", "text"],
        ["context_json", "text"],
      ],
      key: ["delivery_id"],
    },
    {
      name: "_consultchimps_delivery_memberships",
      columns: [
        ["delivery_id", "text"],
        ["capture_id", "text"],
      ],
      key: ["delivery_id", "capture_id"],
    },
  ]);
});
