import { copyFile, mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterAll, afterEach, beforeAll, expect, test } from "vitest";

import { engineOf, inspectDatabase } from "../src/database.js";
import { NodeDuckDbEngine } from "../src/engines/duckdb/node.js";
import { NodeSqliteEngine } from "../src/engines/sqlite/node.js";
import type { DatabaseEngine } from "../src/internal/engine.js";
import {
  APPLICATION_TABLE,
  CAPTURE_ROW_TABLE,
  SOURCE_CONTENT_TABLE,
  SOURCE_FILE_TABLE,
  SOURCE_NAME_TABLE,
} from "../src/metadata.js";
import {
  createDatabase,
  createPreparedImport,
  openDatabase,
  openPreparedImport,
} from "../src/node.js";
import {
  PREPARED_BINDING_TABLE,
  PREPARED_CAPTURE_TABLE,
  PREPARED_METADATA_TABLE,
  PREPARED_ROW_TABLE,
} from "../src/prepared.js";
import type { DatabaseFormat } from "../src/schema.js";

interface ConstraintMutation {
  readonly name: string;
  create(format: DatabaseFormat): string;
  readonly table: string;
}

const mutations: readonly ConstraintMutation[] = [
  {
    name: "missing simple primary key",
    table: SOURCE_CONTENT_TABLE,
    create: () =>
      `CREATE TABLE ${SOURCE_CONTENT_TABLE} (content_hash VARCHAR, byte_count BIGINT NOT NULL)`,
  },
  {
    name: "missing composite primary key",
    table: SOURCE_NAME_TABLE,
    create: () =>
      `CREATE TABLE ${SOURCE_NAME_TABLE} (source_file_id VARCHAR NOT NULL, display_name VARCHAR NOT NULL)`,
  },
  {
    name: "missing secondary unique constraint",
    table: SOURCE_FILE_TABLE,
    create: () =>
      `CREATE TABLE ${SOURCE_FILE_TABLE} (source_file_id VARCHAR PRIMARY KEY, content_hash VARCHAR NOT NULL, display_name VARCHAR NOT NULL)`,
  },
  {
    name: "changed column type",
    table: CAPTURE_ROW_TABLE,
    create: (format) =>
      `CREATE TABLE ${CAPTURE_ROW_TABLE} (capture_id VARCHAR NOT NULL, source_row BIGINT NOT NULL, values_json BIGINT NOT NULL${format === "sqlite" ? ", PRIMARY KEY(capture_id, source_row)" : ""})`,
  },
  {
    name: "changed column nullability",
    table: CAPTURE_ROW_TABLE,
    create: (format) =>
      `CREATE TABLE ${CAPTURE_ROW_TABLE} (capture_id VARCHAR NOT NULL, source_row BIGINT NOT NULL, values_json VARCHAR${format === "sqlite" ? ", PRIMARY KEY(capture_id, source_row)" : ""})`,
  },
  {
    name: "unsupported extra unique constraint",
    table: SOURCE_NAME_TABLE,
    create: () =>
      `CREATE TABLE ${SOURCE_NAME_TABLE} (source_file_id VARCHAR NOT NULL, display_name VARCHAR NOT NULL, PRIMARY KEY(source_file_id, display_name), UNIQUE(display_name))`,
  },
  {
    name: "missing import application uniqueness",
    table: APPLICATION_TABLE,
    create: () =>
      `CREATE TABLE ${APPLICATION_TABLE} (import_id VARCHAR PRIMARY KEY, application_key VARCHAR NOT NULL UNIQUE, request_id VARCHAR NOT NULL, capture_id VARCHAR NOT NULL, table_name VARCHAR NOT NULL, plan_id VARCHAR NOT NULL, plan_revision BIGINT NOT NULL, row_count BIGINT NOT NULL)`,
  },
];

const directories: string[] = [];
let templateDirectory: string;
let templates: Record<DatabaseFormat, string>;
let preparedTemplate: string;

async function openEngine(
  filePath: string,
  format: DatabaseFormat,
): Promise<DatabaseEngine> {
  return format === "sqlite"
    ? NodeSqliteEngine.open(filePath)
    : NodeDuckDbEngine.create(filePath);
}

beforeAll(async () => {
  templateDirectory = await mkdtemp(
    path.join(tmpdir(), "cc-internal-constraints-template-"),
  );
  templates = {} as Record<DatabaseFormat, string>;
  for (const format of ["sqlite", "duckdb"] as const) {
    const filePath = path.join(templateDirectory, `template.${format}`);
    const { database } = await createDatabase({ path: filePath, format });
    await engineOf(database).execute("CREATE TABLE sentinel (value VARCHAR)");
    await engineOf(database).execute("INSERT INTO sentinel VALUES (?)", [
      "preserved",
    ]);
    await database.checkpoint();
    await database.close();
    templates[format] = filePath;
  }
  const database = await openDatabase({ path: templates.sqlite });
  preparedTemplate = path.join(templateDirectory, "template.ccplan");
  const prepared = await createPreparedImport({
    path: preparedTemplate,
    database,
    recipe: { version: 1, routes: [] },
    baselineRevision: (await inspectDatabase({ database })).revision,
  });
  await prepared.close();
  await database.close();
});

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

afterAll(async () => {
  await rm(templateDirectory, { recursive: true, force: true });
});

for (const format of ["sqlite", "duckdb"] as const) {
  for (const mutation of mutations) {
    test(`${format}: rejects ${mutation.name} on ${mutation.table} without changing the file`, async () => {
      const directory = await mkdtemp(
        path.join(tmpdir(), "cc-internal-constraints-"),
      );
      directories.push(directory);
      const filePath = path.join(directory, `damaged.${format}`);
      await copyFile(templates[format], filePath);
      const engine = await openEngine(filePath, format);
      try {
        await engine.transaction(async (transaction) => {
          await transaction.execute(`DROP TABLE ${mutation.table}`);
          await transaction.execute(mutation.create(format));
        });
        expect(await engine.query("SELECT value FROM sentinel")).toEqual([
          { value: "preserved" },
        ]);
      } finally {
        await engine.close();
      }
      const before = await readFile(filePath);

      const failure = await openDatabase({
        path: filePath,
        readonly: true,
      }).then(
        async (database) => {
          await database.close();
          return undefined;
        },
        (error: unknown) => error,
      );
      expect(failure).toMatchObject({
        code: "DB_CORRUPT_DATABASE",
        details: { table: mutation.table },
      });
      expect((await readFile(filePath)).equals(before)).toBe(true);
    });
  }

  test(`${format}: accepts an ordinary index on an internal table`, async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "cc-internal-index-valid-"),
    );
    directories.push(directory);
    const filePath = path.join(directory, `indexed.${format}`);
    await copyFile(templates[format], filePath);
    const engine = await openEngine(filePath, format);
    try {
      await engine.execute(
        `CREATE INDEX capture_rows_source_row ON ${CAPTURE_ROW_TABLE} (source_row)`,
      );
    } finally {
      await engine.close();
    }
    const before = await readFile(filePath);

    const database = await openDatabase({ path: filePath, readonly: true });
    await database.close();

    expect((await readFile(filePath)).equals(before)).toBe(true);
  });

  test(`${format}: rejects an explicit unique index on an internal table`, async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "cc-internal-index-invalid-"),
    );
    directories.push(directory);
    const filePath = path.join(directory, `indexed.${format}`);
    await copyFile(templates[format], filePath);
    const engine = await openEngine(filePath, format);
    try {
      await engine.execute(
        `CREATE UNIQUE INDEX capture_rows_source_row_unique ON ${CAPTURE_ROW_TABLE} (source_row)`,
      );
    } finally {
      await engine.close();
    }
    const before = await readFile(filePath);

    const failure = await openDatabase({
      path: filePath,
      readonly: true,
    }).then(
      async (database) => {
        await database.close();
        return undefined;
      },
      (error: unknown) => error,
    );
    expect(failure).toMatchObject({
      code: "DB_CORRUPT_DATABASE",
      details: { table: CAPTURE_ROW_TABLE },
    });
    expect((await readFile(filePath)).equals(before)).toBe(true);
  });
}

const preparedMutations = [
  {
    name: "missing prepared-row composite primary key",
    table: PREPARED_ROW_TABLE,
    statements: [
      `DROP TABLE ${PREPARED_ROW_TABLE}`,
      `CREATE TABLE ${PREPARED_ROW_TABLE} (capture_id VARCHAR NOT NULL, source_row BIGINT NOT NULL, values_json VARCHAR NOT NULL)`,
    ],
  },
  {
    name: "nullable required prepared-capture column",
    table: PREPARED_CAPTURE_TABLE,
    statements: [
      `DROP TABLE ${PREPARED_CAPTURE_TABLE}`,
      `CREATE TABLE ${PREPARED_CAPTURE_TABLE} (capture_id VARCHAR PRIMARY KEY, source_file_id VARCHAR, source_key VARCHAR, display_name VARCHAR NOT NULL, selection_key VARCHAR NOT NULL, selection_label VARCHAR NOT NULL, reader_version VARCHAR NOT NULL, content_hash VARCHAR NOT NULL, byte_count BIGINT NOT NULL, reused BIGINT NOT NULL, row_count BIGINT NOT NULL, columns_json VARCHAR NOT NULL, row_checksum VARCHAR NOT NULL)`,
    ],
  },
  {
    name: "changed prepared-metadata column type",
    table: PREPARED_METADATA_TABLE,
    statements: [
      `ALTER TABLE ${PREPARED_METADATA_TABLE} RENAME TO previous_prepared_metadata`,
      `CREATE TABLE ${PREPARED_METADATA_TABLE} (format_version VARCHAR NOT NULL, plan_id VARCHAR PRIMARY KEY, database_id VARCHAR NOT NULL, baseline_revision BIGINT NOT NULL, schema_fingerprint VARCHAR NOT NULL, plan_revision BIGINT NOT NULL, state VARCHAR NOT NULL, recipe_json VARCHAR NOT NULL, conflicts_json VARCHAR NOT NULL, decisions_json VARCHAR NOT NULL, review_fingerprint VARCHAR NOT NULL)`,
      `INSERT INTO ${PREPARED_METADATA_TABLE} SELECT * FROM previous_prepared_metadata`,
      "DROP TABLE previous_prepared_metadata",
    ],
  },
  {
    name: "extra prepared-binding unique constraint",
    table: PREPARED_BINDING_TABLE,
    statements: [
      `DROP TABLE ${PREPARED_BINDING_TABLE}`,
      `CREATE TABLE ${PREPARED_BINDING_TABLE} (source_key VARCHAR NOT NULL, selection_key VARCHAR NOT NULL, capture_id VARCHAR NOT NULL, display_name VARCHAR NOT NULL, PRIMARY KEY(source_key, selection_key), UNIQUE(capture_id))`,
    ],
  },
] as const;

for (const readonly of [false, true]) {
  for (const mutation of preparedMutations) {
    test(`prepared readonly=${String(readonly)} rejects ${mutation.name}`, async () => {
      const directory = await mkdtemp(
        path.join(tmpdir(), "cc-prepared-constraints-"),
      );
      directories.push(directory);
      const filePath = path.join(directory, "damaged.ccplan");
      await copyFile(preparedTemplate, filePath);
      const engine = NodeSqliteEngine.open(filePath);
      try {
        await engine.transaction(async (transaction) => {
          for (const statement of mutation.statements) {
            await transaction.execute(statement);
          }
        });
      } finally {
        await engine.close();
      }
      const before = await readFile(filePath);

      const failure = await openPreparedImport({
        path: filePath,
        readonly,
      }).then(
        async (prepared) => {
          await prepared.close();
          return undefined;
        },
        (error: unknown) => error,
      );
      expect(failure).toMatchObject({
        code: "DB_INVALID_PREPARED_IMPORT",
        details: { table: mutation.table },
      });
      expect((await readFile(filePath)).equals(before)).toBe(true);
    });
  }
}

test("prepared captures keep nullable source identity in valid plans", async () => {
  const directory = await mkdtemp(
    path.join(tmpdir(), "cc-prepared-constraints-valid-"),
  );
  directories.push(directory);
  const filePath = path.join(directory, "valid.ccplan");
  await copyFile(preparedTemplate, filePath);
  const engine = NodeSqliteEngine.open(filePath, true);
  try {
    const columns = await engine.query(
      `SELECT name, "notnull" AS not_null FROM pragma_table_info(?) WHERE name = 'source_file_id'`,
      [PREPARED_CAPTURE_TABLE],
    );
    expect(columns).toEqual([{ name: "source_file_id", not_null: 0n }]);
  } finally {
    await engine.close();
  }
  const writable = await openPreparedImport({ path: filePath });
  await writable.close();
  const readonly = await openPreparedImport({ path: filePath, readonly: true });
  await readonly.close();
});
