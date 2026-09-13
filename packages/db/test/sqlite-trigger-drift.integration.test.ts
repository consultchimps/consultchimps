import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import SqliteDatabase from "better-sqlite3";
import { afterEach, expect, test } from "vitest";

import { engineOf, inspectDatabase, type Database } from "../src/database.js";
import {
  applyImport,
  prepareImport,
  resolveImport,
} from "../src/import/operations.js";
import { recordDelivery } from "../src/import/deliveries.js";
import type { ImportRecipe, ImportSource } from "../src/import/types.js";
import {
  APPLICATION_TABLE,
  DATABASE_METADATA_TABLE,
  DELIVERY_TABLE,
  IMPORT_REQUEST_TABLE,
} from "../src/metadata.js";
import {
  createDatabase,
  createPreparedImport,
  openDatabase,
} from "../src/node.js";
import {
  PREPARED_CAPTURE_TABLE,
  PREPARED_METADATA_TABLE,
  PREPARED_ROW_TABLE,
  preparedRef,
  preparedEngineOf,
} from "../src/prepared.js";
import { applySchema, planSchema } from "../src/records.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

function source(): ImportSource {
  const bytes = new TextEncoder().encode("synthetic trigger test");
  return {
    key: "submission",
    readerVersion: "synthetic-1",
    bytes: {
      name: "submission.xlsx",
      size: bytes.length,
      async readAt(offset, length) {
        return bytes.slice(offset, offset + length);
      },
    },
    selections: [
      {
        key: "Inventory",
        label: "Inventory",
        async open() {
          return {
            columns: ["Value"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: {
                    Value: { kind: "string" as const, value: "Synthetic" },
                  },
                },
              ];
            },
            async close() {},
          };
        },
      },
    ],
  };
}

const inventorySchema = {
  name: "Inventory",
  recordId: { prefix: "INV", padding: 6 },
  columns: [{ name: "Value", type: "text" as const }],
  foreignKeys: [],
};

const recipe: ImportRecipe = {
  version: 1,
  routes: [
    {
      source: "submission",
      selection: "Inventory",
      destination: {
        kind: "new-table",
        schema: inventorySchema,
      },
      columns: [{ source: "Value", target: "Value", type: "text" }],
    },
  ],
};

async function readyImport(
  directory: string,
  database: Database,
  selectedRecipe = recipe,
) {
  const prepared = await createPreparedImport({
    path: path.join(directory, `${crypto.randomUUID()}.ccplan`),
    database,
    recipe: selectedRecipe,
    baselineRevision: (await inspectDatabase({ database })).revision,
  });
  await prepareImport({
    database,
    prepared,
    recipe: selectedRecipe,
    sources: [source()],
  });
  const approved = await resolveImport({ database, prepared, decisions: [] });
  if (approved.state !== "ready") throw new Error("Expected a ready import");
  return { prepared, approved };
}

async function emptyWriteState(database: Database) {
  const inspection = await inspectDatabase({ database });
  const receipts = await engineOf(database).query(
    `SELECT count(*) AS count FROM ${IMPORT_REQUEST_TABLE}`,
  );
  const applications = await engineOf(database).query(
    `SELECT count(*) AS count FROM ${APPLICATION_TABLE}`,
  );
  return {
    revision: inspection.revision,
    rowCount: inspection.tables[0]?.rowCount ?? 0n,
    receipts: receipts[0]?.["count"],
    applications: applications[0]?.["count"],
    deliveries: inspection.deliveries,
  };
}

test("schema application rejects a trigger before changing schema or revision", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-trigger-schema-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "workspace.sqlite"),
    format: "sqlite",
  });
  const plan = await planSchema({
    database,
    schema: { version: 1, tables: [inventorySchema] },
  });
  const before = await emptyWriteState(database);
  try {
    await engineOf(database).execute(
      `CREATE TRIGGER suppress_revision BEFORE UPDATE ON ${DATABASE_METADATA_TABLE} BEGIN SELECT RAISE(IGNORE); END`,
    );
    await expect(applySchema({ database, plan })).rejects.toMatchObject({
      code: "DB_SCHEMA_DRIFT",
      details: {
        table: DATABASE_METADATA_TABLE,
        trigger: "suppress_revision",
      },
    });
    expect(await emptyWriteState(database)).toEqual(before);
  } finally {
    await database.close();
  }
});

test("a persistent trigger on a managed table is rejected before import writes", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-trigger-managed-"));
  directories.push(directory);
  const databasePath = path.join(directory, "workspace.sqlite");
  const created = await createDatabase({
    path: databasePath,
    format: "sqlite",
    schema: { version: 1, tables: [inventorySchema] },
  });
  await created.database.close();
  const raw = new SqliteDatabase(databasePath);
  raw.exec(
    'CREATE TRIGGER suppress_inventory BEFORE INSERT ON "Inventory" BEGIN SELECT RAISE(IGNORE); END',
  );
  raw.close();

  const database = await openDatabase({ path: databasePath });
  const existingRecipe: ImportRecipe = {
    ...recipe,
    routes: [
      {
        ...recipe.routes[0]!,
        destination: { kind: "existing-table", table: "Inventory" },
      },
    ],
  };
  const { prepared, approved } = await readyImport(
    directory,
    database,
    existingRecipe,
  );
  const before = await emptyWriteState(database);
  try {
    await expect(
      applyImport({ database, prepared, approved, requestId: "triggered" }),
    ).rejects.toMatchObject({
      code: "DB_SCHEMA_DRIFT",
      details: { table: "Inventory", trigger: "suppress_inventory" },
    });
    expect(await emptyWriteState(database)).toEqual(before);
  } finally {
    await prepared.close();
    await database.close();
  }
});

test("a trigger added to the receipt table after open cannot create a false receipt", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-trigger-receipt-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "workspace.sqlite"),
    format: "sqlite",
  });
  const { prepared, approved } = await readyImport(directory, database);
  const before = await emptyWriteState(database);
  try {
    await engineOf(database).execute(
      `CREATE TRIGGER suppress_receipt BEFORE INSERT ON ${IMPORT_REQUEST_TABLE} BEGIN SELECT RAISE(IGNORE); END`,
    );
    await expect(
      applyImport({ database, prepared, approved, requestId: "receipt" }),
    ).rejects.toMatchObject({
      code: "DB_SCHEMA_DRIFT",
      details: { table: IMPORT_REQUEST_TABLE, trigger: "suppress_receipt" },
    });
    expect(await emptyWriteState(database)).toEqual(before);
  } finally {
    await prepared.close();
    await database.close();
  }
});

test("TEMP triggers on managed tables are rejected while unmanaged triggers remain allowed", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-trigger-temp-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "workspace.sqlite"),
    format: "sqlite",
  });
  await engineOf(database).execute("CREATE TABLE scratch (value VARCHAR)");
  await engineOf(database).execute(
    "CREATE TRIGGER scratch_trigger BEFORE INSERT ON scratch BEGIN SELECT RAISE(IGNORE); END",
  );
  const first = await readyImport(directory, database);
  try {
    const applied = await applyImport({
      database,
      prepared: first.prepared,
      approved: first.approved,
      requestId: "unmanaged-trigger",
    });
    expect(applied.metrics.rowsImported).toBe(1);
    const beforeDelivery = await emptyWriteState(database);
    await engineOf(database).execute(
      `CREATE TRIGGER suppress_delivery BEFORE INSERT ON ${DELIVERY_TABLE} BEGIN SELECT RAISE(IGNORE); END`,
    );
    await expect(
      recordDelivery({
        database,
        captureIds: applied.captureIds,
        requestId: "standalone-delivery",
        context: {
          label: "Synthetic delivery",
          scope: { kind: "full" },
          attributes: {},
        },
      }),
    ).rejects.toMatchObject({
      code: "DB_SCHEMA_DRIFT",
      details: { table: DELIVERY_TABLE, trigger: "suppress_delivery" },
    });
    expect(await emptyWriteState(database)).toEqual(beforeDelivery);
  } finally {
    await first.prepared.close();
  }
  await engineOf(database).execute("DROP TRIGGER suppress_delivery");
  const existingRecipe: ImportRecipe = {
    ...recipe,
    routes: [
      {
        ...recipe.routes[0]!,
        destination: { kind: "existing-table", table: "Inventory" },
      },
    ],
  };
  const second = await readyImport(directory, database, existingRecipe);
  try {
    await engineOf(database).execute(
      'CREATE TEMP TRIGGER temp_inventory BEFORE INSERT ON "Inventory" BEGIN SELECT RAISE(IGNORE); END',
    );
    await expect(
      applyImport({
        database,
        prepared: second.prepared,
        approved: second.approved,
        requestId: "temp-trigger",
      }),
    ).rejects.toMatchObject({
      code: "DB_SCHEMA_DRIFT",
      details: { table: "Inventory", trigger: "temp_inventory" },
    });
  } finally {
    await second.prepared.close();
    await database.close();
  }
});

test("prepared-row triggers are rejected before captured rows are staged", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-trigger-plan-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "workspace.sqlite"),
    format: "sqlite",
  });
  const prepared = await createPreparedImport({
    path: path.join(directory, "review.ccplan"),
    database,
    recipe,
    baselineRevision: (await inspectDatabase({ database })).revision,
  });
  try {
    await preparedEngineOf(prepared).execute(
      `CREATE TRIGGER suppress_prepared_rows BEFORE INSERT ON ${PREPARED_ROW_TABLE} BEGIN SELECT RAISE(IGNORE); END`,
    );
    await expect(
      prepareImport({ database, prepared, recipe, sources: [source()] }),
    ).rejects.toMatchObject({
      code: "DB_SCHEMA_DRIFT",
      details: {
        table: PREPARED_ROW_TABLE,
        trigger: "suppress_prepared_rows",
      },
    });
    expect(
      await preparedEngineOf(prepared).query(
        `SELECT count(*) AS count FROM ${PREPARED_CAPTURE_TABLE}`,
      ),
    ).toEqual([{ count: 0n }]);
  } finally {
    await prepared.close();
    await database.close();
  }
});

test("prepared metadata triggers cannot publish a false ready review", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-trigger-review-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "workspace.sqlite"),
    format: "sqlite",
  });
  const emptyRecipe: ImportRecipe = { version: 1, routes: [] };
  const prepared = await createPreparedImport({
    path: path.join(directory, "review.ccplan"),
    database,
    recipe: emptyRecipe,
    baselineRevision: (await inspectDatabase({ database })).revision,
  });
  try {
    await preparedEngineOf(prepared).execute(
      `CREATE TRIGGER suppress_review BEFORE UPDATE ON ${PREPARED_METADATA_TABLE} BEGIN SELECT RAISE(IGNORE); END`,
    );
    await expect(
      prepareImport({
        database,
        prepared,
        recipe: emptyRecipe,
        sources: [],
      }),
    ).rejects.toMatchObject({
      code: "DB_SCHEMA_DRIFT",
      details: { table: PREPARED_METADATA_TABLE, trigger: "suppress_review" },
    });
    expect(await preparedRef(prepared)).toMatchObject({
      state: "needs-review",
      planRevision: 1n,
    });
  } finally {
    await prepared.close();
    await database.close();
  }
});
