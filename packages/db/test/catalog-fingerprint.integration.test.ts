import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { engineOf, inspectDatabase } from "../src/database.js";
import { applyImport, prepareImport } from "../src/import/operations.js";
import type { ImportRecipe, ImportSource } from "../src/import/types.js";
import { createDatabase, createPreparedImport } from "../src/node.js";
import {
  applySchema,
  planSchema,
  readSchemaFingerprint,
} from "../src/records.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

function plannedSchema() {
  return {
    version: 1 as const,
    tables: [
      {
        name: "planned_items",
        recordId: { prefix: "ITEM", padding: 4 },
        columns: [{ name: "value", type: "text" as const }],
      },
    ],
  };
}

const recipe: ImportRecipe = {
  version: 1,
  routes: [
    {
      source: "synthetic",
      selection: "Items",
      destination: { kind: "new-table", schema: plannedSchema().tables[0]! },
      columns: [{ source: "value", target: "value", type: "text" }],
    },
  ],
};

function source(): ImportSource {
  const bytes = new TextEncoder().encode("synthetic catalog fingerprint");
  return {
    key: "synthetic",
    readerVersion: "catalog-fingerprint-test-1",
    bytes: {
      name: "synthetic.txt",
      size: bytes.byteLength,
      async readAt(offset, length) {
        return bytes.slice(offset, offset + length);
      },
    },
    selections: [
      {
        key: "Items",
        label: "Items",
        async open() {
          return {
            columns: ["value"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: {
                    value: { kind: "string" as const, value: "North" },
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

test("DuckDB views deterministically invalidate a schema plan before writes", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-catalog-schema-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "workspace.duckdb"),
    format: "duckdb",
  });
  try {
    const engine = engineOf(database);
    const before = await engine.readTransaction((transaction) =>
      readSchemaFingerprint(transaction, "duckdb"),
    );
    const repeated = await engine.readTransaction((transaction) =>
      readSchemaFingerprint(transaction, "duckdb"),
    );
    expect(repeated).toBe(before);

    const plan = await planSchema({ database, schema: plannedSchema() });
    await engine.execute(
      "CREATE VIEW planned_items AS SELECT 'existing' AS value",
    );
    const withView = await engine.readTransaction((transaction) =>
      readSchemaFingerprint(transaction, "duckdb"),
    );
    expect(withView).not.toBe(before);

    await expect(applySchema({ database, plan })).rejects.toMatchObject({
      code: "DB_STALE_SCHEMA_PLAN",
    });
    expect(await engine.query("SELECT value FROM planned_items")).toEqual([
      { value: "existing" },
    ]);
    await expect(inspectDatabase({ database })).resolves.toMatchObject({
      revision: 0n,
      tables: [],
    });

    await engine.execute("DROP VIEW planned_items");
    const restored = await engine.readTransaction((transaction) =>
      readSchemaFingerprint(transaction, "duckdb"),
    );
    expect(restored).toBe(before);
  } finally {
    await database.close();
  }
});

test("DuckDB views invalidate an approved import before capture application", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-catalog-import-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "workspace.duckdb"),
    format: "duckdb",
  });
  const prepared = await createPreparedImport({
    path: path.join(directory, "review.ccplan"),
    database,
    recipe,
    baselineRevision: 0n,
  });
  try {
    const outcome = await prepareImport({
      database,
      prepared,
      recipe,
      sources: [source()],
    });
    expect(outcome.prepared.state).toBe("ready");
    if (outcome.prepared.state !== "ready")
      throw new Error("The synthetic import review is not ready");

    const engine = engineOf(database);
    await engine.execute(
      "CREATE VIEW planned_items AS SELECT 'existing' AS value",
    );
    await expect(
      applyImport({
        database,
        prepared,
        approved: outcome.prepared,
        requestId: "catalog-view-stale-import",
      }),
    ).rejects.toMatchObject({ code: "DB_STALE_IMPORT_PLAN" });

    expect(await engine.query("SELECT value FROM planned_items")).toEqual([
      { value: "existing" },
    ]);
    await expect(inspectDatabase({ database })).resolves.toMatchObject({
      revision: 0n,
      tables: [],
      captures: 0n,
      completedImports: 0n,
      deliveries: 0n,
    });
  } finally {
    await prepared.close();
    await database.close();
  }
});
