import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { inspectDatabase, valueAsBigInt } from "../src/database.js";
import { prepareImport } from "../src/import/operations.js";
import type { ImportRecipe, ImportSource } from "../src/import/types.js";
import { createDatabase, createPreparedImport } from "../src/node.js";
import {
  PREPARED_BINDING_TABLE,
  PREPARED_CAPTURE_TABLE,
  PREPARED_ROW_TABLE,
  preparedEngineOf,
} from "../src/prepared.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

const recipe: ImportRecipe = {
  version: 1,
  routes: [
    {
      source: "submission",
      selection: "Inventory",
      destination: {
        kind: "new-table-infer",
        name: "Inventory",
        recordId: { prefix: "INV", padding: 6 },
      },
      columns: [],
    },
  ],
};

const explicitRecipe: ImportRecipe = {
  version: 1,
  routes: [
    {
      source: "submission",
      selection: "Inventory",
      destination: {
        kind: "new-table",
        schema: {
          name: "Inventory",
          columns: [{ name: "VendorID", type: "integer" }],
          recordId: { prefix: "INV", padding: 6 },
        },
      },
      columns: [{ source: "record_id", target: "VendorID", type: "integer" }],
    },
  ],
};

const reservedTargetRecipe: ImportRecipe = {
  version: 1,
  routes: [
    {
      ...recipe.routes[0]!,
      columns: [{ source: "record_id", target: "_import_id", type: "integer" }],
    },
  ],
};

function source(
  columns: readonly string[],
  counters: { batches: number; closes: number },
): ImportSource {
  const bytes = new TextEncoder().encode("synthetic reader columns");
  return {
    key: "submission",
    readerVersion: "reader-columns-1",
    bytes: {
      name: "submission.xlsx",
      size: bytes.byteLength,
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
            columns,
            async *batches() {
              counters.batches += 1;
              yield Array.from({ length: 2_001 }, (_, index) => ({
                sourceRow: index + 2,
                cells: {
                  [columns[0] ?? "Value"]: {
                    kind: "number" as const,
                    raw: "1",
                  },
                },
              }));
            },
            async close() {
              counters.closes += 1;
            },
          };
        },
      },
    ],
  };
}

async function expectEmptyStaging(
  prepared: Parameters<typeof preparedEngineOf>[0],
): Promise<void> {
  const engine = preparedEngineOf(prepared);
  for (const table of [
    PREPARED_ROW_TABLE,
    PREPARED_CAPTURE_TABLE,
    PREPARED_BINDING_TABLE,
  ]) {
    const rows = await engine.query(`SELECT COUNT(*) AS count FROM ${table}`);
    expect(valueAsBigInt(rows[0]?.["count"], "staged count")).toBe(0n);
  }
}

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format}: invalid inferred reader columns fail before capture and can be retried`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-reader-columns-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const prepared = await createPreparedImport({
      path: path.join(directory, "review.ccplan"),
      database,
      recipe,
      baselineRevision: (await inspectDatabase({ database })).revision,
    });
    const counters = { batches: 0, closes: 0 };
    try {
      await expect(
        prepareImport({
          database,
          prepared,
          recipe,
          sources: [source(["Value", "value"], counters)],
        }),
      ).rejects.toMatchObject({ code: "DB_INVALID_RECIPE" });
      expect(counters).toEqual({ batches: 0, closes: 1 });
      await expectEmptyStaging(prepared);

      await expect(
        prepareImport({
          database,
          prepared,
          recipe,
          sources: [source([""], counters)],
        }),
      ).rejects.toMatchObject({ code: "DB_INVALID_IDENTIFIER" });
      expect(counters).toEqual({ batches: 0, closes: 2 });
      await expectEmptyStaging(prepared);

      await expect(
        prepareImport({
          database,
          prepared,
          recipe,
          sources: [source(["record_id"], counters)],
        }),
      ).rejects.toMatchObject({ code: "DB_DUPLICATE_COLUMN" });
      expect(counters).toEqual({ batches: 0, closes: 3 });
      await expectEmptyStaging(prepared);

      await expect(
        prepareImport({
          database,
          prepared,
          recipe: reservedTargetRecipe,
          sources: [source(["record_id"], counters)],
        }),
      ).rejects.toMatchObject({ code: "DB_DUPLICATE_COLUMN" });
      expect(counters).toEqual({ batches: 0, closes: 4 });
      await expectEmptyStaging(prepared);

      const outcome = await prepareImport({
        database,
        prepared,
        recipe: explicitRecipe,
        sources: [source(["record_id"], counters)],
      });
      expect(outcome.prepared.state).toBe("ready");
      expect(counters).toEqual({ batches: 1, closes: 5 });
      const rows = await preparedEngineOf(prepared).query(
        `SELECT COUNT(*) AS count FROM ${PREPARED_ROW_TABLE}`,
      );
      expect(valueAsBigInt(rows[0]?.["count"], "staged count")).toBe(2_001n);
    } finally {
      await prepared.close();
      await database.close();
    }
  });
}
