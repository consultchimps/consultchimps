import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import {
  applyImport,
  inspectImport,
  prepareImport,
} from "../src/import/operations.js";
import { replaceImportRecipe } from "../src/import/resolve.js";
import type { ImportRecipe, ImportSource } from "../src/import/types.js";
import { createDatabase, createPreparedImport } from "../src/node.js";
import { inspectDatabase } from "../src/database.js";
import { preparedEngineOf } from "../src/prepared.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

function recipe(
  table: string,
  sourceKey = "submission",
  selectionKey = "Inventory",
): ImportRecipe {
  return {
    version: 1,
    routes: [
      {
        source: sourceKey,
        selection: selectionKey,
        destination: {
          kind: "new-table",
          schema: {
            name: table,
            columns: [{ name: "Value", type: "text" }],
            recordId: {
              prefix: table === "Inventory" ? "INV" : "REV",
              padding: 6,
            },
          },
        },
        columns: [{ source: "Value", target: "Value", type: "text" }],
      },
    ],
  };
}

function source(
  sourceKey = "submission",
  selectionKey = "Inventory",
  value = "North",
): ImportSource {
  const bytes = new TextEncoder().encode(
    `synthetic inspection snapshot ${sourceKey} ${selectionKey}`,
  );
  return {
    key: sourceKey,
    readerVersion: "synthetic-snapshot-1",
    bytes: {
      name: "submission.synthetic",
      size: bytes.byteLength,
      async readAt(offset, length) {
        return bytes.slice(offset, offset + length);
      },
    },
    selections: [
      {
        key: selectionKey,
        label: selectionKey,
        async open() {
          return {
            columns: ["Value"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: {
                    Value: { kind: "string" as const, value },
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

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format}: inspection returns the review reference for the recipe it displayed`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-review-snapshot-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const initialRecipe = recipe("Inventory");
    const prepared = await createPreparedImport({
      path: path.join(directory, "review.ccplan"),
      database,
      recipe: initialRecipe,
      baselineRevision: 0n,
    });
    const engine = preparedEngineOf(prepared);
    const originalReadTransaction = engine.readTransaction.bind(engine);
    try {
      const initial = await prepareImport({
        database,
        prepared,
        recipe: initialRecipe,
        sources: [source()],
      });
      expect(initial.prepared.state).toBe("ready");
      if (initial.prepared.state !== "ready") {
        throw new Error("The synthetic import plan did not become ready.");
      }

      let revised = false;
      engine.readTransaction = async (read) => {
        const snapshot = await originalReadTransaction(read);
        if (!revised) {
          revised = true;
          engine.readTransaction = originalReadTransaction;
          const replacement = await replaceImportRecipe({
            database,
            prepared,
            recipe: recipe("RevisedInventory"),
          });
          expect(replacement.state).toBe("ready");
        }
        return snapshot;
      };

      const inspection = await inspectImport({
        database,
        prepared,
        page: { limit: 10 },
      });
      expect(inspection.routes[0]?.destination).toMatchObject({
        kind: "new-table",
        schema: { name: "Inventory" },
      });
      expect(inspection.prepared).toEqual(initial.prepared);
      expect(inspection.prepared.state).toBe("ready");
      if (inspection.prepared.state !== "ready") {
        throw new Error("The inspected import plan was not ready.");
      }
      await expect(
        applyImport({
          database,
          prepared,
          approved: inspection.prepared,
          requestId: `snapshot-${format}`,
        }),
      ).rejects.toMatchObject({ code: "DB_STALE_IMPORT_PLAN" });
    } finally {
      engine.readTransaction = originalReadTransaction;
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: apply consumes only captures from its approved review snapshot`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-apply-snapshot-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const initialRecipe = recipe("Inventory");
    const prepared = await createPreparedImport({
      path: path.join(directory, "review.ccplan"),
      database,
      recipe: initialRecipe,
      baselineRevision: 0n,
    });
    const engine = preparedEngineOf(prepared);
    const originalReadTransaction = engine.readTransaction.bind(engine);
    try {
      const initial = await prepareImport({
        database,
        prepared,
        recipe: initialRecipe,
        sources: [source()],
      });
      expect(initial.prepared.state).toBe("ready");
      if (initial.prepared.state !== "ready") {
        throw new Error("The synthetic import plan did not become ready.");
      }
      const addedRecipe = recipe("Added", "added", "Added");
      const expandedRecipe: ImportRecipe = {
        version: 1,
        routes: [...initialRecipe.routes, ...addedRecipe.routes],
      };
      let preparedConcurrently = false;
      engine.readTransaction = async (read) => {
        const snapshot = await originalReadTransaction(read);
        if (!preparedConcurrently) {
          preparedConcurrently = true;
          engine.readTransaction = originalReadTransaction;
          const expanded = await prepareImport({
            database,
            prepared,
            recipe: expandedRecipe,
            sources: [source("added", "Added", "South")],
          });
          expect(expanded.prepared.state).toBe("ready");
        }
        return snapshot;
      };

      await expect(
        applyImport({
          database,
          prepared,
          approved: initial.prepared,
          requestId: `apply-snapshot-${format}`,
        }),
      ).resolves.toMatchObject({
        metrics: { rowsImported: 1, tablesCreated: 1 },
      });
      await expect(inspectDatabase({ database })).resolves.toMatchObject({
        captures: 1n,
        completedImports: 1n,
        tables: [{ name: "Inventory", rowCount: 1n }],
      });
    } finally {
      engine.readTransaction = originalReadTransaction;
      await prepared.close();
      await database.close();
    }
  });
}
