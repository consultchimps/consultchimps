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
import { PREPARED_CAPTURE_TABLE, preparedEngineOf } from "../src/prepared.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

function recipe(table: string): ImportRecipe {
  return {
    version: 1,
    routes: [
      {
        source: "submission",
        selection: "Inventory",
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

function source(): ImportSource {
  const bytes = new TextEncoder().encode("synthetic inspection snapshot");
  return {
    key: "submission",
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
                    Value: { kind: "string" as const, value: "North" },
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

function deferred(): {
  readonly promise: Promise<void>;
  readonly resolve: () => void;
} {
  let resolve!: () => void;
  const promise = new Promise<void>((done) => {
    resolve = done;
  });
  return { promise, resolve };
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
    const originalQuery = engine.query.bind(engine);
    const captureCountReached = deferred();
    const releaseCaptureCount = deferred();
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

      engine.query = async (sql, values) => {
        if (
          sql.includes(`sum(row_count)`) &&
          sql.includes(PREPARED_CAPTURE_TABLE)
        ) {
          captureCountReached.resolve();
          await releaseCaptureCount.promise;
        }
        return originalQuery(sql, values);
      };

      const pendingInspection = inspectImport({
        database,
        prepared,
        page: { limit: 10 },
      });
      await captureCountReached.promise;
      const revised = await replaceImportRecipe({
        database,
        prepared,
        recipe: recipe("RevisedInventory"),
      });
      expect(revised.state).toBe("ready");
      releaseCaptureCount.resolve();

      const inspection = await pendingInspection;
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
      engine.query = originalQuery;
      releaseCaptureCount.resolve();
      await prepared.close();
      await database.close();
    }
  });
}
