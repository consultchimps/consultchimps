import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import {
  inspectImport,
  prepareImport,
  resolveImport,
} from "../src/import/operations.js";
import type { ImportRecipe, ImportSource } from "../src/import/types.js";
import type { DatabaseEngine } from "../src/internal/engine.js";
import { createDatabase, createPreparedImport } from "../src/node.js";
import { preparedEngineOf } from "../src/prepared.js";

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
      source: "original",
      selection: "Inventory",
      destination: {
        kind: "new-table",
        schema: {
          name: "Inventory",
          columns: [{ name: "Value", type: "text" }],
          recordId: { prefix: "INV", padding: 6 },
        },
      },
      columns: [{ source: "Value", target: "Value", type: "text" }],
    },
  ],
};

function source(options: {
  readonly key: string;
  readonly selection: string;
  readonly value: string;
}): ImportSource {
  const bytes = new TextEncoder().encode(
    `synthetic ${options.key} ${options.selection}`,
  );
  return {
    key: options.key,
    readerVersion: "synthetic-review-snapshot-1",
    bytes: {
      name: `${options.key}.synthetic`,
      size: bytes.byteLength,
      async readAt(offset, length) {
        return bytes.slice(offset, offset + length);
      },
    },
    selections: [
      {
        key: options.selection,
        label: options.selection,
        async open() {
          return {
            columns: ["Value"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: {
                    Value: { kind: "string" as const, value: options.value },
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
  test(`${format}: stale review resolution cannot approve a concurrently added capture`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-review-write-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const prepared = await createPreparedImport({
      path: path.join(directory, "review.ccplan"),
      database,
      recipe,
      baselineRevision: 0n,
    });
    const engine = preparedEngineOf(prepared);
    const originalTransaction = engine.transaction.bind(engine);
    const staleWriteReached = deferred();
    const releaseStaleWrite = deferred();
    try {
      const initial = await prepareImport({
        database,
        prepared,
        recipe,
        sources: [
          source({ key: "original", selection: "Inventory", value: "North" }),
        ],
      });
      expect(initial.prepared.state).toBe("ready");

      let interceptNextTransaction = true;
      engine.transaction = async <T>(
        work: Parameters<DatabaseEngine["transaction"]>[0],
      ): Promise<T> => {
        if (interceptNextTransaction) {
          interceptNextTransaction = false;
          staleWriteReached.resolve();
          await releaseStaleWrite.promise;
        }
        return originalTransaction(work) as Promise<T>;
      };

      const staleResolution = resolveImport({
        database,
        prepared,
        decisions: [],
      });
      await staleWriteReached.promise;

      const concurrent = await prepareImport({
        database,
        prepared,
        recipe,
        sources: [
          source({ key: "additional", selection: "Extra", value: "South" }),
        ],
      });
      expect(concurrent.prepared.state).toBe("needs-review");
      releaseStaleWrite.resolve();
      await expect(staleResolution).rejects.toMatchObject({
        code: "DB_STALE_IMPORT_PLAN",
      });

      const inspection = await inspectImport({
        database,
        prepared,
        page: { limit: 10 },
      });
      expect(inspection.prepared.state).toBe("needs-review");
      expect(inspection.routes).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            source: "additional",
            selection: "Extra",
            destination: null,
          }),
        ]),
      );
      expect(inspection.conflicts).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "missing-destination",
            source: "additional",
            selection: "Extra",
          }),
        ]),
      );
    } finally {
      engine.transaction = originalTransaction;
      releaseStaleWrite.resolve();
      await prepared.close();
      await database.close();
    }
  });
}
