import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { inspectDatabase } from "../src/database.js";
import {
  applyImport,
  inspectImport,
  prepareImport,
} from "../src/import/operations.js";
import { preparedCaptures } from "../src/import/planning.js";
import type { ImportRecipe, ImportSource } from "../src/import/types.js";
import { createDatabase, createPreparedImport } from "../src/node.js";
import { PREPARED_CAPTURE_TABLE, preparedEngineOf } from "../src/prepared.js";
import type { DatabaseFormat } from "../src/schema.js";

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

function source(): ImportSource {
  const bytes = new TextEncoder().encode("synthetic capture counts");
  return {
    key: "submission",
    readerVersion: "synthetic-counts-1",
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

async function fixture(format: DatabaseFormat) {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-capture-counts-"));
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
  const outcome = await prepareImport({
    database,
    prepared,
    recipe,
    sources: [source()],
  });
  expect(outcome.prepared.state).toBe("ready");
  if (outcome.prepared.state !== "ready") {
    throw new Error("The synthetic import plan did not become ready.");
  }
  return { database, prepared, approved: outcome.prepared };
}

const invalidCounts = [
  { field: "byte_count", value: -1n },
  { field: "byte_count", value: BigInt(Number.MAX_SAFE_INTEGER) + 1n },
  { field: "row_count", value: -1n },
  { field: "row_count", value: BigInt(Number.MAX_SAFE_INTEGER) + 1n },
  { field: "reused", value: 2n },
] as const;

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format}: invalid persisted capture counts fail before destination writes`, async () => {
    const { database, prepared, approved } = await fixture(format);
    const engine = preparedEngineOf(prepared);
    try {
      for (const [index, invalid] of invalidCounts.entries()) {
        await engine.execute(
          `UPDATE ${PREPARED_CAPTURE_TABLE} SET ${invalid.field} = ?`,
          [invalid.value],
        );

        await expect(preparedCaptures(prepared)).rejects.toMatchObject({
          code: "DB_INVALID_PREPARED_IMPORT",
        });
        await expect(
          inspectImport({ database, prepared, page: { limit: 10 } }),
        ).rejects.toMatchObject({ code: "DB_INVALID_PREPARED_IMPORT" });
        await expect(
          applyImport({
            database,
            prepared,
            approved,
            requestId: `invalid-count-${format}-${index}`,
          }),
        ).rejects.toMatchObject({ code: "DB_INVALID_PREPARED_IMPORT" });
        expect(await inspectDatabase({ database })).toMatchObject({
          revision: 0n,
          tables: [],
          captures: 0n,
          completedImports: 0n,
        });

        await engine.execute(
          `UPDATE ${PREPARED_CAPTURE_TABLE} SET byte_count = ?, reused = 0, row_count = 1`,
          [BigInt(source().bytes.size)],
        );
      }
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: persisted capture count boundaries are accepted`, async () => {
    const { database, prepared } = await fixture(format);
    const engine = preparedEngineOf(prepared);
    const maximum = BigInt(Number.MAX_SAFE_INTEGER);
    try {
      await engine.execute(
        `UPDATE ${PREPARED_CAPTURE_TABLE} SET byte_count = ?, reused = 0, row_count = ?`,
        [maximum, maximum],
      );
      await expect(preparedCaptures(prepared)).resolves.toEqual([
        expect.objectContaining({
          byteCount: maximum,
          reused: false,
          rowCount: maximum,
        }),
      ]);

      await engine.execute(
        `UPDATE ${PREPARED_CAPTURE_TABLE} SET byte_count = 0, reused = 1, row_count = 0`,
      );
      await expect(preparedCaptures(prepared)).resolves.toEqual([
        expect.objectContaining({
          byteCount: 0n,
          reused: true,
          rowCount: 0n,
        }),
      ]);
    } finally {
      await prepared.close();
      await database.close();
    }
  });
}
