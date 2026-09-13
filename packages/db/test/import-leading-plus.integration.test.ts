import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { engineOf } from "../src/database.js";
import {
  applyImport,
  inspectImport,
  prepareImport,
  resolveImport,
} from "../src/import/operations.js";
import { draftImportRecipe } from "../src/import/recipe.js";
import { valueForColumn } from "../src/import/inference.js";
import type { ImportSource } from "../src/import/types.js";
import { createDatabase, createPreparedImport } from "../src/node.js";

const directories: string[] = [];

test("leading-plus integers retain the signed 64-bit range", () => {
  expect(() =>
    valueForColumn(
      { kind: "number", raw: "+9223372036854775808" },
      { name: "Value", type: "integer" },
    ),
  ).toThrow(
    expect.objectContaining({
      code: "DB_IMPORT_VALUE_CONFLICT",
    }),
  );
});

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

function source(): ImportSource {
  const bytes = new TextEncoder().encode("synthetic leading plus numbers");
  return {
    key: "numbers",
    readerVersion: "synthetic-leading-plus-1",
    bytes: {
      name: "numbers.synthetic",
      size: bytes.length,
      async readAt(offset, length) {
        return bytes.slice(offset, offset + length);
      },
    },
    selections: [
      {
        key: "Numbers",
        label: "Numbers",
        async open() {
          return {
            columns: ["Value"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: { Value: { kind: "number" as const, raw: "+0" } },
                },
                {
                  sourceRow: 3,
                  cells: {
                    Value: {
                      kind: "formula" as const,
                      formula: "=1",
                      cached: { kind: "number" as const, raw: "+1" },
                    },
                  },
                },
                {
                  sourceRow: 4,
                  cells: {
                    Value: {
                      kind: "number" as const,
                      raw: "+9223372036854775807",
                    },
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
  test(`${format}: leading-plus integers infer and apply as integers`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-plus-integer-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `numbers.${format}`),
      format,
    });
    const input = source();
    const recipe = await draftImportRecipe({ sources: [input] });
    const prepared = await createPreparedImport({
      path: path.join(directory, "numbers.ccplan"),
      database,
      recipe,
      baselineRevision: 0n,
    });
    try {
      await prepareImport({
        database,
        prepared,
        recipe,
        sources: [input],
      });
      const inspection = await inspectImport({
        prepared,
        page: { limit: 3 },
      });
      expect(inspection.routes[0]?.inferredColumns).toEqual([
        { name: "Value", type: "integer" },
      ]);

      const approved = await resolveImport({
        database,
        prepared,
        decisions: [],
      });
      expect(approved.state).toBe("ready");
      if (approved.state !== "ready") throw new Error("Plan was not ready");
      await applyImport({
        database,
        prepared,
        approved,
        requestId: `leading-plus-${format}`,
      });

      await expect(
        engineOf(database).query(
          'SELECT "Value" AS value FROM "Numbers" ORDER BY record_id',
        ),
      ).resolves.toEqual([
        { value: 0n },
        { value: 1n },
        { value: 9_223_372_036_854_775_807n },
      ]);
    } finally {
      await prepared.close();
      await database.close();
    }
  });
}
