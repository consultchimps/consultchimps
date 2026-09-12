import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { inspectDatabase } from "../src/database.js";
import { inspectImport } from "../src/import/inspection.js";
import {
  applyImport,
  prepareImport,
  resolveImport,
} from "../src/import/operations.js";
import type {
  ImportCell,
  ImportRecipe,
  ImportSource,
} from "../src/import/types.js";
import { createDatabase, createPreparedImport } from "../src/node.js";
import { PREPARED_ROW_TABLE, preparedEngineOf } from "../src/prepared.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

function sourceWithSelections(
  values: readonly (readonly [string, ImportCell])[],
  verifyUnchanged?: () => Promise<void>,
): ImportSource {
  const bytes = new TextEncoder().encode("synthetic atomic import");
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
    ...(verifyUnchanged === undefined ? {} : { verifyUnchanged }),
    selections: values.map(([key, value]) => ({
      key,
      label: key,
      async open() {
        return {
          columns: ["Value"],
          async *batches() {
            yield [{ sourceRow: 2, cells: { Value: value } }];
          },
          async close() {},
        };
      },
    })),
  };
}

function recipeFor(...selections: readonly string[]): ImportRecipe {
  return {
    version: 1,
    routes: selections.map((selection) => ({
      source: "submission",
      selection,
      destination: {
        kind: "new-table",
        schema: {
          name: selection,
          columns: [{ name: "Value", type: "text" }],
          recordId: { prefix: selection.toUpperCase(), padding: 6 },
          foreignKeys: [],
        },
      },
      columns: [{ source: "Value", target: "Value", type: "text" }],
    })),
  };
}

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format}: a row failure rolls back the complete import`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-import-rollback-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const recipe = recipeFor("Good", "Bad");
    const prepared = await createPreparedImport({
      path: path.join(directory, "review.data"),
      database,
      recipe,
      baselineRevision: 0n,
    });
    try {
      await prepareImport({
        database,
        prepared,
        recipe,
        sources: [
          sourceWithSelections([
            ["Good", { kind: "string", value: "kept only on success" }],
            ["Bad", { kind: "string", value: "valid during review" }],
          ]),
        ],
      });
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [],
      });
      expect(approved.state).toBe("ready");
      if (approved.state !== "ready") throw new Error("Plan failed review");
      await preparedEngineOf(prepared).execute(
        `UPDATE ${PREPARED_ROW_TABLE} SET values_json = ? WHERE values_json LIKE ?`,
        [
          JSON.stringify({ Value: { kind: "error", error: "#VALUE!" } }),
          "%valid during review%",
        ],
      );
      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: "atomic-apply",
        }),
      ).rejects.toMatchObject({ code: "DB_IMPORT_ERROR_CELL" });
      const inspection = await inspectDatabase({ database });
      expect(inspection.tables).toEqual([]);
      expect(inspection.captures).toBe(0n);
      expect(inspection.completedImports).toBe(0n);
      expect(inspection.revision).toBe(0n);
    } finally {
      await prepared.close();
      await database.close();
    }
  });
}

test("source mutation after parsing removes staged rows", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-import-mutation-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "workspace.sqlite"),
    format: "sqlite",
  });
  const recipe = recipeFor("Inventory");
  const prepared = await createPreparedImport({
    path: path.join(directory, "review.data"),
    database,
    recipe,
    baselineRevision: 0n,
  });
  let verificationCount = 0;
  try {
    await expect(
      prepareImport({
        database,
        prepared,
        recipe,
        sources: [
          sourceWithSelections(
            [["Inventory", { kind: "string", value: "North" }]],
            async () => {
              verificationCount += 1;
              if (verificationCount === 2) {
                throw new Error("source changed after parsing");
              }
            },
          ),
        ],
      }),
    ).rejects.toThrow("source changed after parsing");
    const inspection = await inspectImport({
      prepared,
      page: { limit: 10 },
    });
    expect(inspection.examples).toEqual([]);
    expect(inspection.routes).toEqual([]);
    expect(inspection.capturedRows).toBe(0n);
  } finally {
    await prepared.close();
    await database.close();
  }
});

test("reusing a prepared binding rejects changed source content", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-import-rebind-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "workspace.sqlite"),
    format: "sqlite",
  });
  const recipe = recipeFor("Inventory");
  const prepared = await createPreparedImport({
    path: path.join(directory, "review.data"),
    database,
    recipe,
    baselineRevision: 0n,
  });
  try {
    const original = sourceWithSelections([
      ["Inventory", { kind: "string", value: "North" }],
    ]);
    await prepareImport({
      database,
      prepared,
      recipe,
      sources: [original],
    });
    const changedBytes = new TextEncoder().encode("changed submission bytes");
    await expect(
      prepareImport({
        database,
        prepared,
        recipe,
        sources: [
          {
            ...original,
            bytes: {
              name: "submission.xlsx",
              size: changedBytes.length,
              async readAt(offset, length) {
                return changedBytes.slice(offset, offset + length);
              },
            },
          },
        ],
      }),
    ).rejects.toMatchObject({ code: "DB_PREPARED_SOURCE_CHANGED" });
  } finally {
    await prepared.close();
    await database.close();
  }
});
