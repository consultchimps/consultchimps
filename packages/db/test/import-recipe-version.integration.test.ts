import { access, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, describe, expect, test } from "vitest";

import { inspectDatabase, type DatabaseId } from "../src/database.js";
import type { DatabaseEngine } from "../src/internal/engine.js";
import type { ImportRecipe } from "../src/import/types.js";
import { createDatabase, createPreparedImport } from "../src/node.js";
import { createPreparedImportHandle } from "../src/prepared.js";
import { validateImportRecipe } from "../src/validators.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

const SYNTHETIC_DATABASE_ID = "DB-recipe-version" as DatabaseId;

function unsupportedRecipe(): ImportRecipe {
  // Simulates an untyped JavaScript caller crossing the public TypeScript contract.
  return { version: 2, routes: [] } as unknown as ImportRecipe;
}

test("rejects an unsupported runtime recipe version before prepared schema staging", async () => {
  let transactionCalls = 0;
  const unexpected = async (): Promise<never> => {
    throw new Error("The prepared engine must remain untouched");
  };
  const engine: DatabaseEngine = {
    format: "sqlite",
    interruptible: false,
    execute: unexpected,
    query: unexpected,
    bulkInsert: unexpected,
    async transaction() {
      transactionCalls += 1;
      throw new Error("The prepared transaction must not start");
    },
    readTransaction: unexpected,
    checkpoint: unexpected,
    interrupt() {},
    close: async () => {},
  };

  expect(() => validateImportRecipe(unsupportedRecipe())).toThrowError(
    expect.objectContaining({ code: "DB_INVALID_RECIPE" }),
  );
  await expect(
    createPreparedImportHandle({
      engine,
      databaseId: SYNTHETIC_DATABASE_ID,
      baselineRevision: 0n,
      baselineSchemaFingerprint: "synthetic-fingerprint",
      recipe: unsupportedRecipe(),
    }),
  ).rejects.toMatchObject({ code: "DB_INVALID_RECIPE" });
  expect(transactionCalls).toBe(0);
});

describe.each(["sqlite", "duckdb"] as const)(
  "%s prepared recipe version validation",
  (format) => {
    test("preserves absent and existing destinations when the version is unsupported", async () => {
      const directory = await mkdtemp(
        path.join(tmpdir(), "cc-recipe-version-"),
      );
      directories.push(directory);
      const { database } = await createDatabase({
        path: path.join(directory, `workspace.${format}`),
        format,
      });
      const absentPath = path.join(directory, "absent.ccplan");
      const existingPath = path.join(directory, "existing.ccplan");
      const sentinel = new TextEncoder().encode("existing plan sentinel");
      await writeFile(existingPath, sentinel);

      try {
        const baselineRevision = (await inspectDatabase({ database })).revision;
        await expect(
          createPreparedImport({
            path: absentPath,
            database,
            baselineRevision,
            recipe: unsupportedRecipe(),
          }),
        ).rejects.toMatchObject({ code: "DB_INVALID_RECIPE" });
        await expect(access(absentPath)).rejects.toMatchObject({
          code: "ENOENT",
        });

        await expect(
          createPreparedImport({
            path: existingPath,
            database,
            baselineRevision,
            recipe: unsupportedRecipe(),
            overwrite: true,
          }),
        ).rejects.toMatchObject({ code: "DB_INVALID_RECIPE" });
        expect(Buffer.from(await readFile(existingPath)).equals(sentinel)).toBe(
          true,
        );

        const supportedPath = path.join(directory, "supported.ccplan");
        const supported = await createPreparedImport({
          path: supportedPath,
          database,
          baselineRevision,
          recipe: { version: 1, routes: [] },
        });
        await supported.close();
        await expect(access(supportedPath)).resolves.toBeUndefined();
      } finally {
        await database.close();
      }
    });
  },
);
