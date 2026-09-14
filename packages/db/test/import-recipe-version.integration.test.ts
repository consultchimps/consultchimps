import { access, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, describe, expect, test } from "vitest";

import { inspectDatabase, type DatabaseId } from "../src/database.js";
import type { DatabaseEngine } from "../src/internal/engine.js";
import type { ImportProfile } from "../src/import/types.js";
import { createDatabase, createImportBatch } from "../src/node.js";
import { createImportBatchHandle } from "../src/prepared.js";
import { validateImportProfile } from "../src/validators.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

const SYNTHETIC_DATABASE_ID = "DB-profile-version" as DatabaseId;

function unsupportedRecipe(): ImportProfile {
  // Simulates an untyped JavaScript caller crossing the public TypeScript contract.
  return { version: 2, routes: [] } as unknown as ImportProfile;
}

test("rejects an unsupported runtime profile version before prepared schema staging", async () => {
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

  expect(() => validateImportProfile(unsupportedRecipe())).toThrowError(
    expect.objectContaining({ code: "DB_INVALID_RECIPE" }),
  );
  await expect(
    createImportBatchHandle({
      engine,
      databaseId: SYNTHETIC_DATABASE_ID,
      baselineRevision: 0n,
      baselineSchemaFingerprint: "synthetic-fingerprint",
      profile: unsupportedRecipe(),
    }),
  ).rejects.toMatchObject({ code: "DB_INVALID_RECIPE" });
  expect(transactionCalls).toBe(0);
});

describe.each(["sqlite", "duckdb"] as const)(
  "%s prepared profile version validation",
  (format) => {
    test("preserves absent and existing destinations when the version is unsupported", async () => {
      const directory = await mkdtemp(
        path.join(tmpdir(), "cc-profile-version-"),
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
          createImportBatch({
            path: absentPath,
            database,
            baselineRevision,
            profile: unsupportedRecipe(),
          }),
        ).rejects.toMatchObject({ code: "DB_INVALID_RECIPE" });
        await expect(access(absentPath)).rejects.toMatchObject({
          code: "ENOENT",
        });

        await expect(
          createImportBatch({
            path: existingPath,
            database,
            baselineRevision,
            profile: unsupportedRecipe(),
            overwrite: true,
          }),
        ).rejects.toMatchObject({ code: "DB_INVALID_RECIPE" });
        expect(Buffer.from(await readFile(existingPath)).equals(sentinel)).toBe(
          true,
        );

        const supportedPath = path.join(directory, "supported.ccplan");
        const supported = await createImportBatch({
          path: supportedPath,
          database,
          baselineRevision,
          profile: { version: 1, routes: [] },
        });
        await supported.close();
        await expect(access(supportedPath)).resolves.toBeUndefined();
      } finally {
        await database.close();
      }
    });
  },
);
