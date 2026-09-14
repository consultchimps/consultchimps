import { access, mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, describe, expect, test } from "vitest";

import { inspectDatabase, type DatabaseId } from "../src/database.js";
import type {
  DatabaseEngine,
  EngineRow,
  EngineTransaction,
} from "../src/internal/engine.js";
import { NodeSqliteEngine } from "../src/engines/sqlite/node.js";
import {
  createDatabase,
  createImportBatch,
  openImportBatch,
} from "../src/node.js";
import {
  createImportBatchHandle,
  preparedRef,
  readPreparedRecipe,
  updatePreparedPlan,
} from "../src/prepared.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

class SentinelEngine implements DatabaseEngine {
  readonly format = "sqlite" as const;
  readonly interruptible = false;
  transactions = 0;
  executions = 0;

  async execute(): Promise<void> {
    this.executions += 1;
  }

  async query(): Promise<readonly EngineRow[]> {
    return [];
  }

  async bulkInsert(): Promise<void> {}

  async transaction<T>(
    work: (transaction: EngineTransaction) => Promise<T>,
  ): Promise<T> {
    this.transactions += 1;
    return work(this);
  }

  async readTransaction<T>(
    work: (transaction: EngineTransaction) => Promise<T>,
  ): Promise<T> {
    return work(this);
  }

  async checkpoint(): Promise<void> {}

  interrupt(): void {}

  async close(): Promise<void> {}
}

test("rejects a negative baseline before starting prepared schema creation", async () => {
  const engine = new SentinelEngine();
  await expect(
    createImportBatchHandle({
      engine,
      databaseId: "DB-baseline-revision" as DatabaseId,
      baselineRevision: -1n,
      baselineSchemaFingerprint: "synthetic-fingerprint",
      profile: { version: 1, routes: [] },
    }),
  ).rejects.toMatchObject({
    code: "DB_INVALID_PREPARED_IMPORT",
    message: "The import batch has an invalid baseline revision.",
  });
  expect(engine.transactions).toBe(0);
  expect(engine.executions).toBe(0);
});

test("opens independently encoded format 3 metadata after the vocabulary rename", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-format3-profile-"));
  directories.push(directory);
  const planPath = path.join(directory, "saved.ccplan");
  const engine = NodeSqliteEngine.create(planPath);
  try {
    await engine.transaction(async (transaction) => {
      await transaction.execute(
        "CREATE TABLE _consultchimps_prepared (format_version BIGINT NOT NULL, plan_id VARCHAR PRIMARY KEY, database_id VARCHAR NOT NULL, baseline_revision BIGINT NOT NULL, schema_fingerprint VARCHAR NOT NULL, plan_revision BIGINT NOT NULL, state VARCHAR NOT NULL, recipe_json VARCHAR NOT NULL, conflicts_json VARCHAR NOT NULL, decisions_json VARCHAR NOT NULL, review_fingerprint VARCHAR NOT NULL)",
      );
      await transaction.execute(
        "CREATE TABLE _consultchimps_prepared_captures (capture_id VARCHAR PRIMARY KEY, source_file_id VARCHAR, source_key VARCHAR NOT NULL, display_name VARCHAR NOT NULL, selection_key VARCHAR NOT NULL, selection_label VARCHAR NOT NULL, reader_version VARCHAR NOT NULL, content_hash VARCHAR NOT NULL, byte_count BIGINT NOT NULL, reused BIGINT NOT NULL, row_count BIGINT NOT NULL, columns_json VARCHAR NOT NULL, row_checksum VARCHAR NOT NULL)",
      );
      await transaction.execute(
        "CREATE TABLE _consultchimps_prepared_bindings (source_key VARCHAR NOT NULL, selection_key VARCHAR NOT NULL, capture_id VARCHAR NOT NULL, display_name VARCHAR NOT NULL, PRIMARY KEY(source_key, selection_key))",
      );
      await transaction.execute(
        "CREATE TABLE _consultchimps_prepared_rows (capture_id VARCHAR NOT NULL, source_row BIGINT NOT NULL, values_json VARCHAR NOT NULL, PRIMARY KEY(capture_id, source_row))",
      );
      await transaction.execute(
        "INSERT INTO _consultchimps_prepared VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        [
          3n,
          "PLAN-format3-vocabulary",
          "DB-format3-vocabulary",
          0n,
          "schema-fingerprint",
          1n,
          "ready",
          '{"routes":[],"version":1}',
          "[]",
          "[]",
          "eb00b2f7931bf4aeb598d463c48ff65bd8e942ab05d1c1b371b9a973420a757d",
        ],
      );
    });
  } finally {
    await engine.close();
  }

  const batch = await openImportBatch({ path: planPath, readonly: true });
  try {
    await expect(readPreparedRecipe(batch)).resolves.toMatchObject({
      profile: { version: 1, routes: [] },
    });
  } finally {
    await batch.close();
  }
});

describe.each(["sqlite", "duckdb"] as const)(
  "%s prepared baseline revision validation",
  (format) => {
    test("preserves new and overwritten destinations and accepts non-negative revisions", async () => {
      const directory = await mkdtemp(
        path.join(tmpdir(), "cc-prepared-baseline-"),
      );
      directories.push(directory);
      const { database } = await createDatabase({
        path: path.join(directory, `workspace.${format}`),
        format,
      });
      const baselineRevision = (await inspectDatabase({ database })).revision;
      const absentPath = path.join(directory, "absent.ccplan");
      const existingPath = path.join(directory, "existing.ccplan");
      const existing = await createImportBatch({
        path: existingPath,
        database,
        baselineRevision,
        profile: { version: 1, routes: [] },
      });
      await existing.close();
      const existingBytes = await readFile(existingPath);

      try {
        await expect(
          createImportBatch({
            path: absentPath,
            database,
            baselineRevision: -1n,
            profile: { version: 1, routes: [] },
          }),
        ).rejects.toMatchObject({ code: "DB_INVALID_PREPARED_IMPORT" });
        await expect(access(absentPath)).rejects.toMatchObject({
          code: "ENOENT",
        });

        await expect(
          createImportBatch({
            path: existingPath,
            database,
            baselineRevision: -1n,
            profile: { version: 1, routes: [] },
            overwrite: true,
          }),
        ).rejects.toMatchObject({ code: "DB_INVALID_PREPARED_IMPORT" });
        expect(
          Buffer.from(await readFile(existingPath)).equals(existingBytes),
        ).toBe(true);

        for (const validRevision of [0n, 7n]) {
          const validPath = path.join(
            directory,
            `valid-${validRevision.toString()}.ccplan`,
          );
          const valid = await createImportBatch({
            path: validPath,
            database,
            baselineRevision: validRevision,
            profile: { version: 1, routes: [] },
          });
          const beforeInvalidUpdate = await preparedRef(valid);
          expect(beforeInvalidUpdate.baselineRevision).toBe(validRevision);
          await expect(
            updatePreparedPlan({
              prepared: valid,
              profile: { version: 1, routes: [] },
              conflicts: [],
              ready: true,
              baselineRevision: -1n,
            }),
          ).rejects.toMatchObject({
            code: "DB_INVALID_PREPARED_IMPORT",
            message: "The import batch has an invalid baseline revision.",
          });
          expect(await preparedRef(valid)).toEqual(beforeInvalidUpdate);
          await valid.close();
        }
      } finally {
        await database.close();
      }
    });
  },
);
