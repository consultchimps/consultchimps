import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, describe, expect, test } from "vitest";

import {
  createDatabaseHandle,
  engineOf,
  inspectDatabase,
} from "../src/database.js";
import type {
  DatabaseEngine,
  EngineRow,
  EngineTransaction,
} from "../src/internal/engine.js";
import {
  CAPTURE_TABLE,
  SOURCE_CONTENT_TABLE,
  SOURCE_FILE_TABLE,
} from "../src/metadata.js";
import { createDatabase, createImportBatch } from "../src/node.js";
import type { ImportBatch } from "../src/prepared.js";
import { applySchema, planSchema } from "../src/records.js";
import { applyImport } from "../src/import/apply.js";
import { recordBatch } from "../src/import/deliveries.js";
import { prepareImport } from "../src/import/prepare.js";
import { resolveImport } from "../src/import/resolve.js";
import { checkpointDatabaseWrite } from "../src/write-completion.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

class CheckpointEngine implements DatabaseEngine {
  readonly format = "sqlite" as const;
  readonly interruptible = false;
  checkpointCalls = 0;
  checkpointFailure: unknown;

  async execute(): Promise<void> {}

  async query(): Promise<readonly EngineRow[]> {
    return [];
  }

  async bulkInsert(): Promise<void> {}

  async transaction<T>(
    work: (transaction: EngineTransaction) => Promise<T>,
  ): Promise<T> {
    return work(this);
  }

  async readTransaction<T>(
    work: (transaction: EngineTransaction) => Promise<T>,
  ): Promise<T> {
    return work(this);
  }

  async checkpoint(): Promise<void> {
    this.checkpointCalls += 1;
    if (this.checkpointFailure !== undefined) throw this.checkpointFailure;
  }

  interrupt(): void {}

  async close(): Promise<void> {}
}

test("checkpoint completion retains a failed cause and retries only the checkpoint", async () => {
  const engine = new CheckpointEngine();
  const database = await createDatabaseHandle(engine);
  const result = {
    operation: "db.synthetic",
    artifacts: [],
    warnings: [],
    metrics: {},
    databaseWrite: "committed",
  } as const;
  const failure = new Error("synthetic checkpoint failure");
  engine.checkpointFailure = failure;

  const failed = await checkpointDatabaseWrite({ database, result });
  expect(failed.result).toBe(result);
  expect(failed.checkpoint).toMatchObject({
    state: "checkpoint-required",
    code: "DB_CHECKPOINT_REQUIRED",
    error: { cause: failure },
  });
  expect(engine.checkpointCalls).toBe(1);

  engine.checkpointFailure = undefined;
  await expect(
    checkpointDatabaseWrite({ database, result: failed.result }),
  ).resolves.toEqual({
    result,
    checkpoint: { state: "checkpoint-completed" },
  });
  expect(engine.checkpointCalls).toBe(2);

  await checkpointDatabaseWrite({
    database,
    result: { ...result, databaseWrite: "unchanged" },
  });
  expect(engine.checkpointCalls).toBe(3);
  await database.close();
});

describe.each(["sqlite", "duckdb"] as const)(
  "%s database write completion",
  (format) => {
    test("marks schema, import receipt, and batch mutations at their transaction branches", async () => {
      const directory = await mkdtemp(path.join(tmpdir(), "cc-db-write-"));
      directories.push(directory);
      const { database } = await createDatabase({
        path: path.join(directory, `workspace.${format}`),
        format,
      });
      let batch: ImportBatch | undefined;
      try {
        const emptySchema = await planSchema({
          database,
          schema: { version: 1, tables: [] },
        });
        await expect(
          applySchema({ database, plan: emptySchema }),
        ).resolves.toMatchObject({ databaseWrite: "unchanged" });

        const tableSchema = await planSchema({
          database,
          schema: {
            version: 1,
            tables: [
              {
                name: "Records",
                recordId: { prefix: "REC", padding: 4 },
                columns: [{ name: "Value", type: "text" }],
              },
            ],
          },
        });
        await expect(
          applySchema({ database, plan: tableSchema }),
        ).resolves.toMatchObject({ databaseWrite: "committed" });

        batch = await createImportBatch({
          path: path.join(directory, "empty.ccplan"),
          database,
          profile: { version: 1, routes: [] },
          baselineRevision: (await inspectDatabase({ database })).revision,
        });
        await prepareImport({
          database,
          prepared: batch,
          profile: { version: 1, routes: [] },
          sources: [],
        });
        const approved = await resolveImport({
          database,
          prepared: batch,
          decisions: [],
        });
        if (approved.state !== "ready") throw new Error("Plan needs review");
        const applied = await applyImport({
          database,
          prepared: batch,
          approved,
          requestId: `${format}-empty-import`,
        });
        expect(applied).toMatchObject({
          databaseWrite: "committed",
          metrics: { rowsImported: 0, rowsReused: 0 },
        });
        await expect(
          applyImport({
            database,
            prepared: batch,
            approved,
            requestId: `${format}-empty-import`,
          }),
        ).resolves.toMatchObject({ databaseWrite: "unchanged" });

        await engineOf(database).transaction(async (transaction) => {
          await transaction.execute(
            `INSERT INTO ${SOURCE_CONTENT_TABLE} VALUES (?, ?)`,
            [`${format}-content`, 0n],
          );
          await transaction.execute(
            `INSERT INTO ${SOURCE_FILE_TABLE} VALUES (?, ?, ?)`,
            ["SRC-000001", `${format}-content`, "synthetic.xlsx"],
          );
          await transaction.execute(
            `INSERT INTO ${CAPTURE_TABLE} VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
            [
              "CAP-000001",
              "SRC-000001",
              "synthetic",
              "Sheet1",
              "Sheet1",
              "synthetic-reader-1",
              "completed",
              0n,
              "[]",
            ],
          );
        });
        const recorded = await recordBatch({
          database,
          captureIds: ["CAP-000001"],
          context: { label: "Synthetic batch", scope: { kind: "full" } },
          requestId: `${format}-batch`,
        });
        expect(recorded.databaseWrite).toBe("committed");
        await expect(
          recordBatch({
            database,
            captureIds: ["CAP-000001"],
            context: { label: "Synthetic batch", scope: { kind: "full" } },
            requestId: `${format}-batch`,
          }),
        ).resolves.toMatchObject({ databaseWrite: "unchanged" });
      } finally {
        await batch?.close();
        await database.close();
      }
    });
  },
);
