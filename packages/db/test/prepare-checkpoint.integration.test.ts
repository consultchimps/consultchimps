import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { expect, test, vi } from "vitest";

import {
  applyImport,
  draftImportProfile,
  inspectDatabase,
  inspectImport,
  prepareImport,
  resolveImport,
  type ImportSource,
} from "../src/index.js";
import {
  createDatabase,
  createImportBatch,
  openImportBatch,
} from "../src/node.js";
import { preparedEngineOf } from "../src/prepared.js";

test.each(["sqlite", "duckdb"] as const)(
  "%s reports a committed preparation after checkpoint failure and recovers without source reads",
  async (format) => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "cc-batch-checkpoint-"),
    );
    const { database } = await createDatabase({
      path: path.join(directory, `working.${format}`),
      format,
    });
    const bytes = new TextEncoder().encode("synthetic checkpoint capture");
    const open = vi.fn(async () => ({
      columns: ["Value"],
      async *batches() {
        yield [
          {
            sourceRow: 2,
            cells: { Value: { kind: "number" as const, raw: "42" } },
          },
        ];
      },
      async close() {},
    }));
    const source: ImportSource = {
      key: "Inventory",
      readerVersion: "checkpoint-test-1",
      bytes: {
        name: "inventory.xlsx",
        size: bytes.length,
        async readAt(offset, length) {
          return bytes.slice(offset, offset + length);
        },
      },
      selections: [{ key: "Sheet1", label: "Inventory", open }],
    };
    const profile = await draftImportProfile({ sources: [source] });
    const batchPath = path.join(directory, "review.ccplan");
    let prepared = await createImportBatch({
      path: batchPath,
      database,
      profile,
      baselineRevision: 0n,
    });
    const failure = new Error("Injected batch checkpoint failure");
    const checkpoint = vi
      .spyOn(preparedEngineOf(prepared), "checkpoint")
      .mockRejectedValueOnce(failure);
    try {
      const before = await inspectImport({
        database,
        prepared,
        page: { limit: 1 },
      });
      await expect(
        prepareImport({ database, prepared, profile, sources: [source] }),
      ).rejects.toMatchObject({
        code: "DB_BATCH_CHECKPOINT_REQUIRED",
        cause: failure,
        details: {
          batchUpdated: true,
          checkpointRequired: true,
          batchId: prepared.id,
        },
      });
      expect(checkpoint).toHaveBeenCalledOnce();
      const updated = await inspectImport({
        database,
        prepared,
        page: { limit: 1 },
      });
      expect(updated.prepared.planRevision).toBeGreaterThan(
        before.prepared.planRevision,
      );
      expect(updated.prepared.state).toBe("needs-review");
      expect(updated.reviewRows).toBe(1n);
      expect(updated.examples).toHaveLength(1);
      expect(open).toHaveBeenCalledOnce();
      checkpoint.mockRestore();
      await prepared.close();
      prepared = await openImportBatch({ path: batchPath });
      const reopened = await inspectImport({
        database,
        prepared,
        page: { limit: 1 },
      });
      expect(reopened.prepared).toEqual(updated.prepared);
      expect(reopened.reviewRows).toBe(1n);
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [
          {
            kind: "route",
            source: source.key,
            selection: "Sheet1",
            destination: {
              kind: "new-table",
              schema: {
                name: "Inventory",
                recordId: { prefix: "INV", padding: 6 },
                columns: [{ name: "Value", type: "integer" }],
              },
            },
            columns: [{ source: "Value", target: "Value", type: "integer" }],
          },
        ],
      });
      if (approved.state !== "ready") throw new Error("Expected ready batch");
      await applyImport({
        database,
        prepared,
        approved,
        requestId: "recover-checkpoint",
      });
      expect((await inspectDatabase({ database })).tables[0]?.rowCount).toBe(
        1n,
      );
      expect(open).toHaveBeenCalledOnce();
    } finally {
      checkpoint.mockRestore();
      await prepared.close();
      await database.close();
      await rm(directory, { recursive: true, force: true });
    }
  },
);
