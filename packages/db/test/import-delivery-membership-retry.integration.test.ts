import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { engineOf, inspectDatabase } from "../src/database.js";
import {
  applyImport,
  prepareImport,
  resolveImport,
} from "../src/import/operations.js";
import type { ImportProfile, ImportSource } from "../src/import/types.js";
import { DELIVERY_MEMBERSHIP_TABLE } from "../src/metadata.js";
import { createDatabase, createImportBatch } from "../src/node.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

const selections = ["North", "South"] as const;
const profile: ImportProfile = {
  version: 1,
  routes: selections.map((selection) => ({
    source: "submission",
    selection,
    destination: {
      kind: "new-table",
      schema: {
        name: `${selection}Inventory`,
        columns: [{ name: "Value", type: "text" }],
        recordId: { prefix: selection.toUpperCase(), padding: 4 },
        foreignKeys: [],
      },
    },
    columns: [{ source: "Value", target: "Value", type: "text" }],
  })),
};

function importSource(): ImportSource {
  const bytes = new TextEncoder().encode("synthetic delivery membership");
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
    selections: selections.map((selection, index) => ({
      key: selection,
      label: selection,
      async open() {
        return {
          columns: ["Value"],
          async *batches() {
            yield [
              {
                sourceRow: index + 2,
                cells: {
                  Value: { kind: "string" as const, value: selection },
                },
              },
            ];
          },
          async close() {},
        };
      },
    })),
  };
}

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format}: import delivery retries reject damaged capture memberships without writes`, async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "cc-delivery-membership-"),
    );
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const prepared = await createImportBatch({
      path: path.join(directory, "review.ccplan"),
      database,
      profile,
      baselineRevision: 0n,
    });
    const requestId = `delivery-membership-${format}`;
    const delivery = {
      label: "Synthetic delivery",
      scope: { kind: "full" as const },
    };
    try {
      await prepareImport({
        database,
        prepared,
        profile,
        sources: [importSource()],
      });
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [],
      });
      if (approved.state !== "ready") throw new Error("Plan failed review");
      const applied = await applyImport({
        database,
        prepared,
        approved,
        requestId,
        batchContext: delivery,
      });
      const batchId = applied.batchId;
      const firstCapture = applied.captureIds[0];
      if (batchId === undefined || firstCapture === undefined) {
        throw new Error("Import did not return its delivery and captures");
      }
      expect(applied.captureIds).toHaveLength(2);
      const engine = engineOf(database);

      const assertRejectedWithoutMutation = async (): Promise<void> => {
        const beforeInspection = await inspectDatabase({ database });
        const beforeMemberships = await engine.query(
          `SELECT delivery_id, capture_id FROM ${DELIVERY_MEMBERSHIP_TABLE} ORDER BY capture_id`,
        );
        await expect(
          applyImport({
            database,
            prepared,
            approved,
            requestId,
            batchContext: delivery,
          }),
        ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });
        expect(await inspectDatabase({ database })).toEqual(beforeInspection);
        expect(
          await engine.query(
            `SELECT delivery_id, capture_id FROM ${DELIVERY_MEMBERSHIP_TABLE} ORDER BY capture_id`,
          ),
        ).toEqual(beforeMemberships);
      };

      await engine.execute(
        `DELETE FROM ${DELIVERY_MEMBERSHIP_TABLE} WHERE delivery_id = ? AND capture_id = ?`,
        [batchId, firstCapture],
      );
      await assertRejectedWithoutMutation();
      await engine.execute(
        `INSERT INTO ${DELIVERY_MEMBERSHIP_TABLE} VALUES (?, ?)`,
        [batchId, firstCapture],
      );

      await engine.execute(
        `INSERT INTO ${DELIVERY_MEMBERSHIP_TABLE} VALUES (?, ?)`,
        [batchId, "CAP-extra"],
      );
      await assertRejectedWithoutMutation();
      await engine.execute(
        `DELETE FROM ${DELIVERY_MEMBERSHIP_TABLE} WHERE delivery_id = ? AND capture_id = ?`,
        [batchId, "CAP-extra"],
      );

      await engine.execute(
        `DELETE FROM ${DELIVERY_MEMBERSHIP_TABLE} WHERE delivery_id = ? AND capture_id = ?`,
        [batchId, firstCapture],
      );
      await engine.execute(
        `INSERT INTO ${DELIVERY_MEMBERSHIP_TABLE} VALUES (?, ?)`,
        [batchId, "CAP-substituted"],
      );
      await assertRejectedWithoutMutation();
      await engine.execute(
        `DELETE FROM ${DELIVERY_MEMBERSHIP_TABLE} WHERE delivery_id = ? AND capture_id = ?`,
        [batchId, "CAP-substituted"],
      );
      await engine.execute(
        `INSERT INTO ${DELIVERY_MEMBERSHIP_TABLE} VALUES (?, ?)`,
        [batchId, firstCapture],
      );

      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId,
          batchContext: delivery,
        }),
      ).resolves.toMatchObject({
        batchId,
        captureIds: applied.captureIds,
        metrics: {
          rowsImported: 0,
          rowsReused: 2,
          batchesRecorded: 0,
        },
      });
    } finally {
      await prepared.close();
      await database.close();
    }
  });
}
