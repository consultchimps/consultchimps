import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { inspectDatabase } from "../src/database.js";
import { listBatches, recordBatch } from "../src/import/deliveries.js";
import {
  applyImport,
  prepareImport,
  resolveImport,
} from "../src/import/operations.js";
import type { ImportProfile, ImportSource } from "../src/import/types.js";
import { createDatabase, createImportBatch } from "../src/node.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

const profile: ImportProfile = {
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
          recordId: { prefix: "ITEM", padding: 4 },
        },
      },
      columns: [{ source: "Value", target: "Value", type: "text" }],
    },
  ],
};

function source(): ImportSource {
  const bytes = new TextEncoder().encode("synthetic request collision");
  return {
    key: "submission",
    readerVersion: "synthetic-request-collision-1",
    bytes: {
      name: "submission.xlsx",
      size: bytes.length,
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

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format}: request IDs cannot cross import and standalone delivery operations`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-request-kind-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const prepared = await createImportBatch({
      path: path.join(directory, "initial.ccplan"),
      database,
      profile,
      baselineRevision: 0n,
    });
    const delivery = {
      label: "Synthetic delivery",
      scope: { kind: "full" as const },
    };
    let emptyPlan: Awaited<ReturnType<typeof createImportBatch>> | undefined;
    try {
      await prepareImport({
        database,
        prepared,
        profile,
        sources: [source()],
      });
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [],
      });
      if (approved.state !== "ready") throw new Error("Plan failed review");
      const paired = await applyImport({
        database,
        prepared,
        approved,
        requestId: `paired-${format}`,
        batchContext: delivery,
      });
      const pairedRetry = await applyImport({
        database,
        prepared,
        approved,
        requestId: `paired-${format}`,
        batchContext: delivery,
      });
      expect(pairedRetry).toMatchObject({
        importIds: paired.importIds,
        batchId: paired.batchId,
        metrics: { rowsImported: 0, rowsReused: 1, batchesRecorded: 0 },
      });

      const standalone = await recordBatch({
        database,
        captureIds: paired.captureIds,
        context: { ...delivery, label: "Standalone delivery" },
        requestId: `delivery-first-${format}`,
      });
      const standaloneRetry = await recordBatch({
        database,
        captureIds: paired.captureIds,
        context: { ...delivery, label: "Standalone delivery" },
        requestId: `delivery-first-${format}`,
      });
      expect(standaloneRetry.batch).toEqual(standalone.batch);
      expect(standaloneRetry.metrics.batchesRecorded).toBe(0);

      const emptyRecipe: ImportProfile = { version: 1, routes: [] };
      const baseline = await inspectDatabase({ database });
      emptyPlan = await createImportBatch({
        path: path.join(directory, "empty.ccplan"),
        database,
        profile: emptyRecipe,
        baselineRevision: baseline.revision,
      });
      await prepareImport({
        database,
        prepared: emptyPlan,
        profile: emptyRecipe,
        sources: [],
      });
      const emptyApproved = await resolveImport({
        database,
        prepared: emptyPlan,
        decisions: [],
      });
      if (emptyApproved.state !== "ready") {
        throw new Error("Empty plan failed review");
      }

      const beforeDeliveryCollision = await inspectDatabase({ database });
      const beforeDeliveries = await listBatches({ database, limit: 10 });
      for (const withDelivery of [false, true]) {
        await expect(
          applyImport({
            database,
            prepared: emptyPlan,
            approved: emptyApproved,
            requestId: `delivery-first-${format}`,
            ...(withDelivery ? { batchContext: delivery } : {}),
          }),
        ).rejects.toMatchObject({ code: "DB_REQUEST_ID_CONFLICT" });
        expect(await inspectDatabase({ database })).toEqual(
          beforeDeliveryCollision,
        );
        expect(await listBatches({ database, limit: 10 })).toEqual(
          beforeDeliveries,
        );
      }
      expect(
        (
          await recordBatch({
            database,
            captureIds: paired.captureIds,
            context: { ...delivery, label: "Standalone delivery" },
            requestId: `delivery-first-${format}`,
          })
        ).batch,
      ).toEqual(standalone.batch);

      const imported = await applyImport({
        database,
        prepared: emptyPlan,
        approved: emptyApproved,
        requestId: `import-first-${format}`,
      });
      const importedRetry = await applyImport({
        database,
        prepared: emptyPlan,
        approved: emptyApproved,
        requestId: `import-first-${format}`,
      });
      expect(importedRetry.importIds).toEqual(imported.importIds);
      expect(importedRetry.metrics).toMatchObject({
        rowsImported: 0,
        rowsReused: 0,
      });

      const beforeImportCollision = await inspectDatabase({ database });
      const deliveriesBeforeImportCollision = await listBatches({
        database,
        limit: 10,
      });
      await expect(
        recordBatch({
          database,
          captureIds: paired.captureIds,
          context: delivery,
          requestId: `import-first-${format}`,
        }),
      ).rejects.toMatchObject({ code: "DB_REQUEST_ID_CONFLICT" });
      expect(await inspectDatabase({ database })).toEqual(
        beforeImportCollision,
      );
      expect(await listBatches({ database, limit: 10 })).toEqual(
        deliveriesBeforeImportCollision,
      );

      const pairedDeliveryRetry = await recordBatch({
        database,
        captureIds: paired.captureIds,
        context: delivery,
        requestId: `paired-${format}`,
      });
      expect(pairedDeliveryRetry).toMatchObject({
        metrics: { batchesRecorded: 0 },
        batch: { id: paired.batchId },
      });
      for (const changed of [
        {
          captureIds: paired.captureIds,
          context: { ...delivery, label: "Changed delivery" },
        },
        { captureIds: ["CAP-missing"], context: delivery },
      ]) {
        await expect(
          recordBatch({
            database,
            ...changed,
            requestId: `paired-${format}`,
          }),
        ).rejects.toMatchObject({ code: "DB_REQUEST_ID_CONFLICT" });
      }
      expect(await inspectDatabase({ database })).toEqual(
        beforeImportCollision,
      );
      expect(await listBatches({ database, limit: 10 })).toEqual(
        deliveriesBeforeImportCollision,
      );
      expect(
        (
          await applyImport({
            database,
            prepared: emptyPlan,
            approved: emptyApproved,
            requestId: `import-first-${format}`,
          })
        ).importIds,
      ).toEqual(imported.importIds);
    } finally {
      await emptyPlan?.close();
      await prepared.close();
      await database.close();
    }
  });
}
