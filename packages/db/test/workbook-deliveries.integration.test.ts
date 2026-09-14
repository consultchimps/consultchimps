import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  createScratchDirectory,
  openRandomAccessSource,
} from "@consultchimps/files";
import { writeTable } from "@consultchimps/xlsx";
import { afterEach, expect, test } from "vitest";

import { engineOf, inspectDatabase } from "../src/database.js";
import { recordBatch, listBatches } from "../src/import/deliveries.js";
import {
  prepareImport,
  inspectImport,
  resolveImport,
  applyImport,
} from "../src/import/operations.js";
import { draftImportProfile } from "../src/import/profile.js";
import {
  createDatabase,
  createImportBatch,
  openImportBatch,
  exportDatabase,
  openDatabase,
} from "../src/node.js";
import { createWorkbookImportSource } from "../src/workbook.js";
import { COUNTERS_TABLE } from "../src/metadata.js";

const directories: string[] = [];
afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format}: library workbook bridge captures once and delivery history paginates independently`, async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "cc-workbook-delivery-"),
    );
    directories.push(directory);
    const input = path.join(directory, "inventory.xlsx");
    await writeTable(
      input,
      {
        columns: ["Name", "Count"],
        rows: [
          { Name: "North", Count: 2 },
          { Name: "South", Count: 3 },
        ],
      },
      { sheetName: "Inventory" },
    );
    const databasePath = path.join(directory, `inventory.${format}`);
    const { database } = await createDatabase({ path: databasePath, format });
    try {
      const bytes = await openRandomAccessSource(input);
      const scratch = await createScratchDirectory(directory);
      const workbook = await createWorkbookImportSource({
        key: "inventory",
        bytes,
        scratch,
        verifyUnchanged: () => bytes.verifyUnchanged(),
      });
      const profile = await draftImportProfile({ sources: [workbook.source] });
      const planPath = path.join(directory, "review.ccplan");
      const originalPlan = await createImportBatch({
        path: planPath,
        database,
        profile,
        baselineRevision: 0n,
        protectedInputPaths: [input],
      });
      try {
        const captured = await prepareImport({
          database,
          prepared: originalPlan,
          profile,
          sources: [workbook.source],
        });
        expect(captured.result.metrics.rowsCaptured).toBe(2);
      } finally {
        await originalPlan.close();
        await workbook.close();
        await bytes.close();
        await scratch.close();
      }
      await rm(input);
      const prepared = await openImportBatch({ path: planPath });
      try {
        const preview = await inspectImport({ prepared, page: { limit: 1 } });
        expect(preview.examples).toHaveLength(1);
        expect(preview.nextCursor).toBeDefined();
        const second = await inspectImport({
          prepared,
          page: { limit: 1, cursor: preview.nextCursor },
        });
        expect(second.examples[0]?.sourceRow).toBe(3);
        const approved = await resolveImport({
          database,
          prepared,
          decisions: [],
        });
        if (approved.state !== "ready")
          throw new Error("Unresolved synthetic plan.");
        const applied = await applyImport({
          database,
          prepared,
          approved,
          requestId: "import-a",
        });
        const context = {
          label: "Synthetic partial delivery",
          scope: { kind: "partial" as const, description: "Selected records" },
          attributes: { reportedCount: 8 },
        };
        const first = await recordBatch({
          database,
          captureIds: applied.captureIds,
          context,
          requestId: "delivery-a",
        });
        const retried = await recordBatch({
          database,
          captureIds: [...applied.captureIds, ...applied.captureIds],
          context: {
            attributes: { reportedCount: 8 },
            scope: context.scope,
            label: context.label,
          },
          requestId: "delivery-a",
        });
        expect(first.batch.reusedCaptureIds).toEqual([]);
        expect(retried.batch).toEqual(first.batch);
        expect(retried.metrics.batchesRecorded).toBe(0);
        const repeated = await recordBatch({
          database,
          captureIds: applied.captureIds,
          context,
          requestId: "delivery-b",
        });
        expect(repeated.batch.reusedCaptureIds).toEqual(applied.captureIds);
        const firstPage = await listBatches({ database, limit: 1 });
        const secondPage = await listBatches({
          database,
          limit: 1,
          cursor: firstPage.nextCursor,
        });
        expect(firstPage.batches[0]?.id).toBe(first.batch.id);
        expect(firstPage.batches[0]?.reusedCaptureIds).toEqual([]);
        expect(secondPage.batches[0]?.requestId).toBe("delivery-b");
        expect(secondPage.batches[0]?.reusedCaptureIds).toEqual(
          applied.captureIds,
        );
        expect(secondPage.nextCursor).toBeUndefined();
        expect(
          (
            await recordBatch({
              database,
              captureIds: applied.captureIds,
              context,
              requestId: "delivery-a",
            })
          ).batch.reusedCaptureIds,
        ).toEqual([]);
        await engineOf(database).execute(
          `UPDATE ${COUNTERS_TABLE} SET next_value = ? WHERE counter_name = ?`,
          [999_999n, "delivery"],
        );
        const lastPadded = await recordBatch({
          database,
          captureIds: applied.captureIds,
          context,
          requestId: "delivery-999999",
        });
        const firstUnpadded = await recordBatch({
          database,
          captureIds: applied.captureIds,
          context,
          requestId: "delivery-1000000",
        });
        expect(lastPadded.batch.id).toBe("DEL-999999");
        expect(firstUnpadded.batch.id).toBe("DEL-1000000");
        const naturalFirstPage = await listBatches({
          database,
          limit: 3,
        });
        expect(naturalFirstPage.batches.map((delivery) => delivery.id)).toEqual(
          ["DEL-000001", "DEL-000002", "DEL-999999"],
        );
        const naturalSecondPage = await listBatches({
          database,
          limit: 1,
          cursor: naturalFirstPage.nextCursor,
        });
        expect(
          naturalSecondPage.batches.map((delivery) => delivery.id),
        ).toEqual(["DEL-1000000"]);
        expect(naturalSecondPage.batches[0]?.reusedCaptureIds).toEqual(
          applied.captureIds,
        );
        await expect(
          recordBatch({
            database,
            captureIds: applied.captureIds,
            context: { ...context, label: "Changed" },
            requestId: "delivery-a",
          }),
        ).rejects.toMatchObject({ code: "DB_REQUEST_ID_CONFLICT" });
        await expect(
          recordBatch({
            database,
            captureIds: ["CAP-missing"],
            context,
            requestId: "delivery-missing",
          }),
        ).rejects.toMatchObject({ code: "DB_CAPTURE_NOT_FOUND" });
        await expect(
          recordBatch({
            database,
            captureIds: [],
            context,
            requestId: "delivery-empty",
          }),
        ).rejects.toMatchObject({ code: "DB_DELIVERY_CAPTURE_REQUIRED" });
        await expect(
          recordBatch({
            database,
            captureIds: applied.captureIds,
            context,
            requestId: " ",
          }),
        ).rejects.toMatchObject({ code: "DB_DELIVERY_REQUEST_ID_REQUIRED" });
        await expect(listBatches({ database, limit: 0 })).rejects.toMatchObject(
          { code: "DB_INVALID_PAGE_SIZE" },
        );
        await expect(
          listBatches({ database, limit: 1, cursor: "DEL-1junk" }),
        ).rejects.toMatchObject({ code: "DB_INVALID_CURSOR" });
        const inspection = await inspectDatabase({ database });
        expect(inspection.tables[0]?.rowCount).toBe(2n);
        expect(inspection.recordedBatches).toBe(4n);
        await expect(
          createImportBatch({
            path: databasePath,
            database,
            profile,
            baselineRevision: inspection.revision,
            overwrite: true,
          }),
        ).rejects.toMatchObject({ code: "FILES_INPUT_OVERWRITE" });
        const output = path.join(directory, `backup.${format}`);
        await exportDatabase({ database, output });
        const reopened = await openDatabase({ path: output });
        try {
          expect(
            (await inspectDatabase({ database: reopened })).recordedBatches,
          ).toBe(4n);
        } finally {
          await reopened.close();
        }
        await expect(
          exportDatabase({ database, output }),
        ).rejects.toMatchObject({ code: "FILES_OUTPUT_EXISTS" });
        await expect(
          exportDatabase({ database, output: databasePath, overwrite: true }),
        ).rejects.toMatchObject({ code: "FILES_INPUT_OVERWRITE" });
      } finally {
        await prepared.close();
      }
    } finally {
      await database.close();
    }
  });
}
