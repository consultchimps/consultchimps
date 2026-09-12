import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  createScratchDirectory,
  openRandomAccessSource,
} from "@consultchimps/files";
import { writeTable } from "@consultchimps/xlsx";
import { afterEach, expect, test } from "vitest";

import { inspectDatabase } from "../src/database.js";
import { recordDelivery, listDeliveries } from "../src/import/deliveries.js";
import {
  prepareImport,
  inspectImport,
  resolveImport,
  applyImport,
} from "../src/import/operations.js";
import { draftImportRecipe } from "../src/import/recipe.js";
import {
  createDatabase,
  createPreparedImport,
  openPreparedImport,
  exportDatabase,
  openDatabase,
} from "../src/node.js";
import { createWorkbookImportSource } from "../src/workbook.js";

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
      const recipe = await draftImportRecipe({ sources: [workbook.source] });
      const planPath = path.join(directory, "review.ccplan");
      const originalPlan = await createPreparedImport({
        path: planPath,
        database,
        recipe,
        baselineRevision: 0n,
        protectedInputPaths: [input],
      });
      try {
        const captured = await prepareImport({
          database,
          prepared: originalPlan,
          recipe,
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
      const prepared = await openPreparedImport({ path: planPath });
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
        const first = await recordDelivery({
          database,
          captureIds: applied.captureIds,
          context,
          requestId: "delivery-a",
        });
        const retried = await recordDelivery({
          database,
          captureIds: [...applied.captureIds, ...applied.captureIds],
          context: {
            attributes: { reportedCount: 8 },
            scope: context.scope,
            label: context.label,
          },
          requestId: "delivery-a",
        });
        expect(retried.delivery).toEqual(first.delivery);
        expect(retried.metrics.deliveriesRecorded).toBe(0);
        await recordDelivery({
          database,
          captureIds: applied.captureIds,
          context,
          requestId: "delivery-b",
        });
        const firstPage = await listDeliveries({ database, limit: 1 });
        const secondPage = await listDeliveries({
          database,
          limit: 1,
          cursor: firstPage.nextCursor,
        });
        expect(firstPage.deliveries[0]?.id).toBe(first.delivery.id);
        expect(secondPage.deliveries[0]?.requestId).toBe("delivery-b");
        expect(secondPage.nextCursor).toBeUndefined();
        await expect(
          recordDelivery({
            database,
            captureIds: applied.captureIds,
            context: { ...context, label: "Changed" },
            requestId: "delivery-a",
          }),
        ).rejects.toMatchObject({ code: "DB_REQUEST_ID_CONFLICT" });
        await expect(
          recordDelivery({
            database,
            captureIds: ["CAP-missing"],
            context,
            requestId: "delivery-missing",
          }),
        ).rejects.toMatchObject({ code: "DB_CAPTURE_NOT_FOUND" });
        await expect(
          recordDelivery({
            database,
            captureIds: [],
            context,
            requestId: "delivery-empty",
          }),
        ).rejects.toMatchObject({ code: "DB_DELIVERY_CAPTURE_REQUIRED" });
        await expect(
          recordDelivery({
            database,
            captureIds: applied.captureIds,
            context,
            requestId: " ",
          }),
        ).rejects.toMatchObject({ code: "DB_DELIVERY_REQUEST_ID_REQUIRED" });
        await expect(
          listDeliveries({ database, limit: 0 }),
        ).rejects.toMatchObject({ code: "DB_INVALID_PAGE_SIZE" });
        await expect(
          listDeliveries({ database, limit: 1, cursor: "DEL-1junk" }),
        ).rejects.toMatchObject({ code: "DB_INVALID_CURSOR" });
        const inspection = await inspectDatabase({ database });
        expect(inspection.tables[0]?.rowCount).toBe(2n);
        expect(inspection.deliveries).toBe(2n);
        await expect(
          createPreparedImport({
            path: databasePath,
            database,
            recipe,
            baselineRevision: inspection.revision,
            overwrite: true,
          }),
        ).rejects.toMatchObject({ code: "FILES_INPUT_OVERWRITE" });
        const output = path.join(directory, `backup.${format}`);
        await exportDatabase({ database, output });
        const reopened = await openDatabase({ path: output });
        try {
          expect(
            (await inspectDatabase({ database: reopened })).deliveries,
          ).toBe(2n);
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
