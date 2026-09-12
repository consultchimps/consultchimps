import { expect, test } from "@playwright/test";

import {
  applyImport,
  inspectDatabase,
  listDeliveries,
  prepareImport,
  recordDelivery,
  resolveImport,
  type ImportRecipe,
  type ImportSource,
} from "@consultchimps/db";
import { createDatabase, createPreparedImport } from "@consultchimps/db/node";

test("shows changes baselines and partial descriptions from an opened database", async ({
  page,
}, testInfo) => {
  const databasePath = testInfo.outputPath("delivery-scopes.sqlite");
  const planPath = testInfo.outputPath("delivery-scopes.ccplan");
  const mixedPlanPath = testInfo.outputPath("delivery-mixed.ccplan");
  const { database } = await createDatabase({
    path: databasePath,
    format: "sqlite",
  });
  const bytes = new TextEncoder().encode("synthetic delivery scope source");
  const source: ImportSource = {
    key: "synthetic-source",
    readerVersion: "synthetic-1",
    bytes: {
      name: "synthetic.xlsx",
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
                    Value: { kind: "string" as const, value: "Synthetic" },
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
  const recipe: ImportRecipe = {
    version: 1,
    routes: [
      {
        source: source.key,
        selection: "Inventory",
        destination: {
          kind: "new-table",
          schema: {
            name: "Inventory",
            recordId: { prefix: "INV", padding: 6 },
            columns: [{ name: "Value", type: "text" }],
            foreignKeys: [],
          },
        },
        columns: [{ source: "Value", target: "Value", type: "text" }],
      },
    ],
  };
  const prepared = await createPreparedImport({
    path: planPath,
    database,
    recipe,
    baselineRevision: 0n,
  });
  try {
    await prepareImport({ database, prepared, sources: [source], recipe });
    const approved = await resolveImport({
      database,
      prepared,
      decisions: [],
    });
    if (approved.state !== "ready") {
      throw new Error("Synthetic delivery import did not become ready");
    }
    const applied = await applyImport({
      database,
      prepared,
      approved,
      requestId: "changes-delivery",
      delivery: {
        label: "Changes delivery",
        scope: { kind: "changes", baseline: "DEL-BASELINE-42" },
        attributes: {
          vendor: "Vendor A",
          entity: "Entity North",
          phase: "Changes",
        },
      },
    });
    await recordDelivery({
      database,
      captureIds: applied.captureIds,
      requestId: "partial-delivery",
      context: {
        label: "Partial delivery",
        scope: {
          kind: "partial",
          description: "Selected business units",
        },
        attributes: {
          vendor: "Vendor B",
          entity: "Entity South",
          phase: "Partial",
        },
      },
    });
    await recordDelivery({
      database,
      captureIds: applied.captureIds,
      requestId: "label-only-delivery",
      context: {
        label: "September client handoff",
        scope: { kind: "full" },
      },
    });
    const changedBytes = new TextEncoder().encode(
      "synthetic changed delivery source",
    );
    const changedSource: ImportSource = {
      key: "changed-source",
      readerVersion: "synthetic-1",
      bytes: {
        name: "changed.xlsx",
        size: changedBytes.length,
        async readAt(offset, length) {
          return changedBytes.slice(offset, offset + length);
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
                      Value: { kind: "string" as const, value: "Changed" },
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
    const existingRoute = (importSource: ImportSource) => ({
      source: importSource.key,
      selection: "Inventory",
      destination: { kind: "existing-table" as const, table: "Inventory" },
      columns: [{ source: "Value", target: "Value", type: "text" as const }],
    });
    const mixedRecipe: ImportRecipe = {
      version: 1,
      routes: [existingRoute(source), existingRoute(changedSource)],
    };
    const mixedPrepared = await createPreparedImport({
      path: mixedPlanPath,
      database,
      recipe: mixedRecipe,
      baselineRevision: (await inspectDatabase({ database })).revision,
    });
    try {
      await prepareImport({
        database,
        prepared: mixedPrepared,
        sources: [source, changedSource],
        recipe: mixedRecipe,
      });
      const mixedApproved = await resolveImport({
        database,
        prepared: mixedPrepared,
        decisions: [],
      });
      if (mixedApproved.state !== "ready") {
        throw new Error("Synthetic mixed delivery import did not become ready");
      }
      const mixed = await applyImport({
        database,
        prepared: mixedPrepared,
        approved: mixedApproved,
        requestId: "mixed-delivery",
        delivery: {
          label: "Mixed delivery",
          scope: { kind: "full" },
        },
      });
      expect(mixed.metrics).toMatchObject({ rowsImported: 1, rowsReused: 1 });
      const history = await listDeliveries({ database, limit: 10 });
      expect(history.deliveries.at(-1)).toMatchObject({
        context: { label: "Mixed delivery" },
        captureIds: expect.arrayContaining([...mixed.captureIds]),
        reusedCaptureIds: [applied.captureIds[0]],
      });
      expect(history.deliveries.at(-1)?.captureIds).toHaveLength(2);
    } finally {
      await mixedPrepared.close();
    }
  } finally {
    await prepared.close();
    await database.close();
  }

  await page.goto("/workspace");
  await page.getByTestId("workspace-open-input").setInputFiles(databasePath);
  await expect(page.getByTestId("workspace-summary")).toBeVisible();
  await page.getByTestId("workspace-deliveries-refresh").click();
  const deliveries = page.getByTestId("workspace-delivery");
  await expect(deliveries).toHaveCount(4);
  await expect(deliveries.nth(0)).toContainText("Changes delivery");
  await expect(deliveries.nth(0)).toContainText("Vendor A");
  await expect(deliveries.nth(0)).toContainText(
    "Changes since DEL-BASELINE-42",
  );
  await expect(deliveries.nth(1)).toContainText("Partial delivery");
  await expect(deliveries.nth(1)).toContainText("Vendor B");
  await expect(deliveries.nth(1)).toContainText(
    "Partial coverage: Selected business units",
  );
  await expect(deliveries.nth(2)).toContainText("September client handoff");
  await expect(deliveries.nth(2)).toContainText("Unspecified vendor");
  await expect(deliveries.nth(2)).toContainText("Full coverage");
  await expect(deliveries.nth(3)).toContainText("Mixed delivery");
  await expect(deliveries.nth(3)).toContainText(
    "Includes previously captured data",
  );
  await expect(deliveries.nth(3)).not.toContainText(
    "without adding observation rows",
  );
  const screenshotPath = testInfo.outputPath("delivery-history-labels.png");
  await page
    .getByTestId("workspace-deliveries")
    .screenshot({ path: screenshotPath });
  await testInfo.attach("delivery-history-labels", {
    path: screenshotPath,
    contentType: "image/png",
  });
});
