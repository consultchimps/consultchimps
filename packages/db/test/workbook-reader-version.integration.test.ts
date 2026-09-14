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
import {
  applyImport,
  inspectImport,
  prepareImport,
} from "../src/import/operations.js";
import type { ImportProfile, ImportSource } from "../src/import/types.js";
import { createDatabase, createImportBatch } from "../src/node.js";
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
  test(`${format}: changed workbook readers recapture once before reuse`, async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "cc-workbook-reader-version-"),
    );
    directories.push(directory);
    const workbookPath = path.join(directory, "inventory.xlsx");
    await writeTable(
      workbookPath,
      {
        columns: ["Name", "Count"],
        rows: [
          { Name: "North", Count: 2 },
          { Name: "South", Count: 3 },
        ],
      },
      { sheetName: "Inventory" },
    );

    const bytes = await openRandomAccessSource(workbookPath);
    const scratch = await createScratchDirectory(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    try {
      const legacyWorkbook = await createWorkbookImportSource({
        key: "inventory",
        bytes,
        scratch,
        verifyUnchanged: () => bytes.verifyUnchanged(),
      });
      const selection = legacyWorkbook.source.selections[0];
      if (selection === undefined) throw new Error("Workbook has no selection");
      const legacySource: ImportSource = {
        ...legacyWorkbook.source,
        readerVersion: "consultchimps-xlsx-stream-1",
      };
      const createRecipe: ImportProfile = {
        version: 1,
        routes: [
          {
            source: "inventory",
            selection: selection.key,
            destination: {
              kind: "new-table",
              schema: {
                name: "InventoryRecords",
                columns: [
                  { name: "Name", type: "text" },
                  { name: "Count", type: "integer" },
                ],
                recordId: { prefix: "INV", padding: 4 },
                foreignKeys: [],
              },
            },
            columns: [
              { source: "Name", target: "Name", type: "text" },
              { source: "Count", target: "Count", type: "integer" },
            ],
          },
        ],
      };
      const reuseRecipe: ImportProfile = {
        version: 1,
        routes: [
          {
            ...createRecipe.routes[0]!,
            destination: { kind: "existing-table", table: "InventoryRecords" },
          },
        ],
      };
      const legacyPlan = await createImportBatch({
        path: path.join(directory, "legacy.ccplan"),
        database,
        profile: createRecipe,
        baselineRevision: (await inspectDatabase({ database })).revision,
        protectedInputPaths: [workbookPath],
      });
      try {
        const legacy = await prepareImport({
          database,
          prepared: legacyPlan,
          profile: createRecipe,
          sources: [legacySource],
        });
        expect(legacy.result.metrics).toMatchObject({
          sourcesRead: 1,
          sourcesReused: 0,
          rowsCaptured: 2,
        });
        if (legacy.prepared.state !== "ready") {
          throw new Error("Legacy workbook import was not ready");
        }
        await applyImport({
          database,
          prepared: legacyPlan,
          approved: legacy.prepared,
          requestId: `${format}-legacy-reader`,
        });
      } finally {
        await legacyPlan.close();
        await legacyWorkbook.close();
      }

      const currentWorkbook = await createWorkbookImportSource({
        key: "inventory",
        bytes,
        scratch,
        verifyUnchanged: () => bytes.verifyUnchanged(),
      });
      const currentPlan = await createImportBatch({
        path: path.join(directory, "current.ccplan"),
        database,
        profile: reuseRecipe,
        baselineRevision: (await inspectDatabase({ database })).revision,
        protectedInputPaths: [workbookPath],
      });
      try {
        const current = await prepareImport({
          database,
          prepared: currentPlan,
          profile: reuseRecipe,
          sources: [currentWorkbook.source],
        });
        expect(current.result.metrics).toMatchObject({
          sourcesRead: 1,
          sourcesReused: 0,
          rowsCaptured: 2,
        });
        expect(
          (
            await inspectImport({
              prepared: currentPlan,
              page: { limit: 1 },
            })
          ).routes[0],
        ).toMatchObject({ reused: false, rowCount: 2n });
        expect(await inspectDatabase({ database })).toMatchObject({
          captures: 1n,
          completedImports: 1n,
          tables: [{ name: "InventoryRecords", rowCount: 2n }],
        });
        if (current.prepared.state !== "ready") {
          throw new Error("Current workbook import was not ready");
        }
        await applyImport({
          database,
          prepared: currentPlan,
          approved: current.prepared,
          requestId: `${format}-current-reader`,
        });
      } finally {
        await currentPlan.close();
        await currentWorkbook.close();
      }

      const repeatedWorkbook = await createWorkbookImportSource({
        key: "inventory",
        bytes,
        scratch,
        verifyUnchanged: () => bytes.verifyUnchanged(),
      });
      const repeatedPlan = await createImportBatch({
        path: path.join(directory, "repeated.ccplan"),
        database,
        profile: reuseRecipe,
        baselineRevision: (await inspectDatabase({ database })).revision,
        protectedInputPaths: [workbookPath],
      });
      try {
        const repeated = await prepareImport({
          database,
          prepared: repeatedPlan,
          profile: reuseRecipe,
          sources: [repeatedWorkbook.source],
        });
        expect(repeated.result.metrics).toMatchObject({
          sourcesRead: 0,
          sourcesReused: 1,
          rowsCaptured: 0,
        });
        expect(
          (
            await inspectImport({
              prepared: repeatedPlan,
              page: { limit: 1 },
            })
          ).routes[0],
        ).toMatchObject({ reused: true, rowCount: 2n });
        expect(await inspectDatabase({ database })).toMatchObject({
          captures: 2n,
          completedImports: 2n,
          tables: [{ name: "InventoryRecords", rowCount: 4n }],
        });
      } finally {
        await repeatedPlan.close();
        await repeatedWorkbook.close();
      }
    } finally {
      await database.close();
      await bytes.close();
      await scratch.close();
    }
  });
}
