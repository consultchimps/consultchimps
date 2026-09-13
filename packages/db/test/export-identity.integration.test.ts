import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { inspectDatabase, type Database } from "../src/database.js";
import {
  applyImport,
  prepareImport,
  resolveImport,
} from "../src/import/operations.js";
import { draftImportRecipe } from "../src/import/recipe.js";
import type { ImportSource } from "../src/import/types.js";
import {
  createDatabase,
  createPreparedImport,
  exportDatabase,
  openDatabase,
  openPreparedImport,
} from "../src/node.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

function source(): ImportSource {
  const bytes = new TextEncoder().encode("synthetic snapshot identity source");
  return {
    key: "Inventory",
    readerVersion: "snapshot-identity-1",
    bytes: {
      name: "inventory.xlsx",
      size: bytes.length,
      async readAt(offset, length) {
        return bytes.slice(offset, offset + length);
      },
    },
    selections: [
      {
        key: "inventory",
        label: "Inventory",
        async open() {
          return {
            columns: ["Name"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: {
                    Name: { kind: "string" as const, value: "North" },
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
  test(`${format}: same-format snapshots retain plan identity while conversions require review`, async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), `cc-export-identity-${format}-`),
    );
    directories.push(directory);
    const databasePath = path.join(directory, `source.${format}`);
    const sameFormatPath = path.join(directory, `snapshot.${format}`);
    const convertedFormat = format === "sqlite" ? "duckdb" : "sqlite";
    const convertedPath = path.join(directory, `converted.${convertedFormat}`);
    const { database: sourceDatabase } = await createDatabase({
      path: databasePath,
      format,
    });
    const importSource = source();
    const recipe = await draftImportRecipe({ sources: [importSource] });
    const sourceBefore = await inspectDatabase({ database: sourceDatabase });
    const planPath = path.join(directory, "review.ccplan");
    let prepared = await createPreparedImport({
      path: planPath,
      database: sourceDatabase,
      recipe,
      baselineRevision: sourceBefore.revision,
    });
    let snapshot: Database | undefined;
    let converted: Database | undefined;
    try {
      await prepareImport({
        database: sourceDatabase,
        prepared,
        sources: [importSource],
        recipe,
      });
      const approved = await resolveImport({
        database: sourceDatabase,
        prepared,
        decisions: [],
      });
      expect(approved.state).toBe("ready");
      if (approved.state !== "ready") throw new Error("Plan failed review");

      const sameFormatExport = await exportDatabase({
        database: sourceDatabase,
        output: sameFormatPath,
      });
      expect(sameFormatExport.plan).toMatchObject({
        databaseId: sourceDatabase.id,
        baselineRevision: sourceBefore.revision,
        sourceFormat: format,
        targetFormat: format,
      });
      const conversionExport = await exportDatabase({
        database: sourceDatabase,
        output: convertedPath,
        format: convertedFormat,
      });
      expect(conversionExport.plan).toMatchObject({
        databaseId: sourceDatabase.id,
        baselineRevision: sourceBefore.revision,
        sourceFormat: format,
        targetFormat: convertedFormat,
      });
      await prepared.close();
      prepared = await openPreparedImport({ path: planPath });
      expect(prepared.databaseId).toBe(sourceDatabase.id);

      snapshot = await openDatabase({ path: sameFormatPath });
      converted = await openDatabase({ path: convertedPath });
      const snapshotBefore = await inspectDatabase({ database: snapshot });
      const convertedBefore = await inspectDatabase({ database: converted });
      expect(snapshotBefore).toEqual(sourceBefore);
      expect(convertedBefore.id).not.toBe(sourceDatabase.id);
      expect(convertedBefore).toMatchObject({
        format: convertedFormat,
        revision: sourceBefore.revision,
        tables: [],
        captures: 0n,
        completedImports: 0n,
      });

      await expect(
        applyImport({
          database: converted,
          prepared,
          approved,
          requestId: `${format}-converted-apply`,
        }),
      ).rejects.toMatchObject({ code: "DB_STALE_IMPORT_PLAN" });
      expect(await inspectDatabase({ database: converted })).toEqual(
        convertedBefore,
      );

      const applied = await applyImport({
        database: snapshot,
        prepared,
        approved,
        requestId: `${format}-snapshot-apply`,
      });
      expect(applied.metrics.rowsImported).toBe(1);
      const snapshotAfter = await inspectDatabase({ database: snapshot });
      expect(snapshotAfter).toMatchObject({
        id: sourceDatabase.id,
        format,
        revision: sourceBefore.revision + 1n,
        completedImports: 1n,
        tables: [{ name: "Inventory", rowCount: 1n }],
      });
      expect(await inspectDatabase({ database: sourceDatabase })).toEqual(
        sourceBefore,
      );
    } finally {
      await converted?.close();
      await snapshot?.close();
      await prepared.close();
      await sourceDatabase.close();
    }
  });
}
