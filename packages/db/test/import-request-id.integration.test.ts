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
import { APPLICATION_TABLE, IMPORT_REQUEST_TABLE } from "../src/metadata.js";
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
          foreignKeys: [],
        },
      },
      columns: [{ source: "Value", target: "Value", type: "text" }],
    },
  ],
};

function importSource(): ImportSource {
  const bytes = new TextEncoder().encode("synthetic request ID fixture");
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
  test(`${format}: request IDs and stored receipts are validated without duplicate writes`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-request-id-"));
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
    try {
      const preparedResult = await prepareImport({
        database,
        prepared,
        sources: [importSource()],
        profile,
      });
      expect(preparedResult.prepared.state).toBe("ready");
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [],
      });
      if (approved.state !== "ready") throw new Error("Plan failed review");
      const before = await inspectDatabase({ database });

      for (const requestId of ["", " \t\n"]) {
        await expect(
          applyImport({
            database,
            prepared,
            approved,
            requestId,
            batchContext: {
              label: "Synthetic delivery",
              scope: { kind: "full" },
            },
          }),
        ).rejects.toMatchObject({ code: "DB_IMPORT_REQUEST_ID_REQUIRED" });
        expect(await inspectDatabase({ database })).toEqual(before);
        await expect(
          engineOf(database).query(
            `SELECT count(*) AS count FROM ${IMPORT_REQUEST_TABLE}`,
          ),
        ).resolves.toEqual([{ count: 0n }]);
      }

      const applied = await applyImport({
        database,
        prepared,
        approved,
        requestId: `valid-${format}`,
        batchContext: {
          label: "Synthetic delivery",
          scope: { kind: "full" },
        },
      });
      expect(applied.metrics).toMatchObject({
        rowsImported: 1,
        batchesRecorded: 1,
      });
      const retried = await applyImport({
        database,
        prepared,
        approved,
        requestId: `valid-${format}`,
        batchContext: {
          label: "Synthetic delivery",
          scope: { kind: "full" },
        },
      });
      expect(retried.importIds).toEqual(applied.importIds);
      expect(retried.captureIds).toEqual(applied.captureIds);
      expect(retried.metrics).toMatchObject({
        rowsImported: 0,
        rowsReused: 1,
        batchesRecorded: 0,
      });
      const importId = applied.importIds[0];
      if (importId === undefined) throw new Error("Import ID was not returned");
      expect(await inspectDatabase({ database })).toMatchObject({
        revision: 1n,
        tables: [{ name: "Inventory", rowCount: 1n }],
        captures: 1n,
        completedImports: 1n,
        recordedBatches: 1n,
      });
      await expect(
        engineOf(database).query(
          `SELECT count(*) AS count FROM ${IMPORT_REQUEST_TABLE}`,
        ),
      ).resolves.toEqual([{ count: 1n }]);

      await engineOf(database).execute(
        `UPDATE ${IMPORT_REQUEST_TABLE} SET capture_ids_json = ? WHERE request_id = ?`,
        [JSON.stringify(["CAP-wrong"]), `valid-${format}`],
      );
      const beforeWrongCapture = await inspectDatabase({ database });
      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: `valid-${format}`,
          batchContext: {
            label: "Synthetic delivery",
            scope: { kind: "full" },
          },
        }),
      ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });
      expect(await inspectDatabase({ database })).toEqual(beforeWrongCapture);
      await engineOf(database).execute(
        `UPDATE ${IMPORT_REQUEST_TABLE} SET capture_ids_json = ? WHERE request_id = ?`,
        [JSON.stringify(applied.captureIds), `valid-${format}`],
      );

      for (const corruption of [
        { name: "negative receipt", receipt: -1n, application: 1n },
        { name: "receipt mismatch", receipt: 2n, application: 1n },
        { name: "application mismatch", receipt: 2n, application: 2n },
      ] as const) {
        await engineOf(database).execute(
          `UPDATE ${IMPORT_REQUEST_TABLE} SET row_count = ? WHERE request_id = ?`,
          [corruption.receipt, `valid-${format}`],
        );
        await engineOf(database).execute(
          `UPDATE ${APPLICATION_TABLE} SET row_count = ? WHERE import_id = ?`,
          [corruption.application, importId],
        );
        const beforeCorruptRetry = await inspectDatabase({ database });

        await expect(
          applyImport({
            database,
            prepared,
            approved,
            requestId: `valid-${format}`,
            batchContext: {
              label: "Synthetic delivery",
              scope: { kind: "full" },
            },
          }),
          corruption.name,
        ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });
        expect(await inspectDatabase({ database })).toEqual(beforeCorruptRetry);
        await expect(
          engineOf(database).query(
            `SELECT count(*) AS count FROM ${IMPORT_REQUEST_TABLE}`,
          ),
        ).resolves.toEqual([{ count: 1n }]);
      }
      await engineOf(database).execute(
        `UPDATE ${APPLICATION_TABLE} SET row_count = 1 WHERE import_id = ?`,
        [importId],
      );
      await engineOf(database).execute(
        `UPDATE ${IMPORT_REQUEST_TABLE} SET import_ids_json = ?, row_count = 1 WHERE request_id = ?`,
        [JSON.stringify(["IMP-missing"]), `valid-${format}`],
      );
      const beforeMissingApplication = await inspectDatabase({ database });
      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: `valid-${format}`,
          batchContext: {
            label: "Synthetic delivery",
            scope: { kind: "full" },
          },
        }),
      ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });
      expect(await inspectDatabase({ database })).toEqual(
        beforeMissingApplication,
      );

      const originalRoute = profile.routes[0];
      if (
        originalRoute === undefined ||
        originalRoute.destination.kind !== "new-table"
      ) {
        throw new Error("Expected a new table route");
      }
      const archiveRecipe: ImportProfile = {
        version: 1,
        routes: [
          {
            ...originalRoute,
            destination: {
              kind: "new-table",
              schema: {
                ...originalRoute.destination.schema,
                name: "Archive",
                recordId: { prefix: "ARCHIVE", padding: 4 },
              },
            },
          },
        ],
      };
      const archivePlan = await createImportBatch({
        path: path.join(directory, "archive.ccplan"),
        database,
        profile: archiveRecipe,
        baselineRevision: (await inspectDatabase({ database })).revision,
      });
      try {
        const archivePrepared = await prepareImport({
          database,
          prepared: archivePlan,
          sources: [importSource()],
          profile: archiveRecipe,
        });
        expect(archivePrepared.prepared.state).toBe("ready");
        const archiveApproved = await resolveImport({
          database,
          prepared: archivePlan,
          decisions: [],
        });
        if (archiveApproved.state !== "ready") {
          throw new Error("Archive plan failed review");
        }
        const archiveApplied = await applyImport({
          database,
          prepared: archivePlan,
          approved: archiveApproved,
          requestId: `archive-${format}`,
        });
        const archiveImportId = archiveApplied.importIds[0];
        if (archiveImportId === undefined) {
          throw new Error("Archive import ID was not returned");
        }
        const applicationRows = await engineOf(database).query(
          `SELECT import_id, capture_id, table_name, row_count FROM ${APPLICATION_TABLE} WHERE import_id IN (?, ?) ORDER BY table_name`,
          [importId, archiveImportId],
        );
        expect(applicationRows).toEqual([
          {
            import_id: archiveImportId,
            capture_id: applied.captureIds[0],
            table_name: "Archive",
            row_count: 1n,
          },
          {
            import_id: importId,
            capture_id: applied.captureIds[0],
            table_name: "Inventory",
            row_count: 1n,
          },
        ]);
        await engineOf(database).execute(
          `UPDATE ${IMPORT_REQUEST_TABLE} SET import_ids_json = ?, row_count = 1 WHERE request_id = ?`,
          [JSON.stringify([archiveImportId]), `valid-${format}`],
        );
        const beforeSwappedApplication = await inspectDatabase({ database });

        await expect(
          applyImport({
            database,
            prepared,
            approved,
            requestId: `valid-${format}`,
            batchContext: {
              label: "Synthetic delivery",
              scope: { kind: "full" },
            },
          }),
        ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });
        expect(await inspectDatabase({ database })).toEqual(
          beforeSwappedApplication,
        );
      } finally {
        await archivePlan.close();
      }
    } finally {
      await prepared.close();
      await database.close();
    }
  });
}
