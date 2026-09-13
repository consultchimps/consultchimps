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
import type { ImportRecipe, ImportSource } from "../src/import/types.js";
import { APPLICATION_TABLE, IMPORT_REQUEST_TABLE } from "../src/metadata.js";
import { createDatabase, createPreparedImport } from "../src/node.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

const recipe: ImportRecipe = {
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
    const prepared = await createPreparedImport({
      path: path.join(directory, "review.ccplan"),
      database,
      recipe,
      baselineRevision: 0n,
    });
    try {
      const preparedResult = await prepareImport({
        database,
        prepared,
        sources: [importSource()],
        recipe,
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
            delivery: {
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
        delivery: {
          label: "Synthetic delivery",
          scope: { kind: "full" },
        },
      });
      expect(applied.metrics).toMatchObject({
        rowsImported: 1,
        deliveriesRecorded: 1,
      });
      const retried = await applyImport({
        database,
        prepared,
        approved,
        requestId: `valid-${format}`,
        delivery: {
          label: "Synthetic delivery",
          scope: { kind: "full" },
        },
      });
      expect(retried.importIds).toEqual(applied.importIds);
      expect(retried.captureIds).toEqual(applied.captureIds);
      expect(retried.metrics).toMatchObject({
        rowsImported: 0,
        rowsReused: 1,
        deliveriesRecorded: 0,
      });
      const importId = applied.importIds[0];
      if (importId === undefined) throw new Error("Import ID was not returned");
      expect(await inspectDatabase({ database })).toMatchObject({
        revision: 1n,
        tables: [{ name: "Inventory", rowCount: 1n }],
        captures: 1n,
        completedImports: 1n,
        deliveries: 1n,
      });
      await expect(
        engineOf(database).query(
          `SELECT count(*) AS count FROM ${IMPORT_REQUEST_TABLE}`,
        ),
      ).resolves.toEqual([{ count: 1n }]);

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
            delivery: {
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
          delivery: {
            label: "Synthetic delivery",
            scope: { kind: "full" },
          },
        }),
      ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });
      expect(await inspectDatabase({ database })).toEqual(
        beforeMissingApplication,
      );
    } finally {
      await prepared.close();
      await database.close();
    }
  });
}
