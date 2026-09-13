import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { sha256 } from "@noble/hashes/sha2.js";
import { bytesToHex } from "@noble/hashes/utils.js";
import { afterEach, expect, test } from "vitest";

import { engineOf, inspectDatabase } from "../src/database.js";
import { canonicalJson } from "../src/internal/json.js";
import {
  applyImport,
  inspectImport,
  prepareImport,
  resolveImport,
} from "../src/import/operations.js";
import type { ImportRecipe, ImportSource } from "../src/import/types.js";
import {
  APPLICATION_TABLE,
  IMPORT_REQUEST_TABLE,
  PLAN_TABLE,
} from "../src/metadata.js";
import { createDatabase, createPreparedImport } from "../src/node.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

function source(): ImportSource {
  const bytes = new TextEncoder().encode("synthetic application identity");
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
        key: "Data",
        label: "Data",
        async open() {
          return {
            columns: ["A", "B"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: {
                    A: { kind: "string" as const, value: "North" },
                    B: { kind: "string" as const, value: "South" },
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

function recipe(options: {
  readonly table: string;
  readonly kind: "new-table" | "existing-table";
  readonly swapped?: boolean | undefined;
  readonly reversed?: boolean | undefined;
}): ImportRecipe {
  const columns = options.swapped
    ? [
        { source: "A", target: "Second", type: "text" as const },
        { source: "B", target: "First", type: "text" as const },
      ]
    : [
        { source: "A", target: "First", type: "text" as const },
        { source: "B", target: "Second", type: "text" as const },
      ];
  if (options.reversed) columns.reverse();
  return {
    version: 1,
    routes: [
      {
        source: "submission",
        selection: "Data",
        destination:
          options.kind === "existing-table"
            ? { kind: "existing-table", table: options.table }
            : {
                kind: "new-table",
                schema: {
                  name: options.table,
                  columns: [
                    { name: "First", type: "text" },
                    { name: "Second", type: "text" },
                  ],
                  recordId: {
                    prefix: options.table === "Records" ? "REC" : "ARC",
                    padding: 4,
                  },
                  foreignKeys: [],
                },
              },
        columns,
      },
    ],
  };
}

async function preparePlan(options: {
  readonly database: Parameters<typeof createPreparedImport>[0]["database"];
  readonly path: string;
  readonly recipe: ImportRecipe;
}) {
  const baselineRevision = (
    await inspectDatabase({
      database: options.database,
    })
  ).revision;
  const prepared = await createPreparedImport({
    path: options.path,
    database: options.database,
    recipe: options.recipe,
    baselineRevision,
  });
  const outcome = await prepareImport({
    database: options.database,
    prepared,
    recipe: options.recipe,
    sources: [source()],
  });
  const approved = await resolveImport({
    database: options.database,
    prepared,
    decisions: [],
  });
  if (approved.state !== "ready") throw new Error("Plan failed review");
  expect(outcome.prepared.state).toBe("ready");
  return { prepared, approved };
}

async function historyCounts(
  database: Parameters<typeof createPreparedImport>[0]["database"],
) {
  const engine = engineOf(database);
  const [applications, requests, plans] = await Promise.all([
    engine.query(`SELECT count(*) AS count FROM ${APPLICATION_TABLE}`),
    engine.query(`SELECT count(*) AS count FROM ${IMPORT_REQUEST_TABLE}`),
    engine.query(`SELECT count(*) AS count FROM ${PLAN_TABLE}`),
  ]);
  return {
    applications: applications[0]?.["count"],
    requests: requests[0]?.["count"],
    plans: plans[0]?.["count"],
  };
}

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format}: application reuse requires the same effective mapping`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-application-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    try {
      const initialRecipe = recipe({ table: "Records", kind: "new-table" });
      const initial = await preparePlan({
        database,
        path: path.join(directory, "initial.ccplan"),
        recipe: initialRecipe,
      });
      let captureId: string;
      try {
        const applied = await applyImport({
          database,
          prepared: initial.prepared,
          approved: initial.approved,
          requestId: `${format}-initial`,
        });
        captureId = applied.captureIds[0]!;
      } finally {
        await initial.prepared.close();
      }

      const legacyKey = bytesToHex(
        sha256(
          new TextEncoder().encode(
            canonicalJson([
              captureId,
              "Records",
              initialRecipe.routes[0]!.destination,
              initialRecipe.routes[0]!.columns,
            ]),
          ),
        ),
      );
      await engineOf(database).execute(
        `UPDATE ${APPLICATION_TABLE} SET application_key = ?`,
        ["mapping-v1:not-a-digest"],
      );
      const malformed = await preparePlan({
        database,
        path: path.join(directory, "malformed.ccplan"),
        recipe: recipe({ table: "Records", kind: "existing-table" }),
      });
      const beforeMalformed = await inspectDatabase({ database });
      const beforeMalformedHistory = await historyCounts(database);
      try {
        await expect(
          applyImport({
            database,
            prepared: malformed.prepared,
            approved: malformed.approved,
            requestId: `${format}-malformed`,
          }),
        ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });
      } finally {
        await malformed.prepared.close();
      }
      expect(await inspectDatabase({ database })).toEqual(beforeMalformed);
      expect(await historyCounts(database)).toEqual(beforeMalformedHistory);
      await engineOf(database).execute(
        `UPDATE ${APPLICATION_TABLE} SET application_key = ?`,
        [legacyKey],
      );

      const corruptHistoricalRecipe: ImportRecipe = {
        ...initialRecipe,
        routes: [
          {
            ...initialRecipe.routes[0]!,
            columns: initialRecipe.routes[0]!.columns.map((column) =>
              column.source === "A" ? { ...column, target: "Missing" } : column,
            ),
          },
        ],
      };
      const corruptLegacyKey = bytesToHex(
        sha256(
          new TextEncoder().encode(
            canonicalJson([
              captureId,
              "Records",
              corruptHistoricalRecipe.routes[0]!.destination,
              corruptHistoricalRecipe.routes[0]!.columns,
            ]),
          ),
        ),
      );
      await engineOf(database).execute(
        `UPDATE ${PLAN_TABLE} SET recipe_json = ? WHERE plan_id = ? AND plan_revision = ?`,
        [
          canonicalJson(corruptHistoricalRecipe),
          initial.approved.id,
          initial.approved.planRevision,
        ],
      );
      await engineOf(database).execute(
        `UPDATE ${APPLICATION_TABLE} SET application_key = ?`,
        [corruptLegacyKey],
      );
      const corruptHistory = await preparePlan({
        database,
        path: path.join(directory, "corrupt-history.ccplan"),
        recipe: recipe({ table: "Records", kind: "existing-table" }),
      });
      try {
        await expect(
          applyImport({
            database,
            prepared: corruptHistory.prepared,
            approved: corruptHistory.approved,
            requestId: `${format}-corrupt-history`,
          }),
        ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });
      } finally {
        await corruptHistory.prepared.close();
      }
      await engineOf(database).execute(
        `UPDATE ${PLAN_TABLE} SET recipe_json = ? WHERE plan_id = ? AND plan_revision = ?`,
        [
          canonicalJson(initialRecipe),
          initial.approved.id,
          initial.approved.planRevision,
        ],
      );
      await engineOf(database).execute(
        `UPDATE ${APPLICATION_TABLE} SET application_key = ?`,
        [legacyKey],
      );

      const changedRecipe = recipe({
        table: "Records",
        kind: "existing-table",
        swapped: true,
      });
      const changed = await preparePlan({
        database,
        path: path.join(directory, "changed.ccplan"),
        recipe: changedRecipe,
      });
      const beforeConflict = await inspectDatabase({ database });
      const beforeHistory = await historyCounts(database);
      try {
        expect(
          (
            await inspectImport({
              database,
              prepared: changed.prepared,
              page: { limit: 1 },
            })
          ).routes[0]?.applicationState,
        ).toBe("mapping-conflict");
        await expect(
          applyImport({
            database,
            prepared: changed.prepared,
            approved: changed.approved,
            requestId: `${format}-changed`,
          }),
        ).rejects.toMatchObject({
          code: "DB_IMPORT_APPLICATION_CONFLICT",
          details: { captureId, table: "Records" },
        });
      } finally {
        await changed.prepared.close();
      }
      expect(await inspectDatabase({ database })).toEqual(beforeConflict);
      expect(await historyCounts(database)).toEqual(beforeHistory);
      await expect(
        engineOf(database).query(
          "SELECT First, Second FROM Records ORDER BY record_id",
        ),
      ).resolves.toEqual([{ First: "North", Second: "South" }]);

      const repeatedBase = recipe({
        table: "Ｒｅｃｏｒｄｓ",
        kind: "existing-table",
        reversed: true,
      });
      const repeatedRecipe: ImportRecipe = {
        ...repeatedBase,
        routes: [
          {
            ...repeatedBase.routes[0]!,
            columns: repeatedBase.routes[0]!.columns.map((column) => ({
              ...column,
              target: column.target === "First" ? "Ｆｉｒｓｔ" : "Ｓｅｃｏｎｄ",
            })),
          },
        ],
      };
      const repeated = await preparePlan({
        database,
        path: path.join(directory, "repeated.ccplan"),
        recipe: repeatedRecipe,
      });
      try {
        expect(
          (
            await inspectImport({
              database,
              prepared: repeated.prepared,
              page: { limit: 1 },
            })
          ).routes[0],
        ).toMatchObject({ reused: true, applicationState: "already-applied" });
        await expect(
          applyImport({
            database,
            prepared: repeated.prepared,
            approved: repeated.approved,
            requestId: `${format}-repeated`,
          }),
        ).resolves.toMatchObject({
          metrics: { rowsImported: 0, rowsReused: 1, tablesCreated: 0 },
        });
      } finally {
        await repeated.prepared.close();
      }
      expect((await inspectDatabase({ database })).tables).toMatchObject([
        { name: "Records", rowCount: 1n },
      ]);

      const archiveRecipe = recipe({ table: "Archive", kind: "new-table" });
      const archive = await preparePlan({
        database,
        path: path.join(directory, "archive.ccplan"),
        recipe: archiveRecipe,
      });
      try {
        expect(
          (
            await inspectImport({
              database,
              prepared: archive.prepared,
              page: { limit: 1 },
            })
          ).routes[0],
        ).toMatchObject({ reused: true, applicationState: "not-applied" });
        await expect(
          applyImport({
            database,
            prepared: archive.prepared,
            approved: archive.approved,
            requestId: `${format}-archive`,
          }),
        ).resolves.toMatchObject({
          metrics: { rowsImported: 1, rowsReused: 0, tablesCreated: 1 },
        });
      } finally {
        await archive.prepared.close();
      }
      await expect(
        engineOf(database).query("SELECT First, Second FROM Archive"),
      ).resolves.toEqual([{ First: "North", Second: "South" }]);
    } finally {
      await database.close();
    }
  });
}
