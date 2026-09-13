import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { engineOf, inspectDatabase } from "../src/database.js";
import { NodeSqliteEngine } from "../src/engines/sqlite/node.js";
import {
  applyImport,
  prepareImport,
  resolveImport,
} from "../src/import/operations.js";
import type {
  ImportRecipe,
  ImportSource,
  ReadyImportRef,
} from "../src/import/types.js";
import {
  APPLICATION_TABLE,
  CAPTURE_ROW_TABLE,
  CAPTURE_TABLE,
  DELIVERY_MEMBERSHIP_TABLE,
  DELIVERY_TABLE,
  IMPORT_REQUEST_TABLE,
  PLAN_TABLE,
  SOURCE_CONTENT_TABLE,
  SOURCE_FILE_TABLE,
  SOURCE_NAME_TABLE,
} from "../src/metadata.js";
import {
  createDatabase,
  createPreparedImport,
  openPreparedImport,
} from "../src/node.js";
import { PREPARED_ROW_TABLE } from "../src/prepared.js";
import type { DatabaseFormat } from "../src/schema.js";

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

function importSource(rowCount = 1): ImportSource {
  const bytes = new TextEncoder().encode("synthetic source row fixture");
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
              for (let offset = 0; offset < rowCount; offset += 2_000) {
                yield Array.from(
                  { length: Math.min(2_000, rowCount - offset) },
                  (_, index) => ({
                    sourceRow: offset + index + 2,
                    cells: {
                      Value: {
                        kind: "string" as const,
                        value: `Row ${offset + index + 1}`,
                      },
                    },
                  }),
                );
              }
            },
            async close() {},
          };
        },
      },
    ],
  };
}

async function saveReadyPlan(options: {
  readonly database: Parameters<typeof createPreparedImport>[0]["database"];
  readonly planPath: string;
  readonly rowCount?: number | undefined;
  readonly importRecipe?: ImportRecipe | undefined;
  readonly baselineRevision?: bigint | undefined;
}): Promise<ReadyImportRef> {
  const importRecipe = options.importRecipe ?? recipe;
  const prepared = await createPreparedImport({
    path: options.planPath,
    database: options.database,
    recipe: importRecipe,
    baselineRevision: options.baselineRevision ?? 0n,
  });
  try {
    await prepareImport({
      database: options.database,
      prepared,
      recipe: importRecipe,
      sources: [importSource(options.rowCount)],
    });
    const approved = await resolveImport({
      database: options.database,
      prepared,
      decisions: [],
    });
    if (approved.state !== "ready") throw new Error("Plan failed review");
    return approved;
  } finally {
    await prepared.close();
  }
}

const copyRecipe: ImportRecipe = {
  version: 1,
  routes: [
    {
      ...recipe.routes[0]!,
      destination: {
        kind: "new-table",
        schema: {
          name: "InventoryCopy",
          columns: [{ name: "Value", type: "text" }],
          recordId: { prefix: "COPY", padding: 4 },
          foreignKeys: [],
        },
      },
    },
  ],
};

async function expectNoImportHistory(
  database: Parameters<typeof createPreparedImport>[0]["database"],
) {
  expect(await inspectDatabase({ database })).toMatchObject({
    revision: 0n,
    tables: [],
    captures: 0n,
    completedImports: 0n,
    deliveries: 0n,
    appliedImportPlans: 0n,
  });
  const tables = [
    SOURCE_CONTENT_TABLE,
    SOURCE_FILE_TABLE,
    SOURCE_NAME_TABLE,
    CAPTURE_TABLE,
    CAPTURE_ROW_TABLE,
    APPLICATION_TABLE,
    IMPORT_REQUEST_TABLE,
    PLAN_TABLE,
    DELIVERY_TABLE,
    DELIVERY_MEMBERSHIP_TABLE,
  ];
  for (const table of tables) {
    await expect(
      engineOf(database).query(`SELECT count(*) AS count FROM ${table}`),
    ).resolves.toEqual([{ count: 0n }]);
  }
}

const corruptions = [
  {
    name: "zero",
    rowCount: 1,
    sql: `UPDATE ${PREPARED_ROW_TABLE} SET source_row = 0`,
  },
  {
    name: "negative",
    rowCount: 1,
    sql: `UPDATE ${PREPARED_ROW_TABLE} SET source_row = -2`,
  },
  {
    name: "unsafe",
    rowCount: 2_001,
    sql: `UPDATE ${PREPARED_ROW_TABLE} SET source_row = 9007199254740992 WHERE source_row = 2002`,
  },
  {
    name: "non-integer",
    rowCount: 1,
    sql: `UPDATE ${PREPARED_ROW_TABLE} SET source_row = 'not-an-integer'`,
  },
  {
    name: "missing-after-first-page",
    rowCount: 2_001,
    sql: `DELETE FROM ${PREPARED_ROW_TABLE} WHERE source_row = 2002`,
  },
] as const;

for (const format of [
  "sqlite",
  "duckdb",
] as const satisfies readonly DatabaseFormat[]) {
  test(`${format}: damaged prepared source rows cannot mutate the database`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-source-row-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    try {
      for (const corruption of corruptions) {
        const planPath = path.join(directory, `${corruption.name}.ccplan`);
        const approved = await saveReadyPlan({
          database,
          planPath,
          rowCount: corruption.rowCount,
        });
        const planEngine = NodeSqliteEngine.open(planPath);
        try {
          await planEngine.execute(corruption.sql);
        } finally {
          await planEngine.close();
        }

        const prepared = await openPreparedImport({ path: planPath });
        try {
          await expect(
            applyImport({
              database,
              prepared,
              approved,
              requestId: `${format}-${corruption.name}`,
            }),
          ).rejects.toMatchObject({ code: "DB_INVALID_PREPARED_IMPORT" });
          await expectNoImportHistory(database);
        } finally {
          await prepared.close();
        }
      }

      const planPath = path.join(directory, "valid.ccplan");
      const approved = await saveReadyPlan({ database, planPath });
      const prepared = await openPreparedImport({ path: planPath });
      try {
        await expect(
          applyImport({
            database,
            prepared,
            approved,
            requestId: `${format}-valid`,
          }),
        ).resolves.toMatchObject({ metrics: { rowsImported: 1 } });
      } finally {
        await prepared.close();
      }
      await expect(
        engineOf(database).query(
          "SELECT record_id, Value, _source_row FROM Inventory",
        ),
      ).resolves.toEqual([
        { record_id: "ITEM-0001", Value: "Row 1", _source_row: 2n },
      ]);
    } finally {
      await database.close();
    }
  });

  test(`${format}: damaged reused capture rows roll back a new application`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-reused-row-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    try {
      const initialPath = path.join(directory, "initial.ccplan");
      const initialApproval = await saveReadyPlan({
        database,
        planPath: initialPath,
      });
      const initialPlan = await openPreparedImport({ path: initialPath });
      let captureId: string;
      try {
        const applied = await applyImport({
          database,
          prepared: initialPlan,
          approved: initialApproval,
          requestId: `${format}-initial`,
        });
        captureId = applied.captureIds[0]!;
      } finally {
        await initialPlan.close();
      }

      const copyPath = path.join(directory, "copy.ccplan");
      const copyApproval = await saveReadyPlan({
        database,
        planPath: copyPath,
        importRecipe: copyRecipe,
        baselineRevision: 1n,
      });
      const before = await inspectDatabase({ database });
      await engineOf(database).execute(
        `UPDATE ${CAPTURE_ROW_TABLE} SET source_row = 0 WHERE capture_id = ?`,
        [captureId],
      );
      const copyPlan = await openPreparedImport({ path: copyPath });
      try {
        await expect(
          applyImport({
            database,
            prepared: copyPlan,
            approved: copyApproval,
            requestId: `${format}-copy`,
          }),
        ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });
        expect(await inspectDatabase({ database })).toEqual(before);

        await engineOf(database).execute(
          `UPDATE ${CAPTURE_ROW_TABLE} SET source_row = 2 WHERE capture_id = ?`,
          [captureId],
        );
        await expect(
          applyImport({
            database,
            prepared: copyPlan,
            approved: copyApproval,
            requestId: `${format}-copy`,
          }),
        ).resolves.toMatchObject({ metrics: { rowsImported: 1 } });
      } finally {
        await copyPlan.close();
      }
      await expect(
        engineOf(database).query(
          "SELECT record_id, Value, _source_row FROM InventoryCopy",
        ),
      ).resolves.toEqual([
        { record_id: "COPY-0001", Value: "Row 1", _source_row: 2n },
      ]);
    } finally {
      await database.close();
    }
  });
}
