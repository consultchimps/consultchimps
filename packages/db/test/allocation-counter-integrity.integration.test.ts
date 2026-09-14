import { copyFile, mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterAll, afterEach, beforeAll, expect, test } from "vitest";

import { engineOf, inspectDatabase } from "../src/database.js";
import {
  applyImport,
  prepareImport,
  resolveImport,
} from "../src/import/operations.js";
import { recordBatch } from "../src/import/deliveries.js";
import type { ImportProfile, ImportSource } from "../src/import/types.js";
import type { EngineTransaction } from "../src/internal/engine.js";
import {
  APPLICATION_TABLE,
  CAPTURE_TABLE,
  COUNTERS_TABLE,
  DELIVERY_MEMBERSHIP_TABLE,
  DELIVERY_TABLE,
  IMPORT_REQUEST_TABLE,
  SOURCE_FILE_TABLE,
  TABLE_REGISTRY_TABLE,
} from "../src/metadata.js";
import {
  createDatabase,
  createImportBatch,
  openDatabase,
  openImportBatch,
} from "../src/node.js";
import type { DatabaseFormat } from "../src/schema.js";

const directories: string[] = [];
let templateDirectory: string;
let databaseTemplates: Record<DatabaseFormat, string>;
let planTemplates: Record<DatabaseFormat, string>;
let reusedPlanTemplates: Record<DatabaseFormat, string>;
let widthPlanTemplates: Record<DatabaseFormat, string>;

const initialRecipe: ImportProfile = {
  version: 1,
  routes: [
    {
      source: "initial",
      selection: "Data",
      destination: {
        kind: "new-table",
        schema: {
          name: "Records",
          columns: [{ name: "Value", type: "text" }],
          recordId: { prefix: "REC", padding: 1 },
          foreignKeys: [],
        },
      },
      columns: [{ source: "Value", target: "Value", type: "text" }],
    },
  ],
};

const nextRecipe: ImportProfile = {
  version: 1,
  routes: [
    {
      source: "next",
      selection: "Data",
      destination: { kind: "existing-table", table: "Records" },
      columns: [{ source: "Value", target: "Value", type: "text" }],
    },
  ],
};

const reusedRecipe: ImportProfile = {
  ...nextRecipe,
  routes: [{ ...nextRecipe.routes[0]!, source: "initial" }],
};

const widthRecipe: ImportProfile = {
  ...nextRecipe,
  routes: [{ ...nextRecipe.routes[0]!, source: "width" }],
};

function source(key: string, input: string | readonly string[]): ImportSource {
  const values = typeof input === "string" ? [input] : input;
  const valueText = values.join(",");
  const bytes = new TextEncoder().encode(`${key}:${valueText}`);
  return {
    key,
    readerVersion: "synthetic-1",
    bytes: {
      name: `${key}.xlsx`,
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
            columns: ["Value"],
            async *batches() {
              yield values.map((value, index) => ({
                sourceRow: index + 2,
                cells: { Value: { kind: "string" as const, value } },
              }));
            },
            async close() {},
          };
        },
      },
    ],
  };
}

async function prepareReady(options: {
  readonly database: Parameters<typeof createImportBatch>[0]["database"];
  readonly filePath: string;
  readonly profile: ImportProfile;
  readonly source: ImportSource;
}) {
  const prepared = await createImportBatch({
    path: options.filePath,
    database: options.database,
    profile: options.profile,
    baselineRevision: (await inspectDatabase({ database: options.database }))
      .revision,
  });
  await prepareImport({
    database: options.database,
    prepared,
    profile: options.profile,
    sources: [options.source],
  });
  const approved = await resolveImport({
    database: options.database,
    prepared,
    decisions: [],
  });
  if (approved.state !== "ready") throw new Error("Plan failed review");
  return { prepared, approved };
}

beforeAll(async () => {
  templateDirectory = await mkdtemp(
    path.join(tmpdir(), "cc-allocation-counter-template-"),
  );
  databaseTemplates = {} as Record<DatabaseFormat, string>;
  planTemplates = {} as Record<DatabaseFormat, string>;
  reusedPlanTemplates = {} as Record<DatabaseFormat, string>;
  widthPlanTemplates = {} as Record<DatabaseFormat, string>;
  for (const format of ["sqlite", "duckdb"] as const) {
    const databasePath = path.join(templateDirectory, `workspace.${format}`);
    const { database } = await createDatabase({ path: databasePath, format });
    const initial = await prepareReady({
      database,
      filePath: path.join(templateDirectory, `initial-${format}.ccplan`),
      profile: initialRecipe,
      source: source("initial", "First"),
    });
    try {
      await applyImport({
        database,
        prepared: initial.prepared,
        approved: initial.approved,
        requestId: `initial-${format}`,
        batchContext: { label: "Initial delivery", scope: { kind: "full" } },
      });
    } finally {
      await initial.prepared.close();
    }
    const reusedPlanPath = path.join(
      templateDirectory,
      `reused-${format}.ccplan`,
    );
    const reused = await prepareReady({
      database,
      filePath: reusedPlanPath,
      profile: reusedRecipe,
      source: source("initial", "First"),
    });
    await reused.prepared.close();
    const planPath = path.join(templateDirectory, `next-${format}.ccplan`);
    const next = await prepareReady({
      database,
      filePath: planPath,
      profile: nextRecipe,
      source: source("next", "Second"),
    });
    await next.prepared.close();
    const widthPlanPath = path.join(
      templateDirectory,
      `width-${format}.ccplan`,
    );
    const width = await prepareReady({
      database,
      filePath: widthPlanPath,
      profile: widthRecipe,
      source: source("width", ["Ninth", "Tenth", "Eleventh"]),
    });
    await width.prepared.close();
    await database.checkpoint();
    await database.close();
    databaseTemplates[format] = databasePath;
    planTemplates[format] = planPath;
    reusedPlanTemplates[format] = reusedPlanPath;
    widthPlanTemplates[format] = widthPlanPath;
  }
});

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

afterAll(async () => {
  await rm(templateDirectory, { recursive: true, force: true });
});

async function fixture(
  format: DatabaseFormat,
  planTemplate = planTemplates[format],
) {
  const directory = await mkdtemp(
    path.join(tmpdir(), "cc-allocation-counter-"),
  );
  directories.push(directory);
  const databasePath = path.join(directory, `workspace.${format}`);
  const planPath = path.join(directory, "next.ccplan");
  await Promise.all([
    copyFile(databaseTemplates[format], databasePath),
    copyFile(planTemplate, planPath),
  ]);
  return {
    database: await openDatabase({ path: databasePath }),
    prepared: await openImportBatch({ path: planPath }),
  };
}

async function databaseCopy(format: DatabaseFormat, suffix: string) {
  const directory = await mkdtemp(
    path.join(tmpdir(), "cc-allocation-counter-open-"),
  );
  directories.push(directory);
  const filePath = path.join(directory, `${suffix}.${format}`);
  await copyFile(databaseTemplates[format], filePath);
  return filePath;
}

async function stateSnapshot(
  database: Awaited<ReturnType<typeof openDatabase>>,
) {
  const engine = engineOf(database);
  const [inspection, counters, registry, rows, metadataCounts] =
    await Promise.all([
      inspectDatabase({ database }),
      engine.query(
        `SELECT counter_name, next_value FROM ${COUNTERS_TABLE} ORDER BY counter_name`,
      ),
      engine.query(
        `SELECT table_name, next_record_id FROM ${TABLE_REGISTRY_TABLE} ORDER BY table_name`,
      ),
      engine.query(
        `SELECT record_id, _imported_row_id, Value FROM Records ORDER BY record_id`,
      ),
      Promise.all(
        [
          SOURCE_FILE_TABLE,
          CAPTURE_TABLE,
          APPLICATION_TABLE,
          IMPORT_REQUEST_TABLE,
          DELIVERY_TABLE,
          DELIVERY_MEMBERSHIP_TABLE,
        ].map(async (table) => ({
          table,
          rows: await engine.query(`SELECT count(*) AS count FROM ${table}`),
        })),
      ),
    ]);
  return { inspection, counters, registry, rows, metadataCounts };
}

const globalCounters = [
  "source_file",
  "capture",
  "import",
  "imported_row",
] as const;

const metadataCounters = [
  "source_file",
  "capture",
  "import",
  "delivery",
] as const;

for (const format of ["sqlite", "duckdb"] as const) {
  for (const counter of metadataCounters) {
    test(`${format}: public open rejects a ${counter} counter behind persisted IDs`, async () => {
      const filePath = await databaseCopy(format, `${counter}-open`);
      const database = await openDatabase({ path: filePath });
      await engineOf(database).execute(
        `UPDATE ${COUNTERS_TABLE} SET next_value = 1 WHERE counter_name = ?`,
        [counter],
      );
      await database.checkpoint();
      await database.close();
      const before = await readFile(filePath);

      const failure = await openDatabase({
        path: filePath,
        readonly: true,
      }).then(
        async (opened) => {
          await opened.close();
          return undefined;
        },
        (error: unknown) => error,
      );
      expect(failure).toMatchObject({ code: "DB_CORRUPT_DATABASE" });
      expect((await readFile(filePath)).equals(before)).toBe(true);
    });
  }

  for (const counter of globalCounters) {
    test(`${format}: rejects a ${counter} counter behind persisted IDs`, async () => {
      const { database, prepared } = await fixture(format);
      try {
        const approved = await resolveImport({
          database,
          prepared,
          decisions: [],
        });
        if (approved.state !== "ready") throw new Error("Plan failed review");
        await engineOf(database).execute(
          `UPDATE ${COUNTERS_TABLE} SET next_value = 1 WHERE counter_name = ?`,
          [counter],
        );
        const before = await stateSnapshot(database);

        await expect(
          applyImport({
            database,
            prepared,
            approved,
            requestId: `${format}-${counter}-corrupt`,
          }),
        ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });

        expect(await stateSnapshot(database)).toEqual(before);
      } finally {
        await prepared.close();
        await database.close();
      }
    });
  }

  test(`${format}: rejects a table Record ID counter behind persisted IDs`, async () => {
    const { database, prepared } = await fixture(format);
    try {
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [],
      });
      if (approved.state !== "ready") throw new Error("Plan failed review");
      await engineOf(database).execute(
        `UPDATE ${TABLE_REGISTRY_TABLE} SET next_record_id = 1 WHERE table_name = 'Records'`,
      );
      const before = await stateSnapshot(database);

      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: `${format}-record-corrupt`,
        }),
      ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });

      expect(await stateSnapshot(database)).toEqual(before);
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  for (const collision of ["null provenance", "edited older ID"] as const) {
    test(`${format}: rejects a generated Record ID collision from ${collision}`, async () => {
      const { database, prepared } = await fixture(format);
      try {
        const approved = await resolveImport({
          database,
          prepared,
          decisions: [],
        });
        if (approved.state !== "ready") throw new Error("Plan failed review");
        if (collision === "null provenance") {
          await engineOf(database).execute(
            "INSERT INTO Records (record_id, Value) VALUES ('REC-2', 'External')",
          );
        } else {
          await engineOf(database).execute(
            "UPDATE Records SET record_id = 'REC-2' WHERE _imported_row_id = 1",
          );
        }
        const before = await stateSnapshot(database);

        await expect(
          applyImport({
            database,
            prepared,
            approved,
            requestId: `${format}-${collision.replaceAll(" ", "-")}`,
          }),
        ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });

        expect(await stateSnapshot(database)).toEqual(before);
      } finally {
        await prepared.close();
        await database.close();
      }
    });
  }

  test(`${format}: an already-applied route does not scan managed rows`, async () => {
    const { database, prepared } = await fixture(
      format,
      reusedPlanTemplates[format],
    );
    const engine = engineOf(database);
    const originalTransaction = engine.transaction.bind(engine);
    let managedTableReads = 0;
    engine.transaction = async <T>(
      work: (transaction: EngineTransaction) => Promise<T>,
    ): Promise<T> =>
      originalTransaction(async (transaction) =>
        work({
          ...transaction,
          query: async (sql, values) => {
            if (
              /\bFROM\s+"Records"(?=\s|$)/iu.test(sql) &&
              /\bORDER BY _imported_row_id DESC LIMIT 1\b/iu.test(sql)
            ) {
              managedTableReads += 1;
            }
            return transaction.query(sql, values);
          },
        }),
      );
    try {
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [],
      });
      if (approved.state !== "ready") throw new Error("Plan failed review");

      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: `${format}-already-applied`,
        }),
      ).resolves.toMatchObject({
        metrics: { rowsImported: 0, rowsReused: 1 },
      });
      expect(managedTableReads).toBe(0);
    } finally {
      engine.transaction = originalTransaction;
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: rejects a generated Record ID collision across a width boundary`, async () => {
    const { database, prepared } = await fixture(
      format,
      widthPlanTemplates[format],
    );
    try {
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [],
      });
      if (approved.state !== "ready") throw new Error("Plan failed review");
      await engineOf(database).execute(
        `UPDATE ${TABLE_REGISTRY_TABLE} SET next_record_id = 9 WHERE table_name = 'Records'`,
      );
      await engineOf(database).execute(
        "INSERT INTO Records (record_id, Value) VALUES ('REC-10', 'External collision')",
      );
      const before = await stateSnapshot(database);

      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: `${format}-width-collision`,
        }),
      ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });

      expect(await stateSnapshot(database)).toEqual(before);
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: generated Record ID checks allow nearby custom IDs`, async () => {
    const { database, prepared } = await fixture(
      format,
      widthPlanTemplates[format],
    );
    try {
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [],
      });
      if (approved.state !== "ready") throw new Error("Plan failed review");
      await engineOf(database).execute(
        `UPDATE ${TABLE_REGISTRY_TABLE} SET next_record_id = 9 WHERE table_name = 'Records'`,
      );
      await engineOf(database).execute(
        "INSERT INTO Records (record_id, Value) VALUES ('REC-10x', 'External custom ID')",
      );

      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: `${format}-width-custom`,
        }),
      ).resolves.toMatchObject({
        metrics: { rowsImported: 3, rowsReused: 0 },
      });
      await expect(
        engineOf(database).query(
          "SELECT record_id, Value FROM Records WHERE record_id IN ('REC-9', 'REC-10', 'REC-11') ORDER BY record_id",
        ),
      ).resolves.toEqual([
        { record_id: "REC-10", Value: "Tenth" },
        { record_id: "REC-11", Value: "Eleventh" },
        { record_id: "REC-9", Value: "Ninth" },
      ]);
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: rejects a delivery counter behind persisted IDs`, async () => {
    const { database, prepared } = await fixture(format);
    try {
      await engineOf(database).execute(
        `UPDATE ${COUNTERS_TABLE} SET next_value = 1 WHERE counter_name = 'delivery'`,
      );
      const captures = await engineOf(database).query(
        `SELECT capture_id FROM ${CAPTURE_TABLE} ORDER BY capture_id`,
      );
      const captureId = captures[0]?.["capture_id"];
      if (typeof captureId !== "string") {
        throw new Error("Fixture capture was not stored");
      }
      const before = await stateSnapshot(database);

      await expect(
        recordBatch({
          database,
          captureIds: [captureId],
          context: { label: "Next delivery", scope: { kind: "full" } },
          requestId: `${format}-delivery-corrupt`,
        }),
      ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });

      expect(await stateSnapshot(database)).toEqual(before);
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: allocation permits deliberate gaps above persisted IDs`, async () => {
    const { database, prepared } = await fixture(format);
    try {
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [],
      });
      if (approved.state !== "ready") throw new Error("Plan failed review");
      await engineOf(database).execute(
        `UPDATE ${COUNTERS_TABLE} SET next_value = 3 WHERE counter_name IN ('source_file', 'capture', 'import', 'delivery', 'imported_row')`,
      );
      await engineOf(database).execute(
        `UPDATE ${TABLE_REGISTRY_TABLE} SET next_record_id = 3 WHERE table_name = 'Records'`,
      );

      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: `${format}-gap`,
          batchContext: { label: "Gap delivery", scope: { kind: "full" } },
        }),
      ).resolves.toMatchObject({
        importIds: ["IMP-000003"],
        captureIds: ["CAP-000003"],
        batchId: "DEL-000003",
        metrics: { rowsImported: 1, rowsReused: 0, batchesRecorded: 1 },
      });
      await expect(
        engineOf(database).query(
          "SELECT record_id, _imported_row_id, Value FROM Records ORDER BY record_id",
        ),
      ).resolves.toEqual([
        { record_id: "REC-1", _imported_row_id: 1n, Value: "First" },
        { record_id: "REC-3", _imported_row_id: 3n, Value: "Second" },
      ]);
    } finally {
      await prepared.close();
      await database.close();
    }
  });
}
