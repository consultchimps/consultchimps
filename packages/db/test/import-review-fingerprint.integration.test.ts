import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { sha256 } from "@noble/hashes/sha2.js";
import { bytesToHex } from "@noble/hashes/utils.js";
import { afterEach, expect, test } from "vitest";

import { inspectDatabase } from "../src/database.js";
import { inspectImport } from "../src/import/inspection.js";
import { applyImport, prepareImport } from "../src/import/operations.js";
import type {
  ImportRecipe,
  ImportSource,
  ReadyImportRef,
} from "../src/import/types.js";
import { canonicalJson } from "../src/internal/json.js";
import {
  PREPARED_BINDING_TABLE,
  PREPARED_CAPTURE_TABLE,
  PREPARED_METADATA_TABLE,
  preparedEngineOf,
} from "../src/prepared.js";
import {
  createDatabase,
  createPreparedImport,
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

const originalRecipe: ImportRecipe = {
  version: 1,
  routes: [
    {
      source: "submission",
      selection: "Data",
      destination: {
        kind: "new-table",
        schema: {
          name: "Observations",
          columns: [{ name: "Value", type: "text" }],
          recordId: { prefix: "OBS", padding: 4 },
        },
      },
      columns: [{ source: "Approved", target: "Value", type: "text" }],
    },
  ],
};

const changedRecipe: ImportRecipe = {
  ...originalRecipe,
  routes: [
    {
      ...originalRecipe.routes[0]!,
      columns: [{ source: "Other", target: "Value", type: "text" }],
    },
  ],
};

function source(): ImportSource {
  const bytes = new TextEncoder().encode("synthetic review fingerprint");
  return {
    key: "submission",
    readerVersion: "synthetic-1",
    bytes: {
      name: "submission.xlsx",
      size: bytes.byteLength,
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
            columns: ["Approved", "Other"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: {
                    Approved: { kind: "string" as const, value: "reviewed" },
                    Other: { kind: "string" as const, value: "not reviewed" },
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

function namedSource(options: {
  readonly key: string;
  readonly selection: string;
  readonly value: string;
  readonly opens?: { count: number } | undefined;
  readonly failOpen?: boolean | undefined;
}): ImportSource {
  const bytes = new TextEncoder().encode(
    `${options.key}:${options.selection}:${options.value}`,
  );
  return {
    key: options.key,
    readerVersion: "synthetic-1",
    bytes: {
      name: `${options.key}.xlsx`,
      size: bytes.byteLength,
      async readAt(offset, length) {
        return bytes.slice(offset, offset + length);
      },
    },
    selections: [
      {
        key: options.selection,
        label: options.selection,
        async open() {
          if (options.opens !== undefined) options.opens.count += 1;
          if (options.failOpen) throw new Error("synthetic reader failure");
          return {
            columns: ["Value"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: {
                    Value: {
                      kind: "string" as const,
                      value: options.value,
                    },
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

function route(
  sourceKey: string,
  selection: string,
): ImportRecipe["routes"][number] {
  return {
    source: sourceKey,
    selection,
    destination: {
      kind: "new-table",
      schema: {
        name: `${sourceKey}Table`,
        columns: [{ name: "Value", type: "text" }],
        recordId: { prefix: sourceKey.toUpperCase(), padding: 4 },
      },
    },
    columns: [{ source: "Value", target: "Value", type: "text" }],
  };
}

async function fixture(format: "sqlite" | "duckdb") {
  const directory = await mkdtemp(
    path.join(tmpdir(), "cc-review-fingerprint-"),
  );
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, `workspace.${format}`),
    format,
  });
  const planPath = path.join(directory, "review.ccplan");
  const prepared = await createPreparedImport({
    path: planPath,
    database,
    recipe: originalRecipe,
    baselineRevision: 0n,
  });
  const outcome = await prepareImport({
    database,
    prepared,
    recipe: originalRecipe,
    sources: [source()],
  });
  if (outcome.prepared.state !== "ready") {
    throw new Error("Synthetic plan did not become ready");
  }
  return {
    database,
    prepared,
    planPath,
    approved: outcome.prepared,
  };
}

function requiredString(value: unknown, field: string): string {
  if (typeof value === "string") return value;
  throw new Error(`Expected ${field} to be a string`);
}

function requiredBigInt(value: unknown, field: string): bigint {
  if (typeof value === "bigint") return value;
  if (typeof value === "number" && Number.isSafeInteger(value)) {
    return BigInt(value);
  }
  throw new Error(`Expected ${field} to be an integer`);
}

function compareStoredTuple(
  left: readonly (string | null)[],
  right: readonly (string | null)[],
): number {
  const leftKey = JSON.stringify(left);
  const rightKey = JSON.stringify(right);
  return leftKey < rightKey ? -1 : leftKey > rightKey ? 1 : 0;
}

async function reviewedCaptureMetadata(
  engine: ReturnType<typeof preparedEngineOf>,
): Promise<readonly [string, string]> {
  const captureRows = await engine.query(
    `SELECT capture_id, source_file_id, source_key, display_name, selection_key, selection_label, reader_version, content_hash, byte_count, reused, row_count, columns_json FROM ${PREPARED_CAPTURE_TABLE}`,
  );
  const captures = captureRows.map(
    (row) =>
      [
        requiredString(row["capture_id"], "capture ID"),
        row["source_file_id"] === null
          ? null
          : requiredString(row["source_file_id"], "source file ID"),
        requiredString(row["source_key"], "source key"),
        requiredString(row["display_name"], "display name"),
        requiredString(row["selection_key"], "selection key"),
        requiredString(row["selection_label"], "selection label"),
        requiredString(row["reader_version"], "reader version"),
        requiredString(row["content_hash"], "content hash"),
        requiredBigInt(row["byte_count"], "byte count").toString(),
        requiredBigInt(row["reused"], "reuse marker").toString(),
        requiredBigInt(row["row_count"], "row count").toString(),
        requiredString(row["columns_json"], "captured columns"),
      ] satisfies readonly (string | null)[],
  );
  const bindingRows = await engine.query(
    `SELECT source_key, selection_key, capture_id, display_name FROM ${PREPARED_BINDING_TABLE}`,
  );
  const bindings = bindingRows.map(
    (row) =>
      [
        requiredString(row["source_key"], "source key"),
        requiredString(row["selection_key"], "selection key"),
        requiredString(row["capture_id"], "capture ID"),
        requiredString(row["display_name"], "display name"),
      ] satisfies readonly (string | null)[],
  );
  return [
    JSON.stringify(captures.sort(compareStoredTuple)),
    JSON.stringify(bindings.sort(compareStoredTuple)),
  ];
}

async function replaceReviewedRecipe(options: {
  readonly prepared: Parameters<typeof preparedEngineOf>[0];
  readonly recipe: ImportRecipe;
  readonly updateFingerprint: boolean;
}): Promise<void> {
  const engine = preparedEngineOf(options.prepared);
  const recipeJson = canonicalJson(options.recipe);
  if (!options.updateFingerprint) {
    await engine.execute(
      `UPDATE ${PREPARED_METADATA_TABLE} SET recipe_json = ?`,
      [recipeJson],
    );
    return;
  }
  const rows = await engine.query(
    `SELECT format_version, plan_id, database_id, baseline_revision, schema_fingerprint, plan_revision, state, conflicts_json, decisions_json FROM ${PREPARED_METADATA_TABLE}`,
  );
  const row = rows[0];
  if (row === undefined) throw new Error("Prepared metadata is missing");
  const [captureDefinitionsJson, bindingsJson] =
    await reviewedCaptureMetadata(engine);
  const reviewedFields = [
    requiredBigInt(row["format_version"], "format version").toString(),
    requiredString(row["plan_id"], "plan ID"),
    requiredString(row["database_id"], "database ID"),
    requiredBigInt(row["baseline_revision"], "baseline revision").toString(),
    requiredString(row["schema_fingerprint"], "schema fingerprint"),
    requiredBigInt(row["plan_revision"], "plan revision").toString(),
    requiredString(row["state"], "state"),
    recipeJson,
    requiredString(row["conflicts_json"], "conflicts"),
    requiredString(row["decisions_json"], "decisions"),
    captureDefinitionsJson,
    bindingsJson,
  ];
  const fingerprint = bytesToHex(
    sha256(new TextEncoder().encode(JSON.stringify(reviewedFields))),
  );
  await engine.execute(
    `UPDATE ${PREPARED_METADATA_TABLE} SET recipe_json = ?, review_fingerprint = ?`,
    [recipeJson, fingerprint],
  );
}

async function expectUnchangedEmptyDatabase(
  database: Parameters<typeof inspectDatabase>[0]["database"],
): Promise<void> {
  await expect(inspectDatabase({ database })).resolves.toMatchObject({
    revision: 0n,
    tables: [],
    captures: 0n,
    completedImports: 0n,
    deliveries: 0n,
  });
}

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format}: changed reviewed metadata cannot use an earlier approval`, async () => {
    const { database, prepared, approved } = await fixture(format);
    try {
      await replaceReviewedRecipe({
        prepared,
        recipe: changedRecipe,
        updateFingerprint: false,
      });
      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: `${format}-changed-review`,
        }),
      ).rejects.toMatchObject({ code: "DB_INVALID_PREPARED_IMPORT" });
      await expectUnchangedEmptyDatabase(database);
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: a coherent metadata rewrite still invalidates an earlier approval`, async () => {
    const { database, prepared, approved } = await fixture(format);
    try {
      await replaceReviewedRecipe({
        prepared,
        recipe: changedRecipe,
        updateFingerprint: true,
      });
      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: `${format}-stale-review`,
        }),
      ).rejects.toMatchObject({ code: "DB_STALE_IMPORT_PLAN" });
      await expectUnchangedEmptyDatabase(database);
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: capture definitions and source bindings remain part of the approval`, async () => {
    const { database, prepared, approved } = await fixture(format);
    const engine = preparedEngineOf(prepared);
    try {
      const bindingRows = await engine.query(
        `SELECT display_name FROM ${PREPARED_BINDING_TABLE}`,
      );
      const originalDisplayName = requiredString(
        bindingRows[0]?.["display_name"],
        "display name",
      );
      await engine.execute(
        `UPDATE ${PREPARED_BINDING_TABLE} SET display_name = ?`,
        ["changed-name.xlsx"],
      );
      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: `${format}-changed-binding`,
        }),
      ).rejects.toMatchObject({ code: "DB_INVALID_PREPARED_IMPORT" });
      await expectUnchangedEmptyDatabase(database);

      await engine.execute(
        `UPDATE ${PREPARED_BINDING_TABLE} SET display_name = ?`,
        [originalDisplayName],
      );
      await engine.execute(
        `UPDATE ${PREPARED_CAPTURE_TABLE} SET selection_label = ?`,
        ["Changed selection"],
      );
      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: `${format}-changed-capture`,
        }),
      ).rejects.toMatchObject({ code: "DB_INVALID_PREPARED_IMPORT" });
      await expectUnchangedEmptyDatabase(database);
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: an unchanged ready plan reopens and replays its receipt`, async () => {
    const fixtureValue = await fixture(format);
    const { database, planPath } = fixtureValue;
    await fixtureValue.prepared.close();
    const prepared = await openPreparedImport({ path: planPath });
    try {
      const inspection = await inspectImport({
        prepared,
        database,
        page: { limit: 1 },
      });
      if (inspection.prepared.state !== "ready") {
        throw new Error("Reopened plan did not remain ready");
      }
      const approved: ReadyImportRef = inspection.prepared;
      const first = await applyImport({
        database,
        prepared,
        approved,
        requestId: `${format}-reopen-replay`,
      });
      const retry = await applyImport({
        database,
        prepared,
        approved,
        requestId: `${format}-reopen-replay`,
      });
      expect(first.metrics.rowsImported).toBe(1);
      expect(retry.importIds).toEqual(first.importIds);
      expect(retry.metrics).toMatchObject({ rowsImported: 0, rowsReused: 1 });
      await expect(inspectDatabase({ database })).resolves.toMatchObject({
        revision: 1n,
        tables: [{ name: "Observations", rowCount: 1n }],
      });
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: a damaged ready plan is rejected after reopen without further writes`, async () => {
    const fixtureValue = await fixture(format);
    const { database, planPath } = fixtureValue;
    await fixtureValue.prepared.close();
    const editor = await openPreparedImport({ path: planPath });
    await replaceReviewedRecipe({
      prepared: editor,
      recipe: changedRecipe,
      updateFingerprint: false,
    });
    await editor.close();
    const before = await readFile(planPath);
    try {
      await expect(
        openPreparedImport({ path: planPath, readonly: true }),
      ).rejects.toMatchObject({ code: "DB_INVALID_PREPARED_IMPORT" });
      expect((await readFile(planPath)).equals(before)).toBe(true);
      await expectUnchangedEmptyDatabase(database);
    } finally {
      await database.close();
    }
  });

  test(`${format}: partial capture work invalidates an approval and remains retryable`, async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "cc-review-invalidation-"),
    );
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const planPath = path.join(directory, "review.ccplan");
    const initialRecipe: ImportRecipe = {
      version: 1,
      routes: [route("first", "First")],
    };
    const expandedRecipe: ImportRecipe = {
      version: 1,
      routes: [
        ...initialRecipe.routes,
        route("second", "Second"),
        route("third", "Third"),
      ],
    };
    let prepared = await createPreparedImport({
      path: planPath,
      database,
      recipe: initialRecipe,
      baselineRevision: 0n,
    });
    const initial = await prepareImport({
      database,
      prepared,
      recipe: initialRecipe,
      sources: [namedSource({ key: "first", selection: "First", value: "A" })],
    });
    if (initial.prepared.state !== "ready") {
      throw new Error("Initial plan did not become ready");
    }
    const secondOpens = { count: 0 };
    try {
      await expect(
        prepareImport({
          database,
          prepared,
          recipe: expandedRecipe,
          sources: [
            namedSource({
              key: "second",
              selection: "Second",
              value: "B",
              opens: secondOpens,
            }),
            namedSource({
              key: "third",
              selection: "Third",
              value: "C",
              failOpen: true,
            }),
          ],
        }),
      ).rejects.toThrow("synthetic reader failure");
      expect(secondOpens.count).toBe(1);
      await expect(
        applyImport({
          database,
          prepared,
          approved: initial.prepared,
          requestId: `${format}-superseded-review`,
        }),
      ).rejects.toMatchObject({ code: "DB_STALE_IMPORT_PLAN" });
      await expectUnchangedEmptyDatabase(database);

      await prepared.close();
      prepared = await openPreparedImport({ path: planPath });
      const interrupted = await inspectImport({
        prepared,
        database,
        page: { limit: 1 },
      });
      expect(interrupted.prepared).toMatchObject({
        state: "needs-review",
        planRevision: initial.prepared.planRevision + 1n,
      });

      const retried = await prepareImport({
        database,
        prepared,
        recipe: expandedRecipe,
        sources: [
          namedSource({
            key: "second",
            selection: "Second",
            value: "B",
            opens: secondOpens,
          }),
          namedSource({ key: "third", selection: "Third", value: "C" }),
        ],
      });
      expect(secondOpens.count).toBe(1);
      expect(retried.result.metrics.rowsCaptured).toBe(1);
      if (retried.prepared.state !== "ready") {
        throw new Error("Retried plan did not become ready");
      }
      const applied = await applyImport({
        database,
        prepared,
        approved: retried.prepared,
        requestId: `${format}-retried-review`,
      });
      expect(applied.metrics.rowsImported).toBe(3);
    } finally {
      await prepared.close();
      await database.close();
    }
  });
}
