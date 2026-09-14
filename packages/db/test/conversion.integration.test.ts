import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { inspectDatabase } from "../src/database.js";
import {
  applyImport,
  prepareImport,
  resolveImport,
} from "../src/import/operations.js";
import { inspectAppliedImportBatch } from "../src/import/history.js";
import { draftImportProfile } from "../src/import/profile.js";
import type { ImportSource, ReadyImportBatchRef } from "../src/import/types.js";
import {
  createDatabase,
  createImportBatch,
  exportDatabase,
  openDatabase,
} from "../src/node.js";

const directories: string[] = [];
afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

test("converts a managed import from SQLite to DuckDB and back", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-conversion-"));
  directories.push(directory);
  const sqlitePath = path.join(directory, "source.sqlite");
  const { database } = await createDatabase({
    path: sqlitePath,
    format: "sqlite",
  });
  const bytes = new TextEncoder().encode("synthetic conversion source");
  const source: ImportSource = {
    key: "amounts",
    readerVersion: "synthetic-1",
    bytes: {
      name: "amounts.xlsx",
      size: bytes.length,
      async readAt(offset, length) {
        return bytes.slice(offset, offset + length);
      },
    },
    selections: [
      {
        key: "amounts",
        label: "Amounts",
        async open() {
          return {
            columns: ["Code", "Amount", "Active"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: {
                    Code: { kind: "string", value: "001" },
                    Amount: { kind: "number", raw: "1234567890.1234" },
                    Active: { kind: "boolean", value: true },
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
  const profile = await draftImportProfile({ sources: [source] });
  const prepared = await createImportBatch({
    path: path.join(directory, "source.ccplan"),
    database,
    profile,
    baselineRevision: 0n,
  });
  let approvedRef: ReadyImportBatchRef | undefined;
  try {
    await prepareImport({ database, prepared, sources: [source], profile });
    const approved = await resolveImport({ database, prepared, decisions: [] });
    if (approved.state !== "ready") throw new Error("Plan failed review");
    approvedRef = approved;
    await applyImport({
      database,
      prepared,
      approved,
      requestId: "conversion-import",
      batchContext: { label: "September", scope: { kind: "full" } },
    });
  } finally {
    await prepared.close();
  }
  const duckPath = path.join(directory, "converted.duckdb");
  const exported = await exportDatabase({
    database,
    output: duckPath,
    format: "duckdb",
  });
  expect(exported.plan.state).toBe("ready");
  expect(exported.metrics.rowsConverted).toBeGreaterThan(0);
  await database.close();

  const duck = await openDatabase({ path: duckPath });
  const duckInspection = await inspectDatabase({ database: duck });
  expect(duckInspection.tables[0]?.rowCount).toBe(1n);
  expect(duckInspection.appliedImportBatches).toBe(1n);
  if (approvedRef === undefined) throw new Error("Import was not approved");
  const duckPlan = await inspectAppliedImportBatch({
    database: duck,
    planId: approvedRef.id,
    planRevision: approvedRef.planRevision,
  });
  expect(duckPlan).toMatchObject({
    profile: {
      routes: [
        {
          source: "amounts",
          selection: "amounts",
          destination: { kind: "new-table" },
        },
      ],
    },
    decisions: [
      {
        kind: "route",
        source: "amounts",
        selection: "amounts",
      },
    ],
    bindings: [
      {
        source: "amounts",
        displayName: "amounts.xlsx",
        selection: "amounts",
        label: "Amounts",
      },
    ],
  });
  const roundtripPath = path.join(directory, "roundtrip.sqlite");
  await exportDatabase({
    database: duck,
    output: roundtripPath,
    format: "sqlite",
  });
  await duck.close();

  const roundtrip = await openDatabase({ path: roundtripPath });
  try {
    const inspection = await inspectDatabase({ database: roundtrip });
    expect(inspection.tables[0]?.rowCount).toBe(1n);
    expect(inspection.captures).toBe(1n);
    expect(inspection.completedImports).toBe(1n);
    expect(inspection.recordedBatches).toBe(1n);
    expect(inspection.appliedImportBatches).toBe(1n);
    await expect(
      inspectAppliedImportBatch({
        database: roundtrip,
        planId: approvedRef.id,
        planRevision: approvedRef.planRevision,
      }),
    ).resolves.toEqual(duckPlan);
  } finally {
    await roundtrip.close();
  }
});
