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
import { inspectAppliedImportPlan } from "../src/import/history.js";
import type {
  ImportCell,
  ImportRecipe,
  ImportSource,
} from "../src/import/types.js";
import { createDatabase, createPreparedImport } from "../src/node.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

function source(
  selections: Readonly<
    Record<
      string,
      readonly Readonly<{
        sourceRow: number;
        cells: Record<string, ImportCell>;
      }>[]
    >
  >,
): ImportSource {
  const bytes = new TextEncoder().encode(JSON.stringify(selections));
  return {
    key: "submission",
    readerVersion: "synthetic-review-1",
    bytes: {
      name: "submission.xlsx",
      size: bytes.byteLength,
      async readAt(offset, length) {
        return bytes.slice(offset, offset + length);
      },
    },
    selections: Object.entries(selections).map(([key, rows]) => ({
      key,
      label: key,
      async open() {
        return {
          columns: [...new Set(rows.flatMap((row) => Object.keys(row.cells)))],
          async *batches() {
            yield rows;
          },
          async close() {},
        };
      },
    })),
  };
}

async function fixture(format: "sqlite" | "duckdb", recipe: ImportRecipe) {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-import-review-"));
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
  return { database, prepared };
}

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format}: mapping and value errors block readiness before apply`, async () => {
    const recipe: ImportRecipe = {
      version: 1,
      routes: [
        {
          source: "submission",
          selection: "Data",
          destination: {
            kind: "new-table",
            schema: {
              name: "Data",
              columns: [
                { name: "Amount", type: "integer", nullable: false },
                { name: "Required", type: "text", nullable: false },
                { name: "Extra", type: "text" },
              ],
              recordId: { prefix: "DAT", padding: 6 },
              foreignKeys: [],
            },
          },
          columns: [
            { source: "Amount", target: "Amount", type: "integer" },
            { source: "Missing", target: "Required", type: "text" },
            { source: "amount", target: "Extra", type: "text" },
          ],
        },
      ],
    };
    const { database, prepared } = await fixture(format, recipe);
    try {
      const outcome = await prepareImport({
        database,
        prepared,
        recipe,
        sources: [
          source({
            Data: [
              {
                sourceRow: 2,
                cells: {
                  Amount: { kind: "string", value: "not a number" },
                  Missing: { kind: "blank" },
                },
              },
            ],
          }),
        ],
      });
      expect(outcome.prepared.state).toBe("needs-review");
      expect(outcome.result.metrics.conflicts).toBe(3);
      const resolved = await resolveImport({
        database,
        prepared,
        decisions: [],
      });
      expect(resolved.state).toBe("needs-review");
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: related new tables apply in foreign-key order`, async () => {
    const parentSchema = {
      name: "Parents",
      columns: [{ name: "Name", type: "text", nullable: false }],
      recordId: { prefix: "PAR", padding: 6 },
      foreignKeys: [],
    } as const;
    const childSchema = {
      name: "Children",
      columns: [{ name: "ParentId", type: "text", nullable: false }],
      recordId: { prefix: "CHI", padding: 6 },
      foreignKeys: [{ column: "ParentId", referencesTable: "Parents" }],
    } as const;
    const recipe: ImportRecipe = {
      version: 1,
      routes: [
        {
          source: "submission",
          selection: "Children",
          destination: { kind: "new-table", schema: childSchema },
          columns: [{ source: "ParentId", target: "ParentId", type: "text" }],
        },
        {
          source: "submission",
          selection: "Parents",
          destination: { kind: "new-table", schema: parentSchema },
          columns: [{ source: "Name", target: "Name", type: "text" }],
        },
      ],
    };
    const { database, prepared } = await fixture(format, recipe);
    try {
      const outcome = await prepareImport({
        database,
        prepared,
        recipe,
        sources: [
          source({
            Children: [
              {
                sourceRow: 2,
                cells: { ParentId: { kind: "string", value: "PAR-000001" } },
              },
            ],
            Parents: [
              {
                sourceRow: 2,
                cells: { Name: { kind: "string", value: "Parent A" } },
              },
            ],
          }),
        ],
      });
      expect(outcome.prepared.state).toBe("ready");
      if (outcome.prepared.state !== "ready") throw new Error("Plan not ready");
      await applyImport({
        database,
        prepared,
        approved: outcome.prepared,
        requestId: `${format}-connected-tables`,
      });
      const inspection = await inspectDatabase({ database });
      expect(
        inspection.tables.map(({ name, rowCount }) => [name, rowCount]),
      ).toEqual([
        ["Children", 1n],
        ["Parents", 1n],
      ]);
      expect(inspection.appliedImportPlans).toBe(1n);
      const saved = await inspectAppliedImportPlan({
        database,
        planId: outcome.prepared.id,
        planRevision: outcome.prepared.planRevision,
      });
      expect(saved).toMatchObject({
        state: "applied",
        recipe,
        conflicts: [],
        bindings: [
          {
            source: "submission",
            displayName: "submission.xlsx",
            selection: "Children",
            label: "Children",
          },
          {
            source: "submission",
            displayName: "submission.xlsx",
            selection: "Parents",
            label: "Parents",
          },
        ],
      });
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: multiple sources can append to one table declared later`, async () => {
    const sharedSchema = {
      name: "Shared",
      columns: [{ name: "Value", type: "text", nullable: false }],
      recordId: { prefix: "SHA", padding: 6 },
      foreignKeys: [],
    } as const;
    const recipe: ImportRecipe = {
      version: 1,
      routes: [
        {
          source: "first",
          selection: "Data",
          destination: { kind: "existing-table", table: "Shared" },
          columns: [{ source: "Value", target: "Value", type: "text" }],
        },
        {
          source: "second",
          selection: "Data",
          destination: { kind: "new-table", schema: sharedSchema },
          columns: [{ source: "Value", target: "Value", type: "text" }],
        },
      ],
    };
    const { database, prepared } = await fixture(format, recipe);
    try {
      const first = source({
        Data: [
          {
            sourceRow: 2,
            cells: { Value: { kind: "string", value: "A" } },
          },
        ],
      });
      const second = source({
        Data: [
          {
            sourceRow: 2,
            cells: { Value: { kind: "string", value: "B" } },
          },
        ],
      });
      const outcome = await prepareImport({
        database,
        prepared,
        recipe,
        sources: [
          { ...first, key: "first" },
          { ...second, key: "second" },
        ],
      });
      expect(outcome.prepared.state).toBe("ready");
      if (outcome.prepared.state !== "ready") throw new Error("Plan not ready");
      const result = await applyImport({
        database,
        prepared,
        approved: outcome.prepared,
        requestId: `${format}-shared-table`,
      });
      expect(result.metrics).toMatchObject({
        rowsImported: 2,
        tablesCreated: 1,
      });
      expect((await inspectDatabase({ database })).tables).toMatchObject([
        { name: "Shared", rowCount: 2n },
      ]);
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: conflicting declarations for one new table require review`, async () => {
    const recipe: ImportRecipe = {
      version: 1,
      routes: [
        {
          source: "first",
          selection: "Data",
          destination: {
            kind: "new-table",
            schema: {
              name: "Shared",
              columns: [{ name: "Value", type: "text" }],
              recordId: { prefix: "SHA", padding: 6 },
            },
          },
          columns: [{ source: "Value", target: "Value", type: "text" }],
        },
        {
          source: "second",
          selection: "Data",
          destination: {
            kind: "new-table",
            schema: {
              name: "Shared",
              columns: [{ name: "Value", type: "integer" }],
              recordId: { prefix: "SHA", padding: 6 },
            },
          },
          columns: [{ source: "Value", target: "Value", type: "integer" }],
        },
      ],
    };
    const { database, prepared } = await fixture(format, recipe);
    try {
      const text = source({
        Data: [
          {
            sourceRow: 2,
            cells: { Value: { kind: "string", value: "A" } },
          },
        ],
      });
      const integer = source({
        Data: [
          {
            sourceRow: 2,
            cells: { Value: { kind: "number", raw: "1" } },
          },
        ],
      });
      const outcome = await prepareImport({
        database,
        prepared,
        recipe,
        sources: [
          { ...text, key: "first" },
          { ...integer, key: "second" },
        ],
      });
      expect(outcome.prepared.state).toBe("needs-review");
      expect(outcome.result.metrics.conflicts).toBeGreaterThan(0);
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: explicit re-review rebases a saved capture without reading its source again`, async () => {
    const firstRecipe: ImportRecipe = {
      version: 1,
      routes: [
        {
          source: "submission",
          selection: "First",
          destination: {
            kind: "new-table",
            schema: {
              name: "First",
              columns: [{ name: "Value", type: "text" }],
              recordId: { prefix: "FIR", padding: 6 },
            },
          },
          columns: [{ source: "Value", target: "Value", type: "text" }],
        },
      ],
    };
    const secondRecipe: ImportRecipe = {
      version: 1,
      routes: [
        {
          source: "submission",
          selection: "Second",
          destination: {
            kind: "new-table",
            schema: {
              name: "Second",
              columns: [{ name: "Value", type: "text" }],
              recordId: { prefix: "SEC", padding: 6 },
            },
          },
          columns: [{ source: "Value", target: "Value", type: "text" }],
        },
      ],
    };
    const directory = await mkdtemp(path.join(tmpdir(), "cc-import-rebase-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const first = await createPreparedImport({
      path: path.join(directory, "first.ccplan"),
      database,
      recipe: firstRecipe,
      baselineRevision: 0n,
    });
    const second = await createPreparedImport({
      path: path.join(directory, "second.ccplan"),
      database,
      recipe: secondRecipe,
      baselineRevision: 0n,
    });
    try {
      const firstOutcome = await prepareImport({
        database,
        prepared: first,
        recipe: firstRecipe,
        sources: [
          source({
            First: [
              {
                sourceRow: 2,
                cells: { Value: { kind: "string", value: "A" } },
              },
            ],
          }),
        ],
      });
      const secondOutcome = await prepareImport({
        database,
        prepared: second,
        recipe: secondRecipe,
        sources: [
          source({
            Second: [
              {
                sourceRow: 2,
                cells: { Value: { kind: "string", value: "B" } },
              },
            ],
          }),
        ],
      });
      if (
        firstOutcome.prepared.state !== "ready" ||
        secondOutcome.prepared.state !== "ready"
      ) {
        throw new Error("Plans not ready");
      }
      await applyImport({
        database,
        prepared: first,
        approved: firstOutcome.prepared,
        requestId: `${format}-first-plan`,
      });
      await expect(
        applyImport({
          database,
          prepared: second,
          approved: secondOutcome.prepared,
          requestId: `${format}-stale-second-plan`,
        }),
      ).rejects.toMatchObject({ code: "DB_STALE_IMPORT_PLAN" });
      const rebased = await resolveImport({
        database,
        prepared: second,
        decisions: [],
        rebase: true,
      });
      expect(rebased).toMatchObject({
        state: "ready",
        baselineRevision: 1n,
      });
      expect(rebased.planRevision).toBeGreaterThan(
        secondOutcome.prepared.planRevision,
      );
      if (rebased.state !== "ready") throw new Error("Rebase failed review");
      await applyImport({
        database,
        prepared: second,
        approved: rebased,
        requestId: `${format}-second-plan`,
      });
      expect((await inspectDatabase({ database })).tables).toHaveLength(2);
    } finally {
      await Promise.all([first.close(), second.close()]);
      await database.close();
    }
  });
}
