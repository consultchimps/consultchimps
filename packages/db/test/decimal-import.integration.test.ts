import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import {
  applyImport,
  inspectImport,
  prepareImport,
  resolveImport,
} from "../src/import/operations.js";
import { draftImportRecipe } from "../src/import/recipe.js";
import type { ImportSource } from "../src/import/types.js";
import { createDatabase, createPreparedImport } from "../src/node.js";
import { engineOf, inspectDatabase } from "../src/database.js";

const directories: string[] = [];
afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

test("infers decimal capacity from integer digits and scale", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-decimal-import-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "decimal.sqlite"),
    format: "sqlite",
  });
  const bytes = new TextEncoder().encode("decimal inference source");
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
            columns: ["Amount"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: { Amount: { kind: "number", raw: "10000" } },
                },
                {
                  sourceRow: 3,
                  cells: { Amount: { kind: "number", raw: "0.123" } },
                },
              ];
            },
            async close() {},
          };
        },
      },
    ],
  };
  const recipe = await draftImportRecipe({ sources: [source] });
  const prepared = await createPreparedImport({
    path: path.join(directory, "decimal.ccplan"),
    database,
    recipe,
    baselineRevision: 0n,
  });
  try {
    await prepareImport({ database, prepared, sources: [source], recipe });
    const inspection = await inspectImport({ prepared, page: { limit: 1 } });
    expect(inspection.routes[0]?.inferredColumns).toEqual([
      { name: "Amount", type: "decimal", precision: 8, scale: 3 },
    ]);
    const narrowed = await resolveImport({
      database,
      prepared,
      decisions: [
        {
          kind: "route",
          source: "amounts",
          selection: "amounts",
          destination: {
            kind: "new-table",
            schema: {
              name: "amounts",
              recordId: { prefix: "AMOUNT", padding: 6 },
              columns: [
                {
                  name: "Amount",
                  type: "decimal",
                  precision: 7,
                  scale: 2,
                },
              ],
            },
          },
          columns: [{ source: "Amount", target: "Amount", type: "decimal" }],
        },
      ],
    });
    expect(narrowed.state).toBe("needs-review");
    expect(
      (await inspectImport({ prepared, page: { limit: 1 } })).conflicts,
    ).toEqual([expect.objectContaining({ kind: "decimal-capacity" })]);
  } finally {
    await prepared.close();
    await database.close();
  }
});

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format}: exponent-form numbers infer and apply as exact decimals`, async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "cc-decimal-exponent-"),
    );
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `decimal.${format}`),
      format,
    });
    const bytes = new TextEncoder().encode("decimal exponent source");
    const source: ImportSource = {
      key: "amounts",
      readerVersion: "synthetic-exponent-1",
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
              columns: ["Amount"],
              async *batches() {
                yield ["1.25E-2", "-4E+1", "0E+9"].map((raw, index) => ({
                  sourceRow: index + 2,
                  cells: { Amount: { kind: "number" as const, raw } },
                }));
              },
              async close() {},
            };
          },
        },
      ],
    };
    const draft = await draftImportRecipe({ sources: [source] });
    const prepared = await createPreparedImport({
      path: path.join(directory, "decimal.ccplan"),
      database,
      recipe: draft,
      baselineRevision: (await inspectDatabase({ database })).revision,
    });
    try {
      await prepareImport({
        database,
        prepared,
        sources: [source],
        recipe: draft,
      });
      const inspection = await inspectImport({ prepared, page: { limit: 1 } });
      expect(inspection.routes[0]?.inferredColumns).toEqual([
        { name: "Amount", type: "decimal", precision: 6, scale: 4 },
      ]);
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [
          {
            kind: "route",
            source: "amounts",
            selection: "amounts",
            destination: {
              kind: "new-table",
              schema: {
                name: "amounts",
                recordId: { prefix: "AMOUNT", padding: 6 },
                columns: [
                  {
                    name: "Amount",
                    type: "decimal",
                    precision: 6,
                    scale: 4,
                  },
                ],
              },
            },
            columns: [{ source: "Amount", target: "Amount", type: "decimal" }],
          },
        ],
      });
      expect(approved.state).toBe("ready");
      if (approved.state !== "ready") throw new Error("Plan was not ready");
      await applyImport({
        database,
        prepared,
        approved,
        requestId: `decimal-exponent-${format}`,
      });
      expect(
        await engineOf(database).query(
          'SELECT CAST("Amount" AS VARCHAR) AS value FROM "amounts" ORDER BY record_id',
        ),
      ).toEqual([
        { value: "0.0125" },
        { value: "-40.0000" },
        { value: "0.0000" },
      ]);
    } finally {
      await prepared.close();
      await database.close();
    }
  });
}
