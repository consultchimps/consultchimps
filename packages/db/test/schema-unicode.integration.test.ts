import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { engineOf, inspectDatabase } from "../src/database.js";
import { createDatabase, openDatabase } from "../src/node.js";
import { applySchema, planSchema } from "../src/records.js";
import {
  quoteIdentifier,
  validateTableSchema,
  type DatabaseSchema,
  type TableSchema,
} from "../src/schema.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

const ordinaryTable: TableSchema = {
  name: "Ordinary",
  recordId: { prefix: "ROW", padding: 4 },
  columns: [{ name: "Value", type: "text" }],
};

function malformedSchemas(): readonly DatabaseSchema[] {
  const malformedColumn = JSON.parse(
    '{"version":1,"tables":[{"name":"Rows","recordId":{"prefix":"ROW","padding":4},"columns":[{"name":"Bad\\udc00","type":"text"}]}]}',
  ) as DatabaseSchema;
  const malformedReference = JSON.parse(
    '{"version":1,"tables":[{"name":"Children","recordId":{"prefix":"CH","padding":4},"columns":[{"name":"parent_id","type":"text"}],"foreignKeys":[{"column":"parent_id","referencesTable":"Parents\\ud800"}]}]}',
  ) as DatabaseSchema;
  return [
    {
      version: 1,
      tables: [{ ...ordinaryTable, name: "Bad\ud800" }],
    },
    malformedColumn,
    malformedReference,
  ];
}

test("Record ID fields and relationship columns reject unpaired surrogates", () => {
  for (const recordId of [
    { prefix: "ROW\ud800", padding: 4, separator: "\udc00" },
    { prefix: "ROW", padding: 4, separator: "\ud800" },
  ]) {
    expect(() => validateTableSchema({ ...ordinaryTable, recordId })).toThrow(
      expect.objectContaining({ code: "DB_INVALID_RECORD_ID_CONFIG" }),
    );
  }
  expect(() =>
    validateTableSchema({
      ...ordinaryTable,
      foreignKeys: [{ column: "Value\ud800", referencesTable: "Ordinary" }],
    }),
  ).toThrow(expect.objectContaining({ code: "DB_INVALID_IDENTIFIER" }));
});

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format}: malformed Unicode schema names fail without changing the database`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-schema-unicode-"));
    directories.push(directory);
    const databasePath = path.join(directory, `workspace.${format}`);
    const { database } = await createDatabase({
      path: databasePath,
      format,
    });
    const before = await inspectDatabase({ database });
    try {
      for (const schema of malformedSchemas()) {
        await expect(planSchema({ database, schema })).rejects.toMatchObject({
          code: "DB_INVALID_IDENTIFIER",
        });
        expect(await inspectDatabase({ database })).toEqual(before);
      }

      const parent: TableSchema = {
        name: "Parents 🚀",
        recordId: { prefix: "P🚀", padding: 4, separator: "🌟" },
        columns: [{ name: "Name ✨", type: "text" }],
      };
      const child: TableSchema = {
        name: "Children 🐒",
        recordId: { prefix: "C🐒", padding: 4 },
        columns: [{ name: "Parent ID 🔗", type: "text" }],
        foreignKeys: [
          { column: "Parent ID 🔗", referencesTable: "Parents 🚀" },
        ],
      };
      const plan = await planSchema({
        database,
        schema: { version: 1, tables: [child, parent] },
      });
      expect(plan.state).toBe("ready");
      await applySchema({ database, plan });
      await engineOf(database).execute(
        `INSERT INTO ${quoteIdentifier(parent.name)} (${quoteIdentifier("record_id")}, ${quoteIdentifier("Name ✨")}) VALUES (?, ?)`,
        ["P🚀🌟0001", "Synthetic"],
      );
      await engineOf(database).execute(
        `INSERT INTO ${quoteIdentifier(child.name)} (${quoteIdentifier("record_id")}, ${quoteIdentifier("Parent ID 🔗")}) VALUES (?, ?)`,
        ["C🐒-0001", "P🚀🌟0001"],
      );
    } finally {
      await database.close();
    }

    const reopened = await openDatabase({ path: databasePath });
    try {
      const inspection = await inspectDatabase({ database: reopened });
      expect(inspection.tables).toEqual([
        expect.objectContaining({
          name: "Children 🐒",
          rowCount: 1n,
          schema: expect.objectContaining({
            foreignKeys: [
              {
                column: "Parent ID 🔗",
                referencesTable: "Parents 🚀",
              },
            ],
          }),
        }),
        expect.objectContaining({ name: "Parents 🚀", rowCount: 1n }),
      ]);
    } finally {
      await reopened.close();
    }
  });
}
