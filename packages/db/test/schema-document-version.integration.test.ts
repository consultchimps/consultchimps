import { mkdtemp, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, describe, expect, it } from "vitest";

import { inspectDatabase } from "../src/database.js";
import { createDatabase, openDatabase } from "../src/node.js";
import { applySchema, planSchema } from "../src/records.js";
import type { DatabaseFormat, DatabaseSchema } from "../src/schema.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

function schema(table: string): DatabaseSchema {
  return {
    version: 1,
    tables: [
      {
        name: table,
        recordId: { prefix: "ROW", padding: 4 },
        columns: [{ name: "value", type: "text" }],
      },
    ],
  };
}

function unsupportedSchema(version: unknown, table: string): DatabaseSchema {
  return { ...schema(table), version } as unknown as DatabaseSchema;
}

async function temporaryDatabasePath(format: DatabaseFormat): Promise<string> {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-schema-version-"));
  directories.push(directory);
  return path.join(directory, `workspace.${format}`);
}

async function expectAbsent(filePath: string): Promise<void> {
  await expect(stat(filePath)).rejects.toMatchObject({ code: "ENOENT" });
}

for (const format of ["sqlite", "duckdb"] as const) {
  describe(`${format} schema document versions`, () => {
    it.each([2, "1"])(
      "rejects unsupported version %j without publishing a database",
      async (version) => {
        const filePath = await temporaryDatabasePath(format);

        await expect(
          createDatabase({
            path: filePath,
            format,
            schema: unsupportedSchema(version, "rejected_rows"),
          }),
        ).rejects.toMatchObject({ code: "DB_INVALID_SCHEMA_DOCUMENT" });
        await expectAbsent(filePath);
      },
    );

    it("preserves an existing database when overwrite receives an unsupported schema version", async () => {
      const filePath = await temporaryDatabasePath(format);
      const created = await createDatabase({
        path: filePath,
        format,
        schema: schema("existing_rows"),
      });
      await created.database.close();

      await expect(
        createDatabase({
          path: filePath,
          format,
          overwrite: true,
          schema: unsupportedSchema(2, "replacement_rows"),
        }),
      ).rejects.toMatchObject({ code: "DB_INVALID_SCHEMA_DOCUMENT" });

      const reopened = await openDatabase({ path: filePath });
      try {
        await expect(
          inspectDatabase({ database: reopened }),
        ).resolves.toMatchObject({
          revision: 1n,
          tables: [{ schema: { name: "existing_rows" } }],
        });
      } finally {
        await reopened.close();
      }
    });

    it("rejects an unsupported version before planning changes to an open database", async () => {
      const filePath = await temporaryDatabasePath(format);
      const { database } = await createDatabase({
        path: filePath,
        format,
        schema: schema("existing_rows"),
      });
      try {
        const before = await inspectDatabase({ database });

        await expect(
          planSchema({
            database,
            schema: unsupportedSchema(2, "rejected_rows"),
          }),
        ).rejects.toMatchObject({ code: "DB_INVALID_SCHEMA_DOCUMENT" });

        await expect(inspectDatabase({ database })).resolves.toEqual(before);
      } finally {
        await database.close();
      }
    });

    it("adds a table referencing an existing table", async () => {
      const filePath = await temporaryDatabasePath(format);
      const { database } = await createDatabase({
        path: filePath,
        format,
        schema: schema("existing_rows"),
      });
      try {
        const plan = await planSchema({
          database,
          schema: {
            version: 1,
            tables: [
              {
                name: "related_rows",
                recordId: { prefix: "REL", padding: 4 },
                columns: [{ name: "existing_id", type: "text" }],
                foreignKeys: [
                  { column: "existing_id", referencesTable: "existing_rows" },
                ],
              },
            ],
          },
        });
        await applySchema({ database, plan });
        const inspection = await inspectDatabase({ database });
        expect(inspection.tables.map((table) => table.name)).toEqual([
          "existing_rows",
          "related_rows",
        ]);
        expect(inspection.tables[1]?.schema.foreignKeys).toEqual([
          { column: "existing_id", referencesTable: "existing_rows" },
        ]);
      } finally {
        await database.close();
      }
    });

    it("accepts the supported schema version", async () => {
      const filePath = await temporaryDatabasePath(format);
      const { database } = await createDatabase({
        path: filePath,
        format,
        schema: schema("supported_rows"),
      });
      try {
        await expect(inspectDatabase({ database })).resolves.toMatchObject({
          revision: 1n,
          tables: [{ schema: { name: "supported_rows" } }],
        });
      } finally {
        await database.close();
      }
    });
  });
}
