import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { executeConversion, planConversion } from "../src/conversion.js";
import { engineOf } from "../src/database.js";
import { createDatabase } from "../src/node.js";
import { applySchema, planSchema } from "../src/records.js";

const directories: string[] = [];

for (const format of ["sqlite", "duckdb"] as const) {
  test.each(["check", "unique", "default", "missing-primary-key"])(
    `${format}: refuses externally changed %s constraints`,
    async (change) => {
      const directory = await mkdtemp(
        path.join(tmpdir(), "cc-convert-constraint-"),
      );
      directories.push(directory);
      const { database } = await createDatabase({
        path: path.join(directory, `source.${format}`),
        format,
      });
      try {
        const plan = await planSchema({
          database,
          schema: {
            version: 1,
            tables: [
              {
                name: "items",
                columns: [{ name: "value", type: "text" }],
                recordId: { prefix: "ITEM", padding: 4 },
              },
            ],
          },
        });
        await applySchema({ database, plan });
        const engine = engineOf(database);
        await engine.execute("DROP TABLE items");
        await engine.execute(`CREATE TABLE items (
        record_id VARCHAR NOT NULL ${change === "missing-primary-key" ? "" : "PRIMARY KEY"},
        _imported_row_id BIGINT ${format === "sqlite" ? "UNIQUE" : ""},
        _import_id VARCHAR, _source_file_id VARCHAR, _source_selection VARCHAR,
        _source_row BIGINT, value VARCHAR ${change === "default" ? "DEFAULT 'pending'" : ""}
        ${change === "check" ? ", CHECK(value <> 'forbidden')" : ""}
        ${change === "unique" ? ", UNIQUE(value)" : ""}
      )`);
        await expect(
          planConversion({
            database,
            format: format === "sqlite" ? "duckdb" : "sqlite",
          }),
        ).rejects.toMatchObject({ code: "DB_SCHEMA_DRIFT" });
      } finally {
        await database.close();
      }
    },
  );
}

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

test("reports DuckDB macros and custom types before cross-format conversion", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-convert-objects-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "source.duckdb"),
    format: "duckdb",
  });
  try {
    const engine = engineOf(database);
    await engine.execute("CREATE MACRO add_one(value) AS value + 1");
    await engine.execute("CREATE TYPE mood AS ENUM ('calm', 'busy')");
    const plan = await planConversion({ database, format: "sqlite" });
    expect(plan.state).toBe("unsupported");
    expect(plan.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          kind: "unsupported-object",
          name: "add_one",
          objectType: "macro",
        }),
        expect.objectContaining({
          kind: "unsupported-object",
          name: "mood",
          objectType: "type",
        }),
      ]),
    );
  } finally {
    await database.close();
  }
});

test("rejects a conversion when an unsupported object appears after planning", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-convert-stale-"));
  directories.push(directory);
  const { database: source } = await createDatabase({
    path: path.join(directory, "source.duckdb"),
    format: "duckdb",
  });
  const { database: target } = await createDatabase({
    path: path.join(directory, "target.sqlite"),
    format: "sqlite",
  });
  try {
    const plan = await planConversion({ database: source, format: "sqlite" });
    expect(plan.state).toBe("ready");
    await engineOf(source).execute(
      "CREATE MACRO added_after_plan(value) AS value + 1",
    );
    await expect(
      executeConversion({ source, target, plan }),
    ).rejects.toMatchObject({ code: "DB_STALE_CONVERSION_PLAN" });
  } finally {
    await target.close();
    await source.close();
  }
});
