import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { engineOf } from "../src/database.js";
import { createDatabase, exportDatabase, openDatabase } from "../src/node.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

const schema = {
  version: 1 as const,
  tables: [
    {
      name: "values_table",
      recordId: { prefix: "VALUE", padding: 4 },
      columns: [
        { name: "label", type: "text" as const, nullable: false },
        { name: "count", type: "integer" as const, nullable: false },
        { name: "ratio", type: "real" as const, nullable: false },
        {
          name: "amount",
          type: "decimal" as const,
          precision: 12,
          scale: 4,
          nullable: false,
        },
        { name: "active", type: "boolean" as const, nullable: false },
        { name: "observed_on", type: "date" as const, nullable: false },
        {
          name: "observed_at",
          type: "timestamp" as const,
          nullable: false,
        },
      ],
    },
  ],
};

async function sourceDatabase(directory: string) {
  const created = await createDatabase({
    path: path.join(directory, "source.sqlite"),
    format: "sqlite",
    schema,
  });
  await engineOf(created.database).bulkInsert({
    table: "values_table",
    columns: [
      "record_id",
      "_imported_row_id",
      "_source_selection",
      "_source_row",
      "label",
      "count",
      "ratio",
      "amount",
      "active",
      "observed_on",
      "observed_at",
    ],
    rows: [
      [
        "VALUE-0001",
        1n,
        "Data",
        2n,
        "North",
        42n,
        1.25,
        "1.2345e2",
        1n,
        "2024-02-29",
        "2024-02-29T03:04:05.123456Z",
      ],
      [
        "VALUE-0002",
        2n,
        "Data",
        3n,
        "Ancient",
        -42n,
        2.5,
        "0e99999999",
        0n,
        "0001-01-01 (BC)",
        "0001-01-01 (BC) 03:04:05.123456",
      ],
      [
        "VALUE-0003",
        3n,
        "Data",
        4n,
        "Unbounded",
        0n,
        0.5,
        "-0.0000",
        1n,
        "infinity",
        "-infinity",
      ],
    ],
  });
  return created.database;
}

test("converts valid logical values from SQLite through DuckDB and back", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-convert-values-"));
  directories.push(directory);
  const source = await sourceDatabase(directory);
  const duckPath = path.join(directory, "values.duckdb");
  try {
    await exportDatabase({
      database: source,
      output: duckPath,
      format: "duckdb",
    });
  } finally {
    await source.close();
  }

  const duck = await openDatabase({ path: duckPath });
  const roundtripPath = path.join(directory, "roundtrip.sqlite");
  try {
    await expect(
      engineOf(duck).query(
        'SELECT label, "count", ratio, amount::VARCHAR AS amount, active, observed_on::VARCHAR AS observed_on, observed_at::VARCHAR AS observed_at, _source_selection, _source_row FROM values_table',
      ),
    ).resolves.toEqual([
      {
        label: "North",
        count: 42n,
        ratio: 1.25,
        amount: "123.4500",
        active: true,
        observed_on: "2024-02-29",
        observed_at: "2024-02-29 03:04:05.123456",
        _source_selection: "Data",
        _source_row: 2n,
      },
      {
        label: "Ancient",
        count: -42n,
        ratio: 2.5,
        amount: "0.0000",
        active: false,
        observed_on: "0001-01-01 (BC)",
        observed_at: "0001-01-01 (BC) 03:04:05.123456",
        _source_selection: "Data",
        _source_row: 3n,
      },
      {
        label: "Unbounded",
        count: 0n,
        ratio: 0.5,
        amount: "0.0000",
        active: true,
        observed_on: "infinity",
        observed_at: "-infinity",
        _source_selection: "Data",
        _source_row: 4n,
      },
    ]);
    await exportDatabase({
      database: duck,
      output: roundtripPath,
      format: "sqlite",
    });
  } finally {
    await duck.close();
  }

  const roundtrip = await openDatabase({ path: roundtripPath });
  try {
    await expect(
      engineOf(roundtrip).query(
        'SELECT label, "count", ratio, amount, active, observed_on, observed_at, _source_selection, _source_row FROM values_table',
      ),
    ).resolves.toEqual([
      {
        label: "North",
        count: 42n,
        ratio: 1.25,
        amount: "123.4500",
        active: 1n,
        observed_on: "2024-02-29",
        observed_at: "2024-02-29 03:04:05.123456",
        _source_selection: "Data",
        _source_row: 2n,
      },
      {
        label: "Ancient",
        count: -42n,
        ratio: 2.5,
        amount: "0.0000",
        active: 0n,
        observed_on: "0001-01-01 (BC)",
        observed_at: "0001-01-01 (BC) 03:04:05.123456",
        _source_selection: "Data",
        _source_row: 3n,
      },
      {
        label: "Unbounded",
        count: 0n,
        ratio: 0.5,
        amount: "0.0000",
        active: 1n,
        observed_on: "infinity",
        observed_at: "-infinity",
        _source_selection: "Data",
        _source_row: 4n,
      },
    ]);
  } finally {
    await roundtrip.close();
  }
});

const invalidValues = [
  {
    column: "active",
    type: "boolean",
    sql: "UPDATE values_table SET active = 2",
  },
  {
    column: "count",
    type: "integer",
    sql: 'UPDATE values_table SET "count" = 1.5',
  },
  {
    column: "ratio",
    type: "real",
    sql: "UPDATE values_table SET ratio = 'not-a-real'",
  },
  {
    column: "amount",
    type: "decimal",
    sql: "UPDATE values_table SET amount = '12.34567'",
  },
  {
    column: "observed_on",
    type: "date",
    sql: "UPDATE values_table SET observed_on = '2024-02-30'",
  },
  {
    column: "observed_at",
    type: "timestamp",
    sql: "UPDATE values_table SET observed_at = '2024-02-29T03:04:05+04:00'",
  },
  {
    column: "label",
    type: "text",
    sql: "UPDATE values_table SET label = CAST('North' AS BLOB)",
  },
  {
    column: "_source_row",
    type: "integer",
    sql: "UPDATE values_table SET _source_row = 2.5",
  },
] as const;

for (const invalid of invalidValues) {
  test(`rejects an off-type SQLite ${invalid.column} value without replacing the output`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-convert-invalid-"));
    directories.push(directory);
    const source = await sourceDatabase(directory);
    const output = path.join(directory, "preserved.duckdb");
    const original = Buffer.from("existing destination");
    await writeFile(output, original);
    try {
      await engineOf(source).execute(invalid.sql);
      const sourceRows = await engineOf(source).query(
        "SELECT * FROM values_table ORDER BY record_id",
      );
      await expect(
        exportDatabase({
          database: source,
          output,
          format: "duckdb",
          overwrite: true,
        }),
      ).rejects.toMatchObject({
        code: "DB_CONVERSION_INVALID_VALUE",
        details: {
          table: "values_table",
          column: invalid.column,
          type: invalid.type,
        },
      });
      expect((await readFile(output)).equals(original)).toBe(true);
      await expect(
        engineOf(source).query("SELECT * FROM values_table ORDER BY record_id"),
      ).resolves.toEqual(sourceRows);
    } finally {
      await source.close();
    }
  });
}

test("does not publish an earlier valid page when a later value is invalid", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-convert-page-"));
  directories.push(directory);
  const { database: source } = await createDatabase({
    path: path.join(directory, "source.sqlite"),
    format: "sqlite",
    schema: {
      version: 1,
      tables: [
        {
          name: "flags",
          recordId: { prefix: "FLAG", padding: 4 },
          columns: [{ name: "active", type: "boolean", nullable: false }],
        },
      ],
    },
  });
  const rows = Array.from(
    { length: 2_001 },
    (_, index) =>
      [
        `FLAG-${String(index + 1).padStart(4, "0")}`,
        BigInt(index + 1),
        index === 2_000 ? 2n : 1n,
      ] as const,
  );
  await engineOf(source).bulkInsert({
    table: "flags",
    columns: ["record_id", "_imported_row_id", "active"],
    rows,
  });
  const output = path.join(directory, "preserved.duckdb");
  const original = Buffer.from("existing destination");
  await writeFile(output, original);
  try {
    await expect(
      exportDatabase({
        database: source,
        output,
        format: "duckdb",
        overwrite: true,
      }),
    ).rejects.toMatchObject({
      code: "DB_CONVERSION_INVALID_VALUE",
      details: { table: "flags", column: "active", type: "boolean" },
    });
    expect((await readFile(output)).equals(original)).toBe(true);
    await expect(
      engineOf(source).query(
        "SELECT count(*) AS row_count, sum(active) AS active_total FROM flags",
      ),
    ).resolves.toEqual([{ row_count: 2_001n, active_total: 2_002n }]);
  } finally {
    await source.close();
  }
});
