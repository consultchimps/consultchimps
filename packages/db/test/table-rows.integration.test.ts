import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { afterEach, describe, expect, test } from "vitest";

import { engineOf } from "../src/database.js";
import { createDatabase } from "../src/node.js";
import { readTableRows } from "../src/rows.js";
import type { DatabaseSchema } from "../src/schema.js";

const SCHEMA: DatabaseSchema = {
  version: 1,
  tables: [
    {
      name: "Readings",
      recordId: { prefix: "READ", padding: 1 },
      columns: [
        { name: "label", type: "text", nullable: false },
        { name: "amount", type: "integer" },
        { name: "checked", type: "boolean" },
        { name: "price", type: "decimal", precision: 10, scale: 2 },
        { name: "on_day", type: "date" },
        { name: "at_time", type: "timestamp" },
      ],
    },
  ],
};

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

/**
 * Twelve rows with one-digit and two-digit Record IDs, so an order that
 * compared text alone would put READ-10 before READ-2.
 */
async function seeded(format: "sqlite" | "duckdb") {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-table-rows-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, `rows.${format}`),
    format,
    schema: SCHEMA,
  });
  const engine = engineOf(database);
  for (let index = 1; index <= 12; index += 1) {
    await engine.execute(
      `INSERT INTO "Readings" (record_id, label, amount, checked, price, on_day, at_time, _source_file_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        `READ-${String(index)}`,
        `reading ${String(index)}`,
        BigInt(index * 1000),
        index % 2 === 0,
        `${String(index)}.50`,
        "2026-09-16",
        "2026-09-16 12:00:00",
        index <= 3 ? "SRC-1" : null,
      ],
    );
  }
  return database;
}

describe.each(["sqlite", "duckdb"] as const)(
  "readTableRows on %s",
  (format) => {
    test("pages stored rows in Record ID order with a continuing cursor", async () => {
      const database = await seeded(format);
      try {
        const first = await readTableRows({
          database,
          table: "readings",
          page: { limit: 5 },
        });
        expect(first.table).toBe("Readings");
        expect(first.columns.map((column) => column.name)).toEqual([
          "record_id",
          "label",
          "amount",
          "checked",
          "price",
          "on_day",
          "at_time",
          "_source_file_id",
        ]);
        expect(first.rows.map((row) => row[0])).toEqual([
          "READ-1",
          "READ-2",
          "READ-3",
          "READ-4",
          "READ-5",
        ]);
        expect(first.rows[0]?.slice(1, 4)).toEqual([
          "reading 1",
          expect.anything(),
          false,
        ]);
        expect(first.rows[1]?.[3]).toBe(true);
        expect(first.rows[0]?.[7]).toBe("SRC-1");
        expect(first.rows[4]?.[7]).toBeNull();
        // 64-bit integers arrive as decimal strings or as safe numbers, never
        // as bigint, so the page can cross a structured-clone boundary; the
        // decimal, date, and timestamp columns arrive as the engine renders
        // them, always as plain text or numbers.
        expect(String(first.rows[0]?.[2])).toBe("1000");
        for (const row of first.rows) {
          for (const value of row) {
            expect(["string", "number", "boolean", "object"]).toContain(
              typeof value,
            );
            expect(typeof value).not.toBe("bigint");
          }
          expect(String(row[4])).toMatch(/^\d+\.5/u);
          expect(String(row[5])).toContain("2026-09-16");
          expect(String(row[6])).toContain("2026-09-16");
        }
        expect(first.nextCursor).toBeDefined();

        const second = await readTableRows({
          database,
          table: "Readings",
          page: { limit: 5, cursor: first.nextCursor },
        });
        expect(second.rows.map((row) => row[0])).toEqual([
          "READ-6",
          "READ-7",
          "READ-8",
          "READ-9",
          "READ-10",
        ]);
        const third = await readTableRows({
          database,
          table: "Readings",
          page: { limit: 5, cursor: second.nextCursor },
        });
        expect(third.rows.map((row) => row[0])).toEqual(["READ-11", "READ-12"]);
        expect(third.nextCursor).toBeUndefined();

        const whole = await readTableRows({ database, table: "Readings" });
        expect(whole.rows).toHaveLength(12);
        expect(whole.nextCursor).toBeUndefined();
      } finally {
        await database.close();
      }
    });

    test("refuses an unknown table, a bad page size, and a foreign cursor", async () => {
      const database = await seeded(format);
      try {
        await expect(
          readTableRows({ database, table: "Missing" }),
        ).rejects.toMatchObject({ code: "DB_TABLE_NOT_FOUND" });
        await expect(
          readTableRows({ database, table: "Readings", page: { limit: 0 } }),
        ).rejects.toMatchObject({ code: "DB_INVALID_PAGE_SIZE" });
        await expect(
          readTableRows({ database, table: "Readings", page: { limit: 201 } }),
        ).rejects.toMatchObject({ code: "DB_INVALID_PAGE_SIZE" });
        await expect(
          readTableRows({
            database,
            table: "Readings",
            page: { cursor: "not json" },
          }),
        ).rejects.toMatchObject({ code: "DB_INVALID_CURSOR" });
        await expect(
          readTableRows({
            database,
            table: "Readings",
            page: { cursor: JSON.stringify(["other", "READ-1"]) },
          }),
        ).rejects.toMatchObject({ code: "DB_INVALID_CURSOR" });
      } finally {
        await database.close();
      }
    });

    test("refuses a registered table whose storage was changed outside the tool", async () => {
      const database = await seeded(format);
      try {
        await engineOf(database).execute(`DROP TABLE "Readings"`);
        await expect(
          readTableRows({ database, table: "Readings" }),
        ).rejects.toMatchObject({ code: "DB_SCHEMA_DRIFT" });
      } finally {
        await database.close();
      }
    });
  },
);
