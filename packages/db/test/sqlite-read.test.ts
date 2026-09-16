import { beforeAll, describe, expect, it, vi } from "vitest";
import initialize from "@sqlite.org/sqlite-wasm";
import { createRequire } from "node:module";
import { readFile } from "node:fs/promises";
import { openReadOnlySqlite } from "../src/sqlite-read.js";
import type { SqliteReadOptions } from "../src/sqlite-read.js";

let bytes: Uint8Array;
const secret = "SYNTHETIC_PRIVATE_CATALOG";

beforeAll(async () => {
  const sqlite = await initialize();
  const database = new sqlite.oo1.DB(":memory:");
  try {
    database.exec(
      `CREATE TABLE "${secret}" (id INTEGER, label TEXT, payload BLOB, value REAL, missing TEXT); INSERT INTO "${secret}" VALUES(9223372036854775807, 'North', x'0001ff', 1.25, NULL), (-9223372036854775808, 'South', x'', 2.5, NULL);`,
    );
    bytes = sqlite.capi.sqlite3_js_db_export(database.pointer!);
  } finally {
    database.close();
  }
});

function expectPrivateError(error: unknown, code: string): void {
  expect(error).toMatchObject({ code, cause: undefined });
  expect(String(error)).not.toContain(secret);
  expect(JSON.stringify(error)).not.toContain(secret);
}

describe("read-only SQLite bytes", () => {
  it("reads independent typed rows, preserving int64, nulls, blobs and input bytes", async () => {
    const original = bytes.slice();
    const reader = await openReadOnlySqlite(bytes);
    try {
      const result = reader.query(`SELECT * FROM "${secret}" ORDER BY label`);
      expect(result.columns).toEqual([
        "id",
        "label",
        "payload",
        "value",
        "missing",
      ]);
      expect(result.rows).toEqual([
        [
          9223372036854775807n,
          "North",
          new Uint8Array([0, 1, 255]),
          1.25,
          null,
        ],
        [-9223372036854775808n, "South", new Uint8Array(), 2.5, null],
      ]);
      expect(result.resultBytes).toBeGreaterThan(0);
      (result.rows[0]![2] as Uint8Array).fill(0);
      expect(
        reader.query(`SELECT payload FROM "${secret}" WHERE label = ?`, [
          "North",
        ]).rows,
      ).toEqual([[new Uint8Array([0, 1, 255])]]);
      expect(bytes).toEqual(original);
      expect(reader.wasmMemoryBytes).toBeGreaterThan(0);
    } finally {
      reader.close();
    }
    expect(reader.wasmMemoryBytes).toBe(0);
    reader.close();
    expect(() => reader.query("SELECT 1")).toThrow(
      expect.objectContaining({ code: "DB_SQLITE_READ_CLOSED" }),
    );
  });

  it("binds scalar values without truncating Unicode or embedded NULs", async () => {
    const reader = await openReadOnlySqlite(bytes);
    try {
      const values = [
        null,
        "North; 南\0suffix",
        1.25,
        9223372036854775807n,
        new Uint8Array([0, 255]),
        new Uint8Array(),
      ];
      expect(
        reader.query("SELECT ?, ?, ?, ?, ?, ?; -- trailing comment", values)
          .rows,
      ).toEqual([values]);
    } finally {
      reader.close();
    }
  });

  it.each([
    `DELETE FROM "${secret}"`,
    "PRAGMA query_only = OFF",
    "ATTACH ':memory:' AS other",
    "CREATE TEMP TABLE other (x)",
    "VACUUM INTO 'other.sqlite'",
    "BEGIN",
  ])(
    "refuses state-changing SQL and keeps the reader usable: %s",
    async (sql) => {
      const reader = await openReadOnlySqlite(bytes);
      const log = vi
        .spyOn(console, "error")
        .mockImplementation(() => undefined);
      try {
        expect(() => reader.query(sql)).toThrow(
          expect.objectContaining({ code: "DB_SQLITE_READ_ONLY" }),
        );
        expect(reader.query(`SELECT count(*) FROM "${secret}"`).rows).toEqual([
          [2n],
        ]);
        expect(log).not.toHaveBeenCalled();
      } finally {
        log.mockRestore();
        reader.close();
      }
    },
  );

  it("refuses multiple statements, invalid parameters, and private parser diagnostics", async () => {
    const reader = await openReadOnlySqlite(bytes);
    try {
      for (const sql of [
        "SELECT 1; SELECT 2",
        `SELECT broken_${secret}`,
        "",
        "SELECT 1\0",
        "SELECT ?",
      ]) {
        try {
          reader.query(sql);
          throw new Error("Expected refusal");
        } catch (error) {
          expectPrivateError(error, "DB_SQLITE_READ_QUERY_FAILED");
        }
      }
      for (const value of [Infinity, NaN, 1n << 63n]) {
        expect(() => reader.query("SELECT ?", [value])).toThrow();
      }
      expect(reader.query("SELECT ';'").rows).toEqual([[";"]]);
    } finally {
      reader.close();
    }
  });

  it("enforces inclusive input and result bounds", async () => {
    await expect(
      openReadOnlySqlite(bytes, { maxDatabaseBytes: bytes.length - 1 }),
    ).rejects.toMatchObject({ code: "DB_SQLITE_READ_LIMIT_EXCEEDED" });
    const first = await openReadOnlySqlite(bytes, {
      maxDatabaseBytes: bytes.length,
    });
    const sql = `SELECT label FROM "${secret}" ORDER BY label`;
    const required = first.query(sql).resultBytes;
    first.close();
    const exact = await openReadOnlySqlite(bytes, {
      maxRows: 2,
      maxResultBytes: required,
    });
    try {
      expect(exact.query(sql).resultBytes).toBe(required);
    } finally {
      exact.close();
    }
    const limited = await openReadOnlySqlite(bytes, {
      maxResultBytes: required - 1,
    });
    try {
      expect(() => limited.query(sql)).toThrow(
        expect.objectContaining({ code: "DB_SQLITE_READ_LIMIT_EXCEEDED" }),
      );
    } finally {
      limited.close();
    }
    const one = await openReadOnlySqlite(bytes, { maxRows: 1 });
    try {
      expect(() => one.query(sql)).toThrow(
        expect.objectContaining({ details: { option: "maxRows", limit: 1 } }),
      );
      expect(one.query("SELECT 1").rows).toEqual([[1n]]);
    } finally {
      one.close();
    }
  });

  it("interrupts runaway SQL and bounds SQL and engine allocations", async () => {
    const reader = await openReadOnlySqlite(bytes, {
      maxSteps: 1000,
      maxSqlBytes: 1024,
      maxSqliteBytes: 1024 * 1024,
    });
    try {
      expect(() =>
        reader.query(
          "WITH RECURSIVE x(n) AS (VALUES(1) UNION ALL SELECT n+1 FROM x) SELECT sum(n) FROM x",
        ),
      ).toThrow(
        expect.objectContaining({
          details: { option: "maxSteps", limit: 1000 },
        }),
      );
      expect(() => reader.query(`SELECT '${"x".repeat(1024)}'`)).toThrow(
        expect.objectContaining({
          details: { option: "maxSqlBytes", limit: 1024 },
        }),
      );
      expect(() => reader.query("SELECT zeroblob(2000000)")).toThrow(
        expect.objectContaining({ code: "DB_SQLITE_READ_LIMIT_EXCEEDED" }),
      );
      expect(reader.query("SELECT 1").rows).toEqual([[1n]]);
    } finally {
      reader.close();
    }
  });

  it("validates options before reading input or calling a locator", async () => {
    const locator = vi.fn(() => secret);
    const input = new Proxy(new Uint8Array(), {
      get() {
        throw new Error("Input accessed");
      },
    });
    await expect(
      openReadOnlySqlite(input, {
        maxRows: 0,
        maxSqlBytes: NaN,
        runtime: { locateFile: locator, wasmBinary: new Uint8Array([1]) },
      }),
    ).rejects.toMatchObject({
      code: "DB_SQLITE_READ_INVALID_OPTIONS",
      details: {
        invalidOptions: [
          {
            path: "maxSqlBytes",
            requirement: "a positive integer no greater than 2147483647",
          },
          {
            path: "maxRows",
            requirement: "a positive integer no greater than 2147483647",
          },
          {
            path: "runtime",
            requirement: "exactly one locator function or nonempty Uint8Array",
          },
        ],
      },
    });
    expect(locator).not.toHaveBeenCalled();
    for (const runtime of [
      null,
      {},
      { wasmBinary: new Uint8Array() },
      { locateFile: secret },
    ]) {
      await expect(
        openReadOnlySqlite(input, { runtime } as SqliteReadOptions),
      ).rejects.toMatchObject({ code: "DB_SQLITE_READ_INVALID_OPTIONS" });
    }
  });

  it("isolates successful and failed custom initialization in the same process", async () => {
    const require = createRequire(import.meta.url);
    const wasmPath = require.resolve("@sqlite.org/sqlite-wasm/sqlite3.wasm");
    const binary = new Uint8Array(await readFile(wasmPath));
    const first = await openReadOnlySqlite(bytes, {
      runtime: { wasmBinary: binary },
    });
    try {
      const locator = vi.fn(() => wasmPath);
      const second = await openReadOnlySqlite(bytes, {
        runtime: { locateFile: locator },
      });
      second.close();
      expect(locator).toHaveBeenCalledWith("sqlite3.wasm");
      await expect(
        openReadOnlySqlite(bytes, {
          runtime: { wasmBinary: new Uint8Array([1, 2]) },
        }),
      ).rejects.toMatchObject({
        code: "DB_SQLITE_READ_RUNTIME_UNAVAILABLE",
        cause: undefined,
      });
      const third = await openReadOnlySqlite(bytes);
      third.close();
      expect(first.query("SELECT 1").rows).toEqual([[1n]]);
    } finally {
      first.close();
    }
  });

  it("redacts locator failures and refuses corrupt SQLite bytes", async () => {
    await expect(
      openReadOnlySqlite(bytes, {
        runtime: {
          locateFile() {
            throw new Error(secret);
          },
        },
      }),
    ).rejects.toMatchObject({
      code: "DB_SQLITE_READ_RUNTIME_UNAVAILABLE",
      cause: undefined,
    });
    for (const input of [new Uint8Array(), bytes.slice(0, 100)]) {
      await expect(openReadOnlySqlite(input)).rejects.toMatchObject({
        code: "DB_SQLITE_READ_INVALID_DATABASE",
        cause: undefined,
      });
    }
  });

  it("refuses inputs that are not standalone rollback-journal databases", async () => {
    const pageSize = (bytes[16]! << 8) | bytes[17]!;
    const wal = bytes.slice();
    wal[18] = 2;
    wal[19] = 2;
    for (const input of [
      "not bytes" as unknown as Uint8Array,
      new ArrayBuffer(bytes.length) as unknown as Uint8Array,
      wal,
      bytes.slice(0, pageSize),
      bytes.slice(0, Math.floor(bytes.length / 2)),
    ]) {
      await expect(openReadOnlySqlite(input)).rejects.toMatchObject({
        code: "DB_SQLITE_READ_INVALID_DATABASE",
        cause: undefined,
      });
    }
  });

  it("reads an offset view or Buffer without touching bytes outside the view", async () => {
    const backing = new Uint8Array(bytes.length + 40).fill(0xff);
    backing.set(bytes, 20);
    const view = backing.subarray(20, 20 + bytes.length);
    for (const input of [view, Buffer.from(view)]) {
      const reader = await openReadOnlySqlite(input, {
        maxDatabaseBytes: bytes.length,
      });
      try {
        expect(reader.query(`SELECT count(*) FROM "${secret}"`).rows).toEqual([
          [2n],
        ]);
      } finally {
        reader.close();
      }
    }
    expect(backing.subarray(0, 20)).toEqual(new Uint8Array(20).fill(0xff));
    expect(backing.subarray(20 + bytes.length)).toEqual(
      new Uint8Array(20).fill(0xff),
    );
  });

  it("refuses comment-only SQL and schema PRAGMA table functions", async () => {
    const reader = await openReadOnlySqlite(bytes);
    try {
      for (const sql of ["-- nothing", "/* nothing */", ";"]) {
        expect(() => reader.query(sql)).toThrow(
          expect.objectContaining({ code: "DB_SQLITE_READ_QUERY_FAILED" }),
        );
      }
      expect(() =>
        reader.query(`SELECT name FROM pragma_table_info('${secret}')`),
      ).toThrow(expect.objectContaining({ code: "DB_SQLITE_READ_ONLY" }));
      expect(
        reader.query("SELECT name FROM sqlite_schema WHERE type = ?", ["table"])
          .rows,
      ).toEqual([[secret]]);
    } finally {
      reader.close();
    }
  });

  it("opens under tight query limits and applies them to the caller's queries only", async () => {
    const reader = await openReadOnlySqlite(bytes, {
      maxSqlBytes: 40,
      maxResultBytes: 64,
      maxSteps: 1,
    });
    try {
      expect(() => reader.query("SELECT 1")).toThrow(
        expect.objectContaining({ details: { option: "maxSteps", limit: 1 } }),
      );
    } finally {
      reader.close();
    }
    const roomy = await openReadOnlySqlite(bytes, { maxSqlBytes: 40 });
    try {
      expect(roomy.query("SELECT 1").rows).toEqual([[1n]]);
      expect(() => roomy.query(`SELECT 1 -- ${"x".repeat(40)}`)).toThrow(
        expect.objectContaining({
          details: { option: "maxSqlBytes", limit: 40 },
        }),
      );
    } finally {
      roomy.close();
    }
  });

  it("refuses EXPLAIN and reads virtual tables through their internal PRAGMA", async () => {
    const sqlite = await initialize();
    const database = new sqlite.oo1.DB(":memory:");
    let searchable: Uint8Array;
    try {
      database.exec(
        `CREATE VIRTUAL TABLE docs USING fts5(body); INSERT INTO docs VALUES('north wind'), ('south wind'); CREATE VIRTUAL TABLE boxes USING rtree(id, minX, maxX); INSERT INTO boxes VALUES(1, 0, 10);`,
      );
      searchable = sqlite.capi.sqlite3_js_db_export(database.pointer!);
    } finally {
      database.close();
    }
    // A small SQL budget must not reach the statements FTS5 prepares itself.
    const reader = await openReadOnlySqlite(searchable, { maxSqlBytes: 64 });
    try {
      expect(() =>
        reader.query(`SELECT body FROM docs -- ${"x".repeat(64)}`),
      ).toThrow(
        expect.objectContaining({
          details: { option: "maxSqlBytes", limit: 64 },
        }),
      );
      // The internal call names the schema, so the qualified query form and
      // any letter case are indistinguishable from it; assignment is not.
      for (const sql of ["PRAGMA Data_Version", "PRAGMA main.data_version"]) {
        expect(reader.query(sql).rows).toHaveLength(1);
      }
      expect(() => reader.query("PRAGMA data_version = 5")).toThrow(
        expect.objectContaining({ code: "DB_SQLITE_READ_ONLY" }),
      );
      expect(reader.query("SELECT body FROM docs ORDER BY body").rows).toEqual([
        ["north wind"],
        ["south wind"],
      ]);
      expect(
        reader.query("SELECT body FROM docs WHERE docs MATCH ?", ["south"])
          .rows,
      ).toEqual([["south wind"]]);
      // R*Tree cursors prepare writes to their shadow tables, so the module is
      // refused as a whole rather than granted write authorization.
      expect(() => reader.query("SELECT id FROM boxes WHERE maxX > 5")).toThrow(
        expect.objectContaining({ code: "DB_SQLITE_READ_ONLY" }),
      );
      // The authorizer cannot tell a cursor's internal call from a direct one,
      // so this single read-only PRAGMA is permitted; every other one is not.
      expect(reader.query("PRAGMA data_version").rows).toHaveLength(1);
      for (const sql of ["PRAGMA page_size", "PRAGMA table_info(docs)"]) {
        expect(() => reader.query(sql)).toThrow(
          expect.objectContaining({ code: "DB_SQLITE_READ_ONLY" }),
        );
      }
      for (const sql of ["EXPLAIN SELECT 1", "EXPLAIN QUERY PLAN SELECT 1"]) {
        expect(() => reader.query(sql)).toThrow(
          expect.objectContaining({ code: "DB_SQLITE_READ_QUERY_FAILED" }),
        );
      }
      expect(reader.query("VALUES (1), (2)").rows).toEqual([[1n], [2n]]);
    } finally {
      reader.close();
    }
  });

  it("keeps concurrently opened readers isolated", async () => {
    // Three megabytes exceed the small allocator ceiling and stay under the
    // default value-length and result limits.
    const heavy = "SELECT zeroblob(3000000)";
    const readers = await Promise.all([
      openReadOnlySqlite(bytes, { maxSqliteBytes: 1024 * 1024 }),
      openReadOnlySqlite(bytes),
      openReadOnlySqlite(bytes, { maxSqliteBytes: 1024 * 1024 }),
      openReadOnlySqlite(bytes),
    ]);
    try {
      for (const [index, reader] of readers.entries()) {
        expect(reader.query(`SELECT count(*) FROM "${secret}"`).rows).toEqual([
          [2n],
        ]);
        if (index % 2 === 0) {
          expect(() => reader.query(heavy)).toThrow(
            expect.objectContaining({
              details: { option: "maxSqliteBytes", limit: 1024 * 1024 },
            }),
          );
        } else {
          expect(reader.query(heavy).rows[0]![0]).toHaveLength(3000000);
        }
      }
    } finally {
      for (const reader of readers) reader.close();
    }
  });
});
