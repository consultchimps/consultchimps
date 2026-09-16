import { describe, expect, it, vi } from "vitest";
import type { Sqlite3Static } from "@sqlite.org/sqlite-wasm";
import type * as SqliteWasm from "@sqlite.org/sqlite-wasm";
import initialize from "@sqlite.org/sqlite-wasm";
import { openReadOnlySqlite } from "../src/sqlite-read.js";

const observed = vi.hoisted(() => ({ runtimes: [] as Sqlite3Static[] }));
vi.mock("@sqlite.org/sqlite-wasm", async (importOriginal) => {
  const actual = await importOriginal<typeof SqliteWasm>();
  return {
    ...actual,
    default: async (...args: unknown[]) => {
      const sqlite = (await Reflect.apply(
        actual.default,
        undefined,
        args,
      )) as Sqlite3Static;
      vi.spyOn(sqlite.wasm, "allocFromTypedArray");
      vi.spyOn(sqlite.wasm, "dealloc");
      vi.spyOn(sqlite.capi, "sqlite3_finalize");
      observed.runtimes.push(sqlite);
      return sqlite;
    },
  };
});

async function fixture(): Promise<Uint8Array> {
  const sqlite = await initialize();
  const database = new sqlite.oo1.DB(":memory:");
  database.exec(
    "CREATE TABLE example (id INTEGER); INSERT INTO example VALUES(1)",
  );
  try {
    return sqlite.capi.sqlite3_js_db_export(database.pointer!);
  } finally {
    database.close();
  }
}

describe("SQLite reader ownership", () => {
  it("finalizes refused statements and frees the database buffer once on close", async () => {
    const reader = await openReadOnlySqlite(await fixture());
    const sqlite = observed.runtimes.at(-1)!;
    const pointer = vi.mocked(sqlite.wasm.allocFromTypedArray).mock.results[0]!
      .value as number;
    const finalized = vi.mocked(sqlite.capi.sqlite3_finalize).mock.calls.length;
    expect(() => reader.query("SELECT 1; SELECT 2")).toThrow();
    expect(vi.mocked(sqlite.capi.sqlite3_finalize).mock.calls.length).toBe(
      finalized + 2,
    );
    expect(reader.query("SELECT id FROM example").rows).toEqual([[1n]]);
    const callsBeforeClose = vi.mocked(sqlite.wasm.dealloc).mock.calls.length;
    reader.close();
    expect(
      vi.mocked(sqlite.wasm.dealloc).mock.calls.slice(callsBeforeClose),
    ).toContainEqual([pointer]);
    const callsAfterClose = vi.mocked(sqlite.wasm.dealloc).mock.calls.length;
    reader.close();
    expect(vi.mocked(sqlite.wasm.dealloc).mock.calls.length).toBe(
      callsAfterClose,
    );
  });

  it("reports a failed release, blocks queries, and completes on retry", async () => {
    const reader = await openReadOnlySqlite(await fixture());
    const sqlite = observed.runtimes.at(-1)!;
    const pointer = vi.mocked(sqlite.wasm.allocFromTypedArray).mock.results[0]!
      .value as number;
    vi.mocked(sqlite.wasm.dealloc).mockImplementationOnce(() => {
      throw new Error("release failed");
    });
    expect(() => reader.close()).toThrow(
      expect.objectContaining({
        code: "DB_SQLITE_READ_CLEANUP_REQUIRED",
        cause: undefined,
      }),
    );
    expect(() => reader.query("SELECT 1")).toThrow(
      expect.objectContaining({ code: "DB_SQLITE_READ_CLOSED" }),
    );
    const before = vi.mocked(sqlite.wasm.dealloc).mock.calls.length;
    reader.close();
    expect(vi.mocked(sqlite.wasm.dealloc).mock.calls.slice(before)).toEqual([
      [pointer],
    ]);
    expect(reader.wasmMemoryBytes).toBe(0);
  });

  it("keeps the catalog buffer while the database handle refuses to close", async () => {
    const reader = await openReadOnlySqlite(await fixture());
    const sqlite = observed.runtimes.at(-1)!;
    const pointer = vi.mocked(sqlite.wasm.allocFromTypedArray).mock.results[0]!
      .value as number;
    const closing = vi
      .spyOn(sqlite.oo1.DB.prototype, "close")
      .mockImplementationOnce(() => {
        throw new Error("handle busy");
      });
    try {
      const before = vi.mocked(sqlite.wasm.dealloc).mock.calls.length;
      expect(() => reader.close()).toThrow(
        expect.objectContaining({ code: "DB_SQLITE_READ_CLEANUP_REQUIRED" }),
      );
      expect(vi.mocked(sqlite.wasm.dealloc).mock.calls.length).toBe(before);
      expect(() => reader.query("SELECT 1")).toThrow(
        expect.objectContaining({ code: "DB_SQLITE_READ_CLOSED" }),
      );
      reader.close();
      expect(vi.mocked(sqlite.wasm.dealloc).mock.calls.slice(before)).toEqual([
        [pointer],
      ]);
      expect(reader.wasmMemoryBytes).toBe(0);
    } finally {
      closing.mockRestore();
    }
  });

  it("frees an allocated catalog after schema validation fails", async () => {
    const bytes = (await fixture()).slice(0, 100);
    await expect(openReadOnlySqlite(bytes)).rejects.toMatchObject({
      code: "DB_SQLITE_READ_INVALID_DATABASE",
    });
    const sqlite = observed.runtimes.at(-1)!;
    const pointer = vi.mocked(sqlite.wasm.allocFromTypedArray).mock.results[0]!
      .value as number;
    expect(sqlite.wasm.dealloc).toHaveBeenCalledWith(pointer);
  });
});
