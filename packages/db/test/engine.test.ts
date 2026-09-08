import { describe, expect, it } from "vitest";

import { loadSqlDatabase } from "../src/index.js";

describe("SqlDatabase", () => {
  it("runs statements, reads rows, and serializes back to bytes", async () => {
    const database = await loadSqlDatabase();
    database.run("CREATE TABLE t (id INTEGER, label TEXT);");
    database.run("INSERT INTO t (id, label) VALUES (?, ?);", [1, "one"]);
    database.run("INSERT INTO t (id, label) VALUES (?, ?);", [2, "two"]);

    const rows = database.select("SELECT id, label FROM t ORDER BY id;");
    expect(rows).toEqual([
      { id: 1, label: "one" },
      { id: 2, label: "two" },
    ]);

    expect(database.selectValue("SELECT count(*) FROM t;")).toBe(2);
    expect(
      database.selectValue("SELECT label FROM t WHERE id = 9;"),
    ).toBeNull();

    const bytes = database.serialize();
    expect(bytes.byteLength).toBeGreaterThan(0);
    database.close();

    const reloaded = await loadSqlDatabase(bytes);
    expect(reloaded.select("SELECT label FROM t ORDER BY id;")).toEqual([
      { label: "one" },
      { label: "two" },
    ]);
    reloaded.close();
  });

  it("refuses an asynchronous transaction callback and rolls back", async () => {
    const database = await loadSqlDatabase();
    database.run("CREATE TABLE t (a INTEGER);");
    expect(() =>
      database.transaction(async () => {
        database.run("INSERT INTO t (a) VALUES (1);");
      }),
    ).toThrow(expect.objectContaining({ code: "DB_ASYNC_TRANSACTION" }));
    // The statement the callback ran before its first await was rolled back.
    expect(database.selectValue("SELECT count(*) FROM t;")).toBe(0);
    database.close();
  });

  it("supports named parameters", async () => {
    const database = await loadSqlDatabase();
    database.run("CREATE TABLE t (a INTEGER);");
    database.run("INSERT INTO t (a) VALUES ($a);", { $a: 5 });
    expect(database.selectValue("SELECT a FROM t;")).toBe(5);
    database.close();
  });
});
