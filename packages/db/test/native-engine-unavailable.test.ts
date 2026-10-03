import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterAll, beforeAll, expect, test, vi } from "vitest";

const failures = vi.hoisted(() => ({
  duckdb: new Error("Synthetic DuckDB module load failure"),
  // better-sqlite3 loads its addon in the first constructor call, which is
  // where a missing or mismatched binary throws.
  sqlite: Object.assign(new Error("Synthetic SQLite addon load failure"), {
    code: "ERR_DLOPEN_FAILED",
  }),
}));

vi.mock("@duckdb/node-api", () => {
  throw failures.duckdb;
});

vi.mock("../src/engines/sqlite/binding.js", () => ({
  betterSqlite3: () =>
    class {
      constructor() {
        throw failures.sqlite;
      }
    },
}));

import { NodeDuckDbEngine } from "../src/engines/duckdb/node.js";
import { NodeSqliteEngine } from "../src/engines/sqlite/node.js";
import { inspectFileKind, openDatabase, openImportBatch } from "../src/node.js";

let directory: string;

beforeAll(async () => {
  directory = await mkdtemp(path.join(tmpdir(), "cc-native-missing-"));
});

afterAll(async () => {
  await rm(directory, { force: true, recursive: true });
});

function unavailable(format: string, module: string): unknown {
  return expect.objectContaining({
    code: "DB_NATIVE_ENGINE_UNAVAILABLE",
    details: { format, module },
  });
}

test("a DuckDB engine that cannot load is named, not blamed on the file", async () => {
  // Anything without a SQLite header is handed to DuckDB.
  const filePath = path.join(directory, "database.duckdb");
  await writeFile(filePath, "not a SQLite file");
  const expected = unavailable("duckdb", "@duckdb/node-api");
  await expect(NodeDuckDbEngine.create(filePath)).rejects.toThrow(expected);
  await expect(openDatabase({ path: filePath })).rejects.toThrow(expected);
  await expect(inspectFileKind({ path: filePath })).rejects.toThrow(expected);
});

test("a SQLite addon that cannot load is named, not blamed on the file", async () => {
  const filePath = path.join(directory, "database.sqlite");
  const header = Buffer.alloc(100);
  header.write("SQLite format 3\0", "latin1");
  await writeFile(filePath, header);
  const expected = unavailable("sqlite", "better-sqlite3");
  expect(() => NodeSqliteEngine.create(filePath)).toThrow(expected);
  await expect(openDatabase({ path: filePath })).rejects.toThrow(expected);
  await expect(inspectFileKind({ path: filePath })).rejects.toThrow(expected);
  await expect(openImportBatch({ path: filePath })).rejects.toThrow(expected);
});
