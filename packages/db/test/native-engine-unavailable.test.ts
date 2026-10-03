import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { expect, test, vi } from "vitest";

const loadFailure = vi.hoisted(
  () => new Error("Synthetic native load failure"),
);

vi.mock("@duckdb/node-api", () => {
  throw loadFailure;
});

import { NodeDuckDbEngine } from "../src/engines/duckdb/node.js";
import { inspectFileKind, openDatabase } from "../src/node.js";

test("an engine that cannot load is named, not blamed on the file", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-native-missing-"));
  try {
    // Anything without a SQLite header is handed to DuckDB.
    const filePath = path.join(directory, "database.duckdb");
    await writeFile(filePath, "not a SQLite file");
    const unavailable = expect.objectContaining({
      code: "DB_NATIVE_ENGINE_UNAVAILABLE",
      details: { format: "duckdb", module: "@duckdb/node-api" },
    });
    await expect(NodeDuckDbEngine.create(filePath)).rejects.toThrow(
      unavailable,
    );
    await expect(openDatabase({ path: filePath })).rejects.toThrow(unavailable);
    await expect(inspectFileKind({ path: filePath })).rejects.toThrow(
      unavailable,
    );
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
});
