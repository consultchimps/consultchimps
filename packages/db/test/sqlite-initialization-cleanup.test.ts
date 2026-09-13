import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import BetterSqlite3 from "better-sqlite3";
import { expect, test, vi } from "vitest";

import { NodeSqliteEngine } from "../src/engines/sqlite/node.js";

test.each([new Error("Synthetic close failure"), undefined])(
  "preserves SQLite initialization and cleanup failures (%s)",
  async (cleanupFailure) => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-sqlite-init-"));
    const filename = path.join(directory, "database.sqlite");
    const initializationFailure = new Error("Synthetic pragma failure");
    const opened: BetterSqlite3.Database[] = [];
    const pragma = vi
      .spyOn(BetterSqlite3.prototype, "pragma")
      .mockImplementation(function (this: BetterSqlite3.Database) {
        opened.push(this);
        throw initializationFailure;
      });
    const close = vi
      .spyOn(BetterSqlite3.prototype, "close")
      .mockImplementation(() => {
        throw cleanupFailure;
      });
    try {
      expect(() => NodeSqliteEngine.create(filename)).toThrow(
        expect.objectContaining({
          code: "DB_NATIVE_SQLITE_CLEANUP_REQUIRED",
          cause: expect.objectContaining({
            errors: [initializationFailure, cleanupFailure],
          }),
        }),
      );
      expect(close).toHaveBeenCalledTimes(1);
    } finally {
      pragma.mockRestore();
      close.mockRestore();
      for (const database of opened) database.close();
      await rm(directory, { recursive: true, force: true });
    }
  },
);
