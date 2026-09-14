import { beforeAll, describe, expect, test } from "vitest";
import sqlite3InitModule, { type Sqlite3Static } from "@sqlite.org/sqlite-wasm";

import { BrowserSqliteEngine } from "../src/engines/sqlite/browser.js";

let sqlite: Sqlite3Static;

beforeAll(async () => {
  sqlite = await sqlite3InitModule();
});

function createEngine(): BrowserSqliteEngine {
  return new BrowserSqliteEngine(sqlite, new sqlite.oo1.DB(":memory:"));
}

function createControlledEngine(options: {
  readonly failStatements: ReadonlyMap<string, readonly unknown[]>;
  readonly statements: string[];
}): BrowserSqliteEngine {
  const database = new sqlite.oo1.DB(":memory:");
  const failures = new Map(
    [...options.failStatements].map(([sql, values]) => [sql, [...values]]),
  );
  const exec = database.exec.bind(database);
  Object.defineProperty(database, "exec", {
    value(argument: unknown) {
      const sql =
        typeof argument === "string"
          ? argument
          : typeof argument === "object" &&
              argument !== null &&
              "sql" in argument &&
              typeof argument.sql === "string"
            ? argument.sql
            : undefined;
      if (sql !== undefined) {
        options.statements.push(sql);
        const statementFailures = failures.get(sql);
        if (statementFailures !== undefined && statementFailures.length > 0) {
          throw statementFailures.shift();
        }
      }
      if (typeof argument === "string") return exec(argument);
      if (sql === "SELECT 1") return [{ "1": 1 }];
      if (sql !== undefined) return [];
      throw new Error("Unexpected SQLite exec argument");
    },
  });
  return new BrowserSqliteEngine(sqlite, database);
}

function createStoredEngine(owner: object, name: string): BrowserSqliteEngine {
  return new BrowserSqliteEngine(sqlite, new sqlite.oo1.DB(":memory:"), false, {
    owner,
    name,
  });
}

describe("browser SQLite engine", () => {
  test("preserves large integers and blobs through prepared batches", async () => {
    const engine = createEngine();
    try {
      await engine.execute(
        "CREATE TABLE values_table (id INTEGER, payload BLOB, label TEXT)",
      );
      await engine.bulkInsert({
        table: "values_table",
        columns: ["id", "payload", "label"],
        rows: [
          [9_007_199_254_740_993n, new Uint8Array([0, 127, 255]), "first"],
          [-9_007_199_254_740_993n, new Uint8Array([4, 5]), "second"],
        ],
      });

      const rows = await engine.query(
        "SELECT id, payload, label FROM values_table ORDER BY label",
      );
      expect(rows).toEqual([
        {
          id: 9_007_199_254_740_993n,
          payload: new Uint8Array([0, 127, 255]),
          label: "first",
        },
        {
          id: -9_007_199_254_740_993n,
          payload: new Uint8Array([4, 5]),
          label: "second",
        },
      ]);
    } finally {
      await engine.close();
    }
  });

  test("rolls back a failed transaction", async () => {
    const engine = createEngine();
    const failure = new Error("stop");
    try {
      await engine.execute("CREATE TABLE events (value TEXT NOT NULL)");
      await expect(
        engine.transaction(async (transaction) => {
          await transaction.execute("INSERT INTO events (value) VALUES (?)", [
            "temporary",
          ]);
          throw failure;
        }),
      ).rejects.toBe(failure);
      expect(
        await engine.query("SELECT count(*) AS count FROM events"),
      ).toEqual([{ count: 0 }]);
    } finally {
      await engine.close();
    }
  });

  test.each([
    {
      label: "transaction work",
      begin: "BEGIN IMMEDIATE",
      execute: async (engine: BrowserSqliteEngine, failure: Error) =>
        engine.transaction(async () => {
          throw failure;
        }),
    },
    {
      label: "transaction commit",
      begin: "BEGIN IMMEDIATE",
      execute: async (engine: BrowserSqliteEngine) =>
        engine.transaction(async () => 1),
    },
    {
      label: "read transaction work",
      begin: "BEGIN",
      execute: async (engine: BrowserSqliteEngine, failure: Error) =>
        engine.readTransaction(async () => {
          throw failure;
        }),
    },
    {
      label: "snapshot commit",
      begin: "BEGIN",
      execute: async (engine: BrowserSqliteEngine) =>
        engine.copySnapshot(async () => 1),
    },
  ])(
    "quarantines SQLite after failed rollback following $label failure",
    async ({ label, begin, execute }) => {
      const primaryFailure = new Error(`Injected ${label} failure`);
      const rollbackFailure = new Error("Injected rollback failure");
      const statements: string[] = [];
      const failStatements = new Map<string, readonly unknown[]>([
        ["ROLLBACK", [rollbackFailure]],
      ]);
      if (label.endsWith("commit")) {
        failStatements.set("COMMIT", [primaryFailure]);
      }
      const engine = createControlledEngine({ failStatements, statements });

      let failure: unknown;
      try {
        await execute(engine, primaryFailure);
      } catch (error) {
        failure = error;
      }

      expect(statements).toContain(begin);
      expect(failure).toMatchObject({
        code: "DB_TRANSACTION_ROLLBACK_FAILED",
        details: {
          format: "sqlite",
          rollbackFailed: true,
          transactionState: "unknown",
        },
      });
      expect(((failure as Error).cause as AggregateError).errors).toEqual([
        primaryFailure,
        rollbackFailure,
      ]);
      await expect(engine.query("SELECT 2")).rejects.toBe(failure);
      await expect(engine.close()).resolves.toBeUndefined();
    },
  );

  test("does not roll back or quarantine SQLite when beginning fails", async () => {
    const beginFailure = new Error("Injected begin failure");
    const statements: string[] = [];
    const engine = createControlledEngine({
      failStatements: new Map([["BEGIN IMMEDIATE", [beginFailure]]]),
      statements,
    });

    await expect(engine.transaction(async () => 1)).rejects.toBe(beginFailure);
    await expect(engine.query("SELECT 1")).resolves.toEqual([{ "1": 1 }]);
    await engine.close();

    expect(statements).not.toContain("ROLLBACK");
  });

  test("yields to cancellation and rolls back a bounded batch", async () => {
    const engine = createEngine();
    const controller = new AbortController();
    try {
      await engine.execute("CREATE TABLE observations (value INTEGER)");
      const inserting = engine.transaction((transaction) =>
        transaction.bulkInsert({
          table: "observations",
          columns: ["value"],
          rows: Array.from({ length: 2_000 }, (_, value) => [value]),
          signal: controller.signal,
        }),
      );
      globalThis.setTimeout(() => controller.abort(), 0);

      await expect(inserting).rejects.toMatchObject({
        code: "OPERATION_ABORTED",
      });
      expect(
        await engine.query("SELECT count(*) AS count FROM observations"),
      ).toEqual([{ count: 0 }]);
    } finally {
      await engine.close();
    }
  });

  test("serializes concurrent work and refuses work after close", async () => {
    const engine = createEngine();
    await engine.execute("CREATE TABLE ordered_values (value INTEGER)");
    const inserting = engine.bulkInsert({
      table: "ordered_values",
      columns: ["value"],
      rows: Array.from({ length: 1_000 }, (_, value) => [value]),
    });
    const counting = engine.query(
      "SELECT count(*) AS count FROM ordered_values",
    );
    await inserting;
    await expect(counting).resolves.toEqual([{ count: 1_000 }]);

    await engine.checkpoint();
    await engine.close();
    await expect(engine.execute("SELECT 1")).rejects.toThrow(
      "SQLite engine is closed",
    );
  });

  test("holds queued mutations until a snapshot copy finishes", async () => {
    const engine = createEngine();
    let releaseCopy: () => void = () => undefined;
    const copyGate = new Promise<void>((resolve) => {
      releaseCopy = resolve;
    });
    let snapshotStarted: () => void = () => undefined;
    const started = new Promise<void>((resolve) => {
      snapshotStarted = resolve;
    });
    try {
      await engine.execute("CREATE TABLE events (value TEXT NOT NULL)");
      const snapshot = engine.copySnapshot(async () => {
        snapshotStarted();
        await copyGate;
        return "copied";
      });
      await started;

      let mutationFinished = false;
      const mutation = engine
        .execute("INSERT INTO events (value) VALUES ('after snapshot')")
        .then(() => {
          mutationFinished = true;
        });
      await new Promise<void>((resolve) => globalThis.setTimeout(resolve, 0));
      expect(mutationFinished).toBe(false);

      releaseCopy();
      await expect(snapshot).resolves.toBe("copied");
      await mutation;
      expect(
        await engine.query("SELECT value FROM events ORDER BY rowid"),
      ).toEqual([{ value: "after snapshot" }]);
    } finally {
      releaseCopy();
      await engine.close();
    }
  });

  test("releases queued work when a snapshot copy fails", async () => {
    const owner = {};
    const first = createStoredEngine(owner, "/failed.sqlite");
    const second = createStoredEngine(owner, "/failed.sqlite");
    let rejectCopy: (reason: Error) => void = () => undefined;
    const copyGate = new Promise<never>((_resolve, reject) => {
      rejectCopy = reject;
    });
    let snapshotStarted: () => void = () => undefined;
    const started = new Promise<void>((resolve) => {
      snapshotStarted = resolve;
    });
    let snapshot: Promise<never> | undefined;
    try {
      snapshot = first.copySnapshot(async () => {
        snapshotStarted();
        return copyGate;
      });
      await started;

      let secondHandleFinished = false;
      const secondHandleWork = second.execute("SELECT 1").then(() => {
        secondHandleFinished = true;
      });
      await new Promise<void>((resolve) => globalThis.setTimeout(resolve, 0));
      expect(secondHandleFinished).toBe(false);

      const failure = new Error("copy failed");
      rejectCopy(failure);
      await expect(snapshot).rejects.toBe(failure);
      await secondHandleWork;
      expect(secondHandleFinished).toBe(true);
    } finally {
      rejectCopy(new Error("test cleanup"));
      await snapshot?.catch(() => undefined);
      await Promise.all([first.close(), second.close()]);
    }
  });

  test("coordinates snapshot copies across handles for the same storage", async () => {
    const owner = {};
    const first = createStoredEngine(owner, "/shared.sqlite");
    const second = createStoredEngine(owner, "/shared.sqlite");
    let releaseCopy: () => void = () => undefined;
    const copyGate = new Promise<void>((resolve) => {
      releaseCopy = resolve;
    });
    let snapshotStarted: () => void = () => undefined;
    const started = new Promise<void>((resolve) => {
      snapshotStarted = resolve;
    });
    try {
      const snapshot = first.copySnapshot(async () => {
        snapshotStarted();
        await copyGate;
      });
      await started;

      let secondHandleFinished = false;
      const secondHandleWork = second.execute("SELECT 1").then(() => {
        secondHandleFinished = true;
      });
      await new Promise<void>((resolve) => globalThis.setTimeout(resolve, 0));
      expect(secondHandleFinished).toBe(false);

      releaseCopy();
      await snapshot;
      await secondHandleWork;
      expect(secondHandleFinished).toBe(true);
    } finally {
      releaseCopy();
      await Promise.all([first.close(), second.close()]);
    }
  });
});
