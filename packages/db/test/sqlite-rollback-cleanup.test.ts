import { beforeEach, expect, test, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  exec: vi.fn<(sql: string) => void>(),
  close: vi.fn<() => void>(),
  inTransaction: true,
}));

vi.mock("better-sqlite3", () => ({
  default: class {
    pragma() {}
    defaultSafeIntegers() {}
    get inTransaction() {
      return mocks.inTransaction;
    }
    exec(sql: string) {
      mocks.exec(sql);
    }
    prepare() {
      return { all: () => [], run: () => undefined };
    }
    close() {
      mocks.close();
    }
  },
}));

import { NodeSqliteEngine } from "../src/engines/sqlite/node.js";

beforeEach(() => {
  vi.resetAllMocks();
  mocks.inTransaction = true;
});

test.each([
  ["transaction", "work"],
  ["transaction", "commit"],
  ["readTransaction", "work"],
  ["readTransaction", "commit"],
] as const)(
  "%s preserves %s and rollback failures and permits only cleanup",
  async (method, phase) => {
    const original = new Error("Injected operation failure");
    const rollback = new Error("Injected rollback failure");
    mocks.exec.mockImplementation((sql) => {
      if (sql === "COMMIT" && phase === "commit") throw original;
      if (sql === "ROLLBACK") throw rollback;
    });
    const engine = NodeSqliteEngine.create("synthetic.sqlite");
    const failure = await engine[method](async () => {
      if (phase === "work") throw original;
      return 1;
    }).catch((error: unknown) => error);
    expect(failure).toMatchObject({
      code: "DB_TRANSACTION_ROLLBACK_FAILED",
      details: {
        format: "sqlite",
        rollbackFailed: true,
        transactionState: "unknown",
      },
      cause: expect.any(AggregateError),
    });
    expect((failure as Error).cause).toMatchObject({
      errors: [original, rollback],
    });
    const calls = mocks.exec.mock.calls.length;
    const nextWork = vi.fn(async () => undefined);
    await expect(engine.transaction(nextWork)).rejects.toBe(failure);
    await expect(engine.query("SELECT 1")).rejects.toBe(failure);
    await expect(engine.checkpoint()).rejects.toBe(failure);
    expect(nextWork).not.toHaveBeenCalled();
    expect(mocks.exec.mock.calls).toHaveLength(calls);
    const closeFailure = new Error("Injected close failure");
    mocks.close.mockImplementationOnce(() => {
      throw closeFailure;
    });
    await expect(engine.close()).rejects.toBe(closeFailure);
    await engine.close();
    expect(mocks.close).toHaveBeenCalledTimes(2);
  },
);

test.each([true, false])(
  "preserves the primary failure when SQLite inTransaction is %s and cleanup succeeds",
  async (active) => {
    const original = new Error("Injected work failure");
    const engine = NodeSqliteEngine.create("synthetic.sqlite");
    await expect(
      engine.transaction(async () => {
        mocks.inTransaction = active;
        throw original;
      }),
    ).rejects.toBe(original);
    expect(
      mocks.exec.mock.calls.filter(([sql]) => sql === "ROLLBACK"),
    ).toHaveLength(active ? 1 : 0);
    await expect(engine.query("SELECT 1")).resolves.toEqual([]);
    await engine.close();
  },
);

test("does not roll back or quarantine after BEGIN fails", async () => {
  const original = new Error("Injected begin failure");
  mocks.exec.mockImplementationOnce(() => {
    throw original;
  });
  const engine = NodeSqliteEngine.create("synthetic.sqlite");
  const work = vi.fn(async () => undefined);
  await expect(engine.transaction(work)).rejects.toBe(original);
  expect(work).not.toHaveBeenCalled();
  expect(mocks.exec.mock.calls).toEqual([["BEGIN IMMEDIATE"]]);
  await expect(engine.query("SELECT 1")).resolves.toEqual([]);
  await engine.close();
});
