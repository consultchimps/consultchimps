import { afterEach, expect, test, vi } from "vitest";

const fakes = vi.hoisted(() => ({
  sqliteBackup: vi.fn(),
  duckRun: vi.fn(),
  duckRead: vi.fn(),
  duckInterrupt: vi.fn(),
  duckConnectionClose: vi.fn(),
  duckInstanceClose: vi.fn(),
  duckConnect: vi.fn(async (connection: unknown) => connection),
  duckCreate: vi.fn(),
  onDuckCopy: (): void => {},
  rejectDuckCopy: (error: unknown): void => {
    void error;
  },
}));

vi.mock("better-sqlite3", () => ({
  default: class FakeSqliteDatabase {
    pragma(): void {}
    defaultSafeIntegers(): void {}
    backup(
      destination: string,
      options: {
        readonly progress: (info: {
          readonly totalPages: number;
          readonly remainingPages: number;
        }) => number;
      },
    ): Promise<unknown> {
      return fakes.sqliteBackup(destination, options);
    }
    close(): void {}
  },
}));

vi.mock("@duckdb/node-api", () => {
  const connection = {
    run: fakes.duckRun,
    runAndReadAll: fakes.duckRead,
    interrupt: fakes.duckInterrupt,
    closeSync: fakes.duckConnectionClose,
  };
  return {
    blobValue: (value: unknown) => value,
    DuckDBBlobValue: class FakeDuckDbBlobValue {},
    DuckDBInstance: {
      create: async (...args: unknown[]) => {
        fakes.duckCreate(...args);
        return {
          connect: () => fakes.duckConnect(connection),
          closeSync: fakes.duckInstanceClose,
        };
      },
    },
  };
});

import { NodeDuckDbEngine } from "../src/engines/duckdb/node.js";
import { NodeSqliteEngine } from "../src/engines/sqlite/node.js";

afterEach(() => {
  vi.clearAllMocks();
  fakes.onDuckCopy = (): void => {};
  fakes.rejectDuckCopy = (): void => {};
});

test("DuckDB closes its instance when connection creation fails", async () => {
  const connectionFailure = new Error("connect failed");
  fakes.duckConnect.mockRejectedValueOnce(connectionFailure);

  await expect(NodeDuckDbEngine.create("source.duckdb")).rejects.toBe(
    connectionFailure,
  );
  expect(fakes.duckInstanceClose).toHaveBeenCalledOnce();
});

test("DuckDB preserves connection and cleanup failures during creation", async () => {
  const connectionFailure = new Error("connect failed");
  const cleanupFailure = new Error("close failed");
  fakes.duckConnect.mockRejectedValueOnce(connectionFailure);
  fakes.duckInstanceClose.mockImplementationOnce(() => {
    throw cleanupFailure;
  });

  const failure = await NodeDuckDbEngine.create("source.duckdb").catch(
    (error: unknown) => error,
  );
  expect(failure).toMatchObject({
    code: "DB_DUCKDB_OPEN_CLEANUP_FAILED",
    details: { path: "source.duckdb", closeFailed: true },
  });
  expect((failure as Error).cause).toBeInstanceOf(AggregateError);
  expect(((failure as Error).cause as AggregateError).errors).toEqual([
    connectionFailure,
    cleanupFailure,
  ]);
});

test("SQLite backup stops from its progress boundary", async () => {
  const controller = new AbortController();
  fakes.sqliteBackup.mockImplementation(
    async (
      _destination: string,
      options: {
        readonly progress: (info: {
          readonly totalPages: number;
          readonly remainingPages: number;
        }) => number;
      },
    ) => {
      controller.abort("cancelled during backup");
      options.progress({ totalPages: 2, remainingPages: 1 });
    },
  );
  const engine = NodeSqliteEngine.create("source.sqlite");
  try {
    await expect(
      engine.backupTo("private.sqlite", controller.signal),
    ).rejects.toMatchObject({ code: "OPERATION_ABORTED" });
  } finally {
    await engine.close();
  }
});

test("DuckDB cancellation interrupts the active database copy and detaches its target", async () => {
  const controller = new AbortController();
  fakes.duckRead.mockResolvedValue({
    getRowObjects: () => [{ database_name: "source" }],
  });
  fakes.duckRun.mockImplementation(async (sql: string) => {
    if (!sql.startsWith("COPY FROM DATABASE")) return;
    await new Promise<void>((_resolve, reject) => {
      fakes.rejectDuckCopy = reject;
      queueMicrotask(fakes.onDuckCopy);
    });
  });
  fakes.duckInterrupt.mockImplementation(() => {
    fakes.rejectDuckCopy(new Error("interrupted"));
  });
  fakes.onDuckCopy = () => controller.abort("cancelled during copy");
  const engine = await NodeDuckDbEngine.create("source.duckdb");
  try {
    await expect(
      engine.copyTo("private.duckdb", controller.signal),
    ).rejects.toMatchObject({ code: "OPERATION_ABORTED" });
    expect(fakes.duckInterrupt).toHaveBeenCalledOnce();
    expect(
      fakes.duckRun.mock.calls.some(([sql]) =>
        String(sql).startsWith("DETACH"),
      ),
    ).toBe(true);
  } finally {
    await engine.close();
  }
});

test("DuckDB readonly cancellation releases its private export resources", async () => {
  const controller = new AbortController();
  fakes.duckRun.mockImplementation(async (sql: string) => {
    if (!sql.startsWith("COPY FROM DATABASE")) return;
    await new Promise<void>((_resolve, reject) => {
      fakes.rejectDuckCopy = reject;
      queueMicrotask(fakes.onDuckCopy);
    });
  });
  fakes.duckInterrupt.mockImplementation(() => {
    fakes.rejectDuckCopy(new Error("interrupted"));
  });
  fakes.onDuckCopy = () => controller.abort("cancelled during copy");
  const engine = await NodeDuckDbEngine.create("source.duckdb", true);
  try {
    await expect(
      engine.copyTo("private.duckdb", controller.signal),
    ).rejects.toMatchObject({ code: "OPERATION_ABORTED" });
    expect(fakes.duckInterrupt).toHaveBeenCalledOnce();
    expect(fakes.duckConnectionClose).toHaveBeenCalledOnce();
    expect(fakes.duckInstanceClose).toHaveBeenCalledOnce();
  } finally {
    await engine.close();
  }
});

test("DuckDB readonly copy owns its temp directory and attempts every resource close", async () => {
  const engine = await NodeDuckDbEngine.create("source.duckdb", true);
  const copyFailure = new Error("copy failed");
  const connectionCleanupFailure = new Error("connection close failed");
  const instanceCleanupFailure = new Error("instance close failed");
  fakes.duckRun.mockImplementation(async (sql: string) => {
    if (sql.startsWith("COPY FROM DATABASE")) throw copyFailure;
  });
  fakes.duckConnectionClose.mockImplementationOnce(() => {
    throw connectionCleanupFailure;
  });
  fakes.duckInstanceClose.mockImplementationOnce(() => {
    throw instanceCleanupFailure;
  });
  try {
    const failure = await engine
      .copyTo("private.duckdb")
      .catch((error: unknown) => error);
    expect(failure).toMatchObject({
      code: "DB_DUCKDB_EXPORT_CLEANUP_FAILED",
      details: { operation: "readonly-copy" },
    });
    expect(fakes.duckCreate).toHaveBeenNthCalledWith(2, ":memory:", {
      temp_directory: "private.duckdb.tmp",
    });
    expect(fakes.duckInstanceClose).toHaveBeenCalledOnce();
    expect((failure as Error).cause).toBeInstanceOf(AggregateError);
    expect(((failure as Error).cause as AggregateError).errors).toEqual([
      copyFailure,
      connectionCleanupFailure,
      instanceCleanupFailure,
    ]);
  } finally {
    fakes.duckRun.mockResolvedValue(undefined);
    await engine.close();
  }
});

test("DuckDB reports copy and destination detach failures without replacing either", async () => {
  const copyFailure = new Error("copy failed");
  const detachFailure = new Error("detach failed");
  fakes.duckRead.mockResolvedValue({
    getRowObjects: () => [{ database_name: "source" }],
  });
  fakes.duckRun.mockImplementation(async (sql: string) => {
    if (sql.startsWith("COPY FROM DATABASE")) throw copyFailure;
    if (sql.startsWith("DETACH")) throw detachFailure;
  });
  const engine = await NodeDuckDbEngine.create("source.duckdb");
  try {
    const failure = await engine
      .copyTo("private.duckdb")
      .catch((error: unknown) => error);
    expect(failure).toMatchObject({
      code: "DB_DUCKDB_EXPORT_CLEANUP_FAILED",
      details: {
        destination: "private.duckdb",
        attachment: expect.stringMatching(/^cc_export_/),
      },
    });
    expect((failure as Error).cause).toBeInstanceOf(AggregateError);
    expect(((failure as Error).cause as AggregateError).errors).toEqual([
      copyFailure,
      detachFailure,
    ]);
    expect(
      fakes.duckRun.mock.calls.filter(([sql]) =>
        String(sql).startsWith("DETACH"),
      ),
    ).toHaveLength(1);
  } finally {
    fakes.duckRun.mockResolvedValue(undefined);
    await engine.close();
  }
});

test("DuckDB preserves failed destination cleanup when cancellation interrupts copy", async () => {
  const controller = new AbortController();
  const copyFailure = new Error("interrupted");
  const detachFailure = new Error("detach failed");
  fakes.duckRead.mockResolvedValue({
    getRowObjects: () => [{ database_name: "source" }],
  });
  fakes.duckRun.mockImplementation(async (sql: string) => {
    if (sql.startsWith("COPY FROM DATABASE")) {
      controller.abort("cancelled during copy");
      throw copyFailure;
    }
    if (sql.startsWith("DETACH")) throw detachFailure;
  });
  const engine = await NodeDuckDbEngine.create("source.duckdb");
  try {
    const failure = await engine
      .copyTo("private.duckdb", controller.signal)
      .catch((error: unknown) => error);
    expect(failure).toMatchObject({
      code: "DB_DUCKDB_EXPORT_CLEANUP_FAILED",
    });
    expect(((failure as Error).cause as AggregateError).errors).toEqual([
      copyFailure,
      detachFailure,
    ]);
  } finally {
    fakes.duckRun.mockResolvedValue(undefined);
    await engine.close();
  }
});
