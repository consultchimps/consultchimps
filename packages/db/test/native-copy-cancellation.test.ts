import { afterEach, expect, test, vi } from "vitest";

const fakes = vi.hoisted(() => ({
  sqliteBackup: vi.fn(),
  duckRun: vi.fn(),
  duckRead: vi.fn(),
  duckInterrupt: vi.fn(),
  duckConnectionClose: vi.fn(),
  duckInstanceClose: vi.fn(),
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
      create: vi.fn(async () => ({
        connect: async () => connection,
        closeSync: fakes.duckInstanceClose,
      })),
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
