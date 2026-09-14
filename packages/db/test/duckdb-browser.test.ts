import type { RandomAccessFile } from "@consultchimps/core";
import { afterEach, beforeEach, expect, test, vi } from "vitest";

const duckdb = vi.hoisted(() => ({
  accessMode: "",
  openPath: "",
  registrations: [] as string[],
  checkpointFailure: undefined as Error | undefined,
  queryFailures: new Map<string, unknown[]>(),
  connectFailure: undefined as Error | undefined,
  terminationFailures: [] as Error[],
  terminationStarted: undefined as (() => void) | undefined,
  terminationWait: undefined as Promise<void> | undefined,
  queries: [] as string[],
  connectionCloses: 0,
  flushes: 0,
  terminations: 0,
}));

vi.mock("@duckdb/duckdb-wasm/dist/duckdb-browser", () => ({
  DuckDBAccessMode: {
    READ_ONLY: "read-only",
    READ_WRITE: "read-write",
  },
  DuckDBDataProtocol: { BROWSER_FSACCESS: "browser-file" },
  VoidLogger: class {},
  AsyncDuckDB: class {
    async instantiate(): Promise<void> {}

    async registerFileHandle(name: string): Promise<void> {
      duckdb.registrations.push(name);
    }

    async dropFile(): Promise<void> {}

    async open(options: {
      readonly accessMode: string;
      readonly path: string;
    }): Promise<void> {
      duckdb.accessMode = options.accessMode;
      duckdb.openPath = options.path;
    }

    async connect() {
      if (duckdb.connectFailure !== undefined) throw duckdb.connectFailure;
      return {
        async query(sql: string) {
          duckdb.queries.push(sql);
          const failures = duckdb.queryFailures.get(sql);
          if (failures !== undefined && failures.length > 0) {
            throw failures.shift();
          }
          if (sql === "CHECKPOINT" && duckdb.checkpointFailure !== undefined) {
            throw duckdb.checkpointFailure;
          }
          return {
            toArray: () =>
              sql.includes("current_database()")
                ? [{ toJSON: () => ({ database_name: "source" }) }]
                : [],
          };
        },
        async prepare(): Promise<never> {
          throw new Error("Unexpected prepared statement");
        },
        async cancelSent(): Promise<void> {},
        async close(): Promise<void> {
          duckdb.connectionCloses += 1;
        },
      };
    }

    async flushFiles(): Promise<void> {
      duckdb.flushes += 1;
    }

    async terminate(): Promise<void> {
      duckdb.terminations += 1;
      duckdb.terminationStarted?.();
      await duckdb.terminationWait;
      const failure = duckdb.terminationFailures.shift();
      if (failure !== undefined) throw failure;
    }
  },
}));

import {
  BrowserDuckDbEngine,
  BrowserDuckDbOpenCleanupError,
} from "../src/engines/duckdb/browser.js";

function memoryFile(): RandomAccessFile & { bytes(): Uint8Array } {
  let stored = new Uint8Array();
  return {
    name: "copy.duckdb",
    get size() {
      return stored.byteLength;
    },
    async readAt(offset, length) {
      return stored.slice(offset, offset + length);
    },
    async writeAt(offset, bytes) {
      const required = offset + bytes.byteLength;
      if (required > stored.byteLength) {
        const expanded = new Uint8Array(required);
        expanded.set(stored);
        stored = expanded;
      }
      stored.set(bytes, offset);
    },
    async truncate(size) {
      stored = stored.slice(0, size);
    },
    async close() {},
    bytes() {
      return stored.slice();
    },
  };
}

beforeEach(() => {
  duckdb.accessMode = "";
  duckdb.openPath = "";
  duckdb.registrations.length = 0;
  duckdb.checkpointFailure = undefined;
  duckdb.queryFailures.clear();
  duckdb.connectFailure = undefined;
  duckdb.terminationFailures.length = 0;
  duckdb.terminationStarted = undefined;
  duckdb.terminationWait = undefined;
  duckdb.queries.length = 0;
  duckdb.connectionCloses = 0;
  duckdb.flushes = 0;
  duckdb.terminations = 0;
  vi.stubGlobal("Worker", class {});
});

afterEach(() => {
  vi.unstubAllGlobals();
});

async function createEngine(
  storageName = "transaction.duckdb",
): Promise<BrowserDuckDbEngine> {
  return BrowserDuckDbEngine.open({
    wasmUrl: "duckdb.wasm",
    workerUrl: "duckdb.worker.js",
    storageName,
    fileHandle: {
      async getFile() {
        return new Blob();
      },
    },
    walHandle: {
      async getFile() {
        return new Blob();
      },
    },
  });
}

test("preserves the original DuckDB transaction failure after rollback", async () => {
  const engine = await createEngine();
  const workFailure = new Error("Injected transaction failure");

  await expect(
    engine.transaction(async () => {
      throw workFailure;
    }),
  ).rejects.toBe(workFailure);
  await expect(engine.query("SELECT 1")).resolves.toEqual([]);
  await engine.close();

  expect(duckdb.queries).toContain("ROLLBACK");
});

test.each([
  {
    label: "transaction work",
    execute: async (engine: BrowserDuckDbEngine, failure: Error) =>
      engine.transaction(async () => {
        throw failure;
      }),
  },
  {
    label: "transaction commit",
    execute: async (engine: BrowserDuckDbEngine) =>
      engine.transaction(async () => 1),
  },
  {
    label: "read transaction work",
    execute: async (engine: BrowserDuckDbEngine, failure: Error) =>
      engine.readTransaction(async () => {
        throw failure;
      }),
  },
])(
  "quarantines DuckDB after failed rollback following $label failure",
  async ({ label, execute }) => {
    const engine = await createEngine(`${label.replaceAll(" ", "-")}.duckdb`);
    const primaryFailure = new Error(`Injected ${label} failure`);
    const rollbackFailure = new Error("Injected rollback failure");
    if (label === "transaction commit") {
      duckdb.queryFailures.set("COMMIT", [primaryFailure]);
    }
    duckdb.queryFailures.set("ROLLBACK", [rollbackFailure]);

    let failure: unknown;
    try {
      await execute(engine, primaryFailure);
    } catch (error) {
      failure = error;
    }

    expect(failure).toMatchObject({
      code: "DB_TRANSACTION_ROLLBACK_FAILED",
      details: {
        format: "duckdb",
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

test("does not roll back or quarantine DuckDB when beginning fails", async () => {
  const engine = await createEngine("begin-failure.duckdb");
  const beginFailure = new Error("Injected begin failure");
  duckdb.queryFailures.set("BEGIN TRANSACTION", [beginFailure]);

  await expect(engine.transaction(async () => 1)).rejects.toBe(beginFailure);
  await expect(engine.query("SELECT 1")).resolves.toEqual([]);
  await engine.close();

  expect(duckdb.queries).not.toContain("ROLLBACK");
});

test("retains retryable ownership when a failed open cannot terminate", async () => {
  const openFailure = new Error("Injected DuckDB connect failure");
  const cleanupFailure = new Error("Injected DuckDB termination failure");
  duckdb.connectFailure = openFailure;
  duckdb.terminationFailures.push(cleanupFailure);

  let failure: unknown;
  try {
    await BrowserDuckDbEngine.open({
      wasmUrl: "duckdb.wasm",
      workerUrl: "duckdb.worker.js",
      storageName: "failed-open.duckdb",
      fileHandle: {
        async getFile() {
          return new Blob();
        },
      },
      walHandle: {
        async getFile() {
          return new Blob();
        },
      },
    });
  } catch (error) {
    failure = error;
  }

  expect(failure).toBeInstanceOf(BrowserDuckDbOpenCleanupError);
  expect((failure as BrowserDuckDbOpenCleanupError).openCause).toBe(
    openFailure,
  );
  expect((failure as BrowserDuckDbOpenCleanupError).cleanupCause).toBe(
    cleanupFailure,
  );
  expect(((failure as Error).cause as AggregateError).errors).toEqual([
    openFailure,
    cleanupFailure,
  ]);
  await (failure as BrowserDuckDbOpenCleanupError).retryCleanup();
  expect(duckdb.terminations).toBe(2);
});

test("copies and closes a read-only browser database without a write checkpoint", async () => {
  const sourceBytes = new Uint8Array(1024 * 1024 + 17);
  sourceBytes[0] = 7;
  sourceBytes[sourceBytes.length - 1] = 9;
  const engine = await BrowserDuckDbEngine.open({
    wasmUrl: "duckdb.wasm",
    workerUrl: "duckdb.worker.js",
    storageName: "readonly.duckdb",
    fileHandle: {
      async getFile() {
        return new Blob([sourceBytes]);
      },
    },
    walHandle: {
      async getFile() {
        return new Blob();
      },
    },
    readonly: true,
    async createReadonlySnapshot() {
      return {
        name: "snapshot.duckdb",
        path: "/snapshot.duckdb",
        directory: "test",
        fileHandle: {
          async getFile() {
            return new Blob();
          },
        },
        walHandle: {
          async getFile() {
            return new Blob();
          },
        },
        async copyTo(destination) {
          await destination.truncate(0);
          await destination.writeAt(0, sourceBytes);
          await destination.truncate(sourceBytes.byteLength);
          return sourceBytes.byteLength;
        },
        async remove() {},
      };
    },
  });
  const destination = memoryFile();

  await engine.checkpoint();
  await expect(engine.copyTo(destination)).resolves.toBe(
    sourceBytes.byteLength,
  );
  await engine.close();

  expect(duckdb.accessMode).toBe("read-only");
  expect(duckdb.openPath).toBe("readonly.duckdb");
  expect(duckdb.queries).not.toContain("CHECKPOINT");
  expect(destination.bytes()).toEqual(sourceBytes);
  expect(duckdb.connectionCloses).toBe(1);
  expect(duckdb.flushes).toBe(3);
  expect(duckdb.terminations).toBe(1);
});

test("holds read-only engine work until snapshot copying finishes", async () => {
  let releaseCopy = (): void => undefined;
  const copyReleased = new Promise<void>((resolve) => {
    releaseCopy = resolve;
  });
  let copyStarted = (): void => undefined;
  const copyObserved = new Promise<void>((resolve) => {
    copyStarted = resolve;
  });
  const engine = await BrowserDuckDbEngine.open({
    wasmUrl: "duckdb.wasm",
    workerUrl: "duckdb.worker.js",
    storageName: "readonly-concurrent.duckdb",
    fileHandle: {
      async getFile() {
        return new Blob();
      },
    },
    walHandle: {
      async getFile() {
        return new Blob();
      },
    },
    readonly: true,
    async createReadonlySnapshot() {
      return {
        name: "snapshot-concurrent.duckdb",
        path: "/snapshot-concurrent.duckdb",
        directory: "test",
        fileHandle: {
          async getFile() {
            return new Blob();
          },
        },
        walHandle: {
          async getFile() {
            return new Blob();
          },
        },
        async copyTo() {
          copyStarted();
          await copyReleased;
          return 0;
        },
        async remove() {},
      };
    },
  });

  const copying = engine.copyTo(memoryFile());
  await copyObserved;
  const querying = engine.query("SELECT 1");
  await expect(
    Promise.race([
      querying.then(() => "queried"),
      new Promise<string>((resolve) => {
        setTimeout(() => resolve("waiting"), 20);
      }),
    ]),
  ).resolves.toBe("waiting");
  releaseCopy();
  await copying;
  await querying;
  await engine.close();

  expect(duckdb.queries).toContain("SELECT 1");
});

test("releases browser DuckDB resources when a writable close checkpoint fails", async () => {
  const checkpointFailure = new Error("Injected checkpoint failure");
  duckdb.checkpointFailure = checkpointFailure;
  const engine = await BrowserDuckDbEngine.open({
    wasmUrl: "duckdb.wasm",
    workerUrl: "duckdb.worker.js",
    storageName: "writable.duckdb",
    fileHandle: {
      async getFile() {
        return new Blob();
      },
    },
    walHandle: {
      async getFile() {
        return new Blob();
      },
    },
  });

  await expect(engine.close()).rejects.toBe(checkpointFailure);
  expect(duckdb.queries).toContain("CHECKPOINT");
  expect(duckdb.connectionCloses).toBe(1);
  expect(duckdb.flushes).toBe(1);
  expect(duckdb.terminations).toBe(1);
  await expect(engine.close()).resolves.toBeUndefined();
  await expect(engine.query("SELECT 1")).rejects.toThrow(
    "DuckDB engine is closed",
  );
  expect(duckdb.connectionCloses).toBe(1);
  expect(duckdb.flushes).toBe(1);
  expect(duckdb.terminations).toBe(1);
});

test("retries only unfinished browser DuckDB cleanup after termination fails", async () => {
  const terminationFailure = new Error("Injected termination failure");
  duckdb.terminationFailures.push(terminationFailure);
  const engine = await BrowserDuckDbEngine.open({
    wasmUrl: "duckdb.wasm",
    workerUrl: "duckdb.worker.js",
    storageName: "termination-retry.duckdb",
    fileHandle: {
      async getFile() {
        return new Blob();
      },
    },
    walHandle: {
      async getFile() {
        return new Blob();
      },
    },
  });

  await expect(engine.close()).rejects.toBe(terminationFailure);
  await expect(engine.query("SELECT 1")).rejects.toThrow(
    "DuckDB engine is closed",
  );
  expect(duckdb.queries.filter((sql) => sql === "CHECKPOINT")).toHaveLength(1);
  expect(duckdb.connectionCloses).toBe(1);
  expect(duckdb.flushes).toBe(1);
  expect(duckdb.terminations).toBe(1);

  await expect(engine.close()).resolves.toBeUndefined();
  await expect(engine.close()).resolves.toBeUndefined();
  expect(duckdb.queries.filter((sql) => sql === "CHECKPOINT")).toHaveLength(1);
  expect(duckdb.connectionCloses).toBe(1);
  expect(duckdb.flushes).toBe(1);
  expect(duckdb.terminations).toBe(2);
});

test.each([
  { label: "success", checkpointFailure: undefined },
  {
    label: "checkpoint failure",
    checkpointFailure: new Error("Injected concurrent checkpoint failure"),
  },
])(
  "shares one in-flight browser DuckDB close attempt after $label",
  async ({ checkpointFailure }) => {
    duckdb.checkpointFailure = checkpointFailure;
    let releaseTermination = (): void => undefined;
    duckdb.terminationWait = new Promise<void>((resolve) => {
      releaseTermination = resolve;
    });
    let observeTermination = (): void => undefined;
    const terminationStarted = new Promise<void>((resolve) => {
      observeTermination = resolve;
    });
    duckdb.terminationStarted = observeTermination;
    const engine = await BrowserDuckDbEngine.open({
      wasmUrl: "duckdb.wasm",
      workerUrl: "duckdb.worker.js",
      storageName: "concurrent-close.duckdb",
      fileHandle: {
        async getFile() {
          return new Blob();
        },
      },
      walHandle: {
        async getFile() {
          return new Blob();
        },
      },
    });

    const first = engine.close();
    await terminationStarted;
    const second = engine.close();
    expect(duckdb.terminations).toBe(1);
    releaseTermination();
    const results = await Promise.allSettled([first, second]);
    expect(results).toEqual(
      checkpointFailure === undefined
        ? [
            { status: "fulfilled", value: undefined },
            { status: "fulfilled", value: undefined },
          ]
        : [
            { status: "rejected", reason: checkpointFailure },
            { status: "rejected", reason: checkpointFailure },
          ],
    );
    await expect(engine.close()).resolves.toBeUndefined();
    expect(duckdb.connectionCloses).toBe(1);
    expect(duckdb.flushes).toBe(1);
    expect(duckdb.terminations).toBe(1);
  },
);
