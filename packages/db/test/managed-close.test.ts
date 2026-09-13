import { describe, expect, test } from "vitest";

import {
  createDatabaseHandle,
  engineOf,
  type DatabaseId,
} from "../src/database.js";
import type {
  DatabaseEngine,
  EngineRow,
  EngineTransaction,
} from "../src/internal/engine.js";
import {
  createPreparedImportHandle,
  preparedEngineOf,
} from "../src/prepared.js";

interface ManagedHandle {
  readonly isOpen: boolean;
  close(): Promise<void>;
}

class ControlledCloseEngine implements DatabaseEngine {
  readonly format = "sqlite" as const;
  readonly interruptible = true;
  closeCalls = 0;
  nextCloseFailure: unknown;
  closeGate: Promise<void> | undefined;

  async execute(): Promise<void> {}

  async query(): Promise<readonly EngineRow[]> {
    return [];
  }

  async bulkInsert(): Promise<void> {}

  async transaction<T>(
    work: (transaction: EngineTransaction) => Promise<T>,
  ): Promise<T> {
    return work(this);
  }

  async readTransaction<T>(
    work: (transaction: EngineTransaction) => Promise<T>,
  ): Promise<T> {
    return work(this);
  }

  async checkpoint(): Promise<void> {}

  interrupt(): void {}

  async close(): Promise<void> {
    this.closeCalls += 1;
    await this.closeGate;
    const failure = this.nextCloseFailure;
    this.nextCloseFailure = undefined;
    if (failure !== undefined) throw failure;
  }
}

async function verifyRetryableClose(
  create: (engine: ControlledCloseEngine) => Promise<{
    readonly handle: ManagedHandle;
    registered(): DatabaseEngine;
  }>,
): Promise<void> {
  const engine = new ControlledCloseEngine();
  const managed = await create(engine);
  const failure = new Error("Injected handle close failure");
  engine.nextCloseFailure = failure;

  await expect(managed.handle.close()).rejects.toBe(failure);
  expect(managed.handle.isOpen).toBe(true);
  expect(managed.registered()).toBe(engine);

  await expect(managed.handle.close()).resolves.toBeUndefined();
  expect(engine.closeCalls).toBe(2);
  expect(managed.handle.isOpen).toBe(false);
  expect(managed.registered).toThrow();
}

async function verifyConcurrentClose(
  create: (engine: ControlledCloseEngine) => Promise<ManagedHandle>,
): Promise<void> {
  const engine = new ControlledCloseEngine();
  let release = (): void => undefined;
  engine.closeGate = new Promise<void>((resolve) => {
    release = resolve;
  });
  const handle = await create(engine);

  const first = handle.close();
  const second = handle.close();
  expect(first).toBe(second);
  await Promise.resolve();
  expect(engine.closeCalls).toBe(1);
  expect(handle.isOpen).toBe(true);

  release();
  await Promise.all([first, second]);
  expect(engine.closeCalls).toBe(1);
  expect(handle.isOpen).toBe(false);
}

describe("managed handle closure", () => {
  test("keeps a database registered until a retry confirms closure", async () => {
    await verifyRetryableClose(async (engine) => {
      const handle = await createDatabaseHandle(engine);
      return { handle, registered: () => engineOf(handle) };
    });
  });

  test("shares a concurrent database close attempt", async () => {
    await verifyConcurrentClose((engine) => createDatabaseHandle(engine));
  });

  test("keeps a prepared import registered until a retry confirms closure", async () => {
    await verifyRetryableClose(async (engine) => {
      const handle = await createPreparedImportHandle({
        engine,
        databaseId: "DB-managed-close" as DatabaseId,
        baselineRevision: 0n,
        baselineSchemaFingerprint: "schema-fingerprint",
        recipe: { version: 1, routes: [] },
      });
      return { handle, registered: () => preparedEngineOf(handle) };
    });
  });

  test("shares a concurrent prepared-import close attempt", async () => {
    await verifyConcurrentClose((engine) =>
      createPreparedImportHandle({
        engine,
        databaseId: "DB-managed-close" as DatabaseId,
        baselineRevision: 0n,
        baselineSchemaFingerprint: "schema-fingerprint",
        recipe: { version: 1, routes: [] },
      }),
    );
  });
});
