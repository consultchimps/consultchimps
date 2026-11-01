import { describe, expect, it } from "vitest";

import {
  claimDatabaseTool,
  DATABASE_TOOL_LOCK_NAME,
  type OwnershipLockManager,
  waitForDatabaseTool,
} from "./workspace-ownership";

interface Waiter {
  readonly grant: () => void;
  readonly abort: () => void;
}

/**
 * A lock manager with the subset of Web Locks semantics the claims rely on:
 * one exclusive holder per name, `ifAvailable` answering immediately, a
 * waiting request granted in order once the holder's callback settles, and
 * an aborted wait rejecting the request.
 */
class FakeLockManager implements OwnershipLockManager {
  readonly requests: string[] = [];
  #holder: Promise<void> | null = null;
  readonly #waiters: Waiter[] = [];

  get held(): boolean {
    return this.#holder !== null;
  }

  async request<T>(
    name: string,
    options: {
      readonly mode: "exclusive";
      readonly ifAvailable?: boolean;
      readonly signal?: AbortSignal;
    },
    callback: (lock: object | null) => Promise<T>,
  ): Promise<T> {
    this.requests.push(name);
    if (this.#holder !== null) {
      if (options.ifAvailable === true) return callback(null);
      await new Promise<void>((resolve, reject) => {
        const waiter: Waiter = {
          grant: resolve,
          abort: () => reject(new DOMException("aborted", "AbortError")),
        };
        this.#waiters.push(waiter);
        options.signal?.addEventListener("abort", () => {
          const index = this.#waiters.indexOf(waiter);
          if (index >= 0) this.#waiters.splice(index, 1);
          waiter.abort();
        });
      });
    }
    const running = callback({ name });
    this.#holder = running.then(
      () => undefined,
      () => undefined,
    );
    try {
      return await running;
    } finally {
      this.#holder = null;
      this.#waiters.shift()?.grant();
    }
  }

  /** Let a release, or the lack of one, propagate to waiting requests. */
  async settle(): Promise<void> {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

describe("claimDatabaseTool", () => {
  it("owns the tool and holds the lock until released", async () => {
    const locks = new FakeLockManager();
    const ownership = await claimDatabaseTool(locks);
    expect(ownership.state).toBe("owned");
    expect(locks.requests).toEqual([DATABASE_TOOL_LOCK_NAME]);
    expect(locks.held).toBe(true);
    if (ownership.state !== "owned") throw new Error("unreachable");
    ownership.release();
    await locks.settle();
    expect(locks.held).toBe(false);
  });

  it("reports a conflict while another page holds the lock", async () => {
    const locks = new FakeLockManager();
    const first = await claimDatabaseTool(locks);
    const second = await claimDatabaseTool(locks);
    expect(second).toEqual({ state: "conflict" });
    if (first.state !== "owned") throw new Error("unreachable");
    first.release();
    await locks.settle();
    const third = await claimDatabaseTool(locks);
    expect(third.state).toBe("owned");
  });

  it("reports the lock manager as unavailable when the browser has none", async () => {
    await expect(claimDatabaseTool(null)).resolves.toEqual({
      state: "unavailable",
    });
  });

  it("treats a lock manager that throws on request as unavailable", async () => {
    const locks: OwnershipLockManager = {
      request: () => {
        throw new TypeError("navigator.locks.request is not a function");
      },
    };
    await expect(claimDatabaseTool(locks)).resolves.toEqual({
      state: "unavailable",
    });
    await expect(
      waitForDatabaseTool(new AbortController().signal, locks),
    ).resolves.toEqual({ state: "unavailable" });
  });

  it("treats a refused lock request as unavailable", async () => {
    const locks: OwnershipLockManager = {
      request: async () => {
        throw new DOMException("Locks are disabled", "SecurityError");
      },
    };
    await expect(claimDatabaseTool(locks)).resolves.toEqual({
      state: "unavailable",
    });
  });
});

describe("waitForDatabaseTool", () => {
  it("takes over once the previous owner releases", async () => {
    const locks = new FakeLockManager();
    const first = await claimDatabaseTool(locks);
    if (first.state !== "owned") throw new Error("unreachable");
    let settled = false;
    const waiting = waitForDatabaseTool(new AbortController().signal, locks);
    void waiting.then(() => {
      settled = true;
    });
    await locks.settle();
    expect(settled).toBe(false);

    first.release();
    const second = await waiting;
    expect(second.state).toBe("owned");
    expect(locks.held).toBe(true);
    if (second.state !== "owned") throw new Error("unreachable");
    second.release();
    await locks.settle();
    expect(locks.held).toBe(false);
  });

  it("gives up without taking the lock when aborted", async () => {
    const locks = new FakeLockManager();
    const first = await claimDatabaseTool(locks);
    if (first.state !== "owned") throw new Error("unreachable");
    const controller = new AbortController();
    const waiting = waitForDatabaseTool(controller.signal, locks);
    controller.abort();
    await expect(waiting).resolves.toEqual({ state: "unavailable" });

    first.release();
    await locks.settle();
    expect(locks.held).toBe(false);
  });

  it("does not wait when already aborted or without a lock manager", async () => {
    const controller = new AbortController();
    controller.abort();
    await expect(
      waitForDatabaseTool(controller.signal, new FakeLockManager()),
    ).resolves.toEqual({ state: "unavailable" });
    await expect(
      waitForDatabaseTool(new AbortController().signal, null),
    ).resolves.toEqual({ state: "unavailable" });
  });
});
