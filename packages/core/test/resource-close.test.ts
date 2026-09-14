import { expect, test, vi } from "vitest";

import { OwnedResources } from "../src/resource-close.js";

test("retains failed owners and retries only resources that remain open", async () => {
  const failure = new Error("busy");
  const first = { close: vi.fn(async () => {}) };
  const second = {
    close: vi
      .fn<() => Promise<void>>()
      .mockRejectedValueOnce(failure)
      .mockResolvedValue(undefined),
  };
  const resources = new OwnedResources([first, second]);
  expect(await resources.close()).toEqual([
    { resource: second, error: failure },
  ]);
  expect(resources.size).toBe(1);
  expect(await resources.close()).toEqual([]);
  expect(resources.size).toBe(0);
  expect(first.close).toHaveBeenCalledTimes(1);
  expect(second.close).toHaveBeenCalledTimes(2);
});

test("concurrent calls share an attempt and wait for independently closing owners", async () => {
  const pending = Promise.withResolvers<void>();
  const slow = { close: vi.fn(() => pending.promise) };
  const fast = { close: vi.fn(async () => {}) };
  const resources = new OwnedResources([slow, fast]);
  const first = resources.close();
  expect(resources.close()).toBe(first);
  await Promise.resolve();
  expect(slow.close).toHaveBeenCalledTimes(1);
  expect(fast.close).toHaveBeenCalledTimes(1);
  pending.resolve();
  expect(await first).toEqual([]);
  expect(resources.size).toBe(0);
});

test("retains owners registered during an attempt for the next explicit close", async () => {
  const pending = Promise.withResolvers<void>();
  const first = { close: vi.fn(() => pending.promise) };
  const later = { close: vi.fn(async () => {}) };
  const resources = new OwnedResources([first]);
  const closing = resources.close();
  resources.add(later);
  pending.resolve();
  expect(await closing).toEqual([]);
  expect(later.close).not.toHaveBeenCalled();
  expect(resources.size).toBe(1);
  expect(await resources.close()).toEqual([]);
  expect(later.close).toHaveBeenCalledTimes(1);
});

test("settles synchronous throws without skipping other resources or duplicating owners", async () => {
  const failure = new Error("synchronous close failure");
  const failed = {
    close: vi.fn<() => Promise<void>>(() => {
      throw failure;
    }),
  };
  const closed = { close: vi.fn(async () => {}) };
  const resources = new OwnedResources([failed, closed, failed]);
  expect(await resources.close()).toEqual([
    { resource: failed, error: failure },
  ]);
  expect(failed.close).toHaveBeenCalledTimes(1);
  expect(closed.close).toHaveBeenCalledTimes(1);
  expect(resources.size).toBe(1);
});
