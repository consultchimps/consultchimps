import { expect, test, vi } from "vitest";

import { RetryableClose } from "../src/internal/retryable-close.js";

test("retains ownership until release succeeds and shares concurrent retries", async () => {
  const failure = new Error("engine busy");
  const pending = Promise.withResolvers<void>();
  const release = vi
    .fn<() => Promise<void>>()
    .mockRejectedValueOnce(failure)
    .mockImplementationOnce(() => pending.promise);
  const owner = new RetryableClose(release);
  await expect(owner.close()).rejects.toBe(failure);
  expect(owner.isOpen).toBe(true);
  const retry = owner.close();
  expect(owner.close()).toBe(retry);
  expect(owner.isOpen).toBe(true);
  pending.resolve();
  await retry;
  expect(owner.isOpen).toBe(false);
  await owner.close();
  expect(release).toHaveBeenCalledTimes(2);
});

test("retains an owner when release throws synchronously", async () => {
  const failure = new Error("release failure");
  const owner = new RetryableClose(() => {
    throw failure;
  });
  await expect(owner.close()).rejects.toBe(failure);
  expect(owner.isOpen).toBe(true);
  await expect(owner.close()).rejects.toBe(failure);
});
