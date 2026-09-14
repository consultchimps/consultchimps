import { expect, test, vi } from "vitest";

import {
  BrowserOpenCleanupError,
  BrowserOpenCleanupRegistry,
} from "../src/browser-open-cleanup.js";

test("retains only owners whose cleanup retry remains unconfirmed", async () => {
  const registry = new BrowserOpenCleanupRegistry();
  const primary = new Error("Injected open failure");
  const firstClose = new Error("Injected initial close failure");
  const retryClose = vi
    .fn<() => Promise<void>>()
    .mockRejectedValueOnce(undefined)
    .mockResolvedValueOnce(undefined);
  const failure = registry.retain("review.sqlite", "prepared", retryClose, [
    primary,
    firstClose,
  ]);

  expect(failure).toBeInstanceOf(BrowserOpenCleanupError);
  expect(registry.has("review.sqlite")).toBe(true);
  let retryFailure: unknown;
  try {
    await registry.retry("review.sqlite");
  } catch (error) {
    retryFailure = error;
  }
  expect(retryFailure).toMatchObject({
    storageName: "review.sqlite",
    ownerKinds: ["prepared"],
    cleanupCauses: [primary, firstClose, undefined],
  });
  expect(registry.has("review.sqlite")).toBe(true);

  await registry.retry("review.sqlite");
  expect(registry.has("review.sqlite")).toBe(false);
  await registry.retry("review.sqlite");
  expect(retryClose).toHaveBeenCalledTimes(2);
});

test("shares a concurrent cleanup attempt without duplicating its cause", async () => {
  const registry = new BrowserOpenCleanupRegistry();
  const pending = Promise.withResolvers<void>();
  const cleanupFailure = new Error("Injected retry failure");
  const retryClose = vi.fn(async () => {
    await pending.promise;
    throw cleanupFailure;
  });
  registry.retain("review.sqlite", "prepared", retryClose, []);

  const first = registry.retry("review.sqlite");
  const second = registry.retry("review.sqlite");
  pending.resolve();

  const failures = await Promise.all([
    first.catch((error: unknown) => error),
    second.catch((error: unknown) => error),
  ]);
  expect(retryClose).toHaveBeenCalledOnce();
  expect(failures).toEqual([
    expect.objectContaining({ cleanupCauses: [cleanupFailure] }),
    expect.objectContaining({ cleanupCauses: [cleanupFailure] }),
  ]);
});
