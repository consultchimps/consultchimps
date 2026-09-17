import { describe, expect, it, vi } from "vitest";

import {
  closeAndRetainFailures,
  importCleanupError,
  RetryableCleanupOwners,
  savedPlanCleanupError,
  stagedPrivatePlanCleanup,
} from "./workspace-import-cleanup";

describe("workspace import cleanup", () => {
  it("attempts every source close and retains only failures for retry", async () => {
    const failure = new Error("Injected source close failure");
    const successfulClose = vi.fn(async () => undefined);
    const retryableClose = vi
      .fn<() => Promise<void>>()
      .mockRejectedValueOnce(failure)
      .mockResolvedValueOnce(undefined);
    const retained = new RetryableCleanupOwners();

    await expect(
      closeAndRetainFailures(
        [{ close: retryableClose }, { close: successfulClose }],
        retained,
      ),
    ).rejects.toBe(failure);
    expect(successfulClose).toHaveBeenCalledOnce();
    expect(retained.size).toBe(1);

    await expect(retained.close()).resolves.toBeUndefined();
    expect(retryableClose).toHaveBeenCalledTimes(2);
    expect(successfulClose).toHaveBeenCalledOnce();
    expect(retained.size).toBe(0);
  });

  it("retries only unfinished private-plan cleanup stages", async () => {
    const discardFailure = new Error("Injected plan discard failure");
    const close = vi.fn(async () => undefined);
    const discard = vi
      .fn<() => Promise<void>>()
      .mockRejectedValueOnce(discardFailure)
      .mockResolvedValueOnce(undefined);
    const owner = stagedPrivatePlanCleanup({ close, discard });

    await expect(owner.close()).rejects.toBe(discardFailure);
    await expect(owner.close()).resolves.toBeUndefined();
    expect(close).toHaveBeenCalledOnce();
    expect(discard).toHaveBeenCalledTimes(2);
  });

  it("shares a concurrent retained cleanup attempt", async () => {
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    const close = vi.fn(async () => gate);
    const retained = new RetryableCleanupOwners();
    retained.retain({ close });

    const first = retained.close();
    const second = retained.close();
    expect(second).toBe(first);
    await Promise.resolve();
    expect(close).toHaveBeenCalledOnce();
    release();
    await expect(Promise.all([first, second])).resolves.toEqual([
      undefined,
      undefined,
    ]);
    expect(retained.size).toBe(0);
  });

  it("reports completed preparation and preserves each cleanup cause", () => {
    const sourceFailure = new Error("Injected source cleanup failure");
    const planFailure = new Error("Injected plan cleanup failure");

    const error = importCleanupError({
      sourceCleanupFailure: { error: sourceFailure },
      planCleanupFailure: { error: planFailure },
      preparationCompleted: true,
    });

    expect(error).toMatchObject({
      code: "DB_BROWSER_IMPORT_CLEANUP_REQUIRED",
      details: {
        preparationCompleted: true,
        sourceCleanupFailed: true,
        planCleanupFailed: true,
      },
      cause: expect.objectContaining({
        errors: [sourceFailure, planFailure],
      }),
    });
  });

  it("preserves saved-plan inspection and retained cleanup causes", () => {
    const inspectionFailure = new Error("Injected plan inspection failure");
    const closeFailure = new Error("Injected plan close failure");

    const error = savedPlanCleanupError({
      operationFailures: [inspectionFailure],
      cleanupFailures: [closeFailure],
    });

    expect(error).toMatchObject({
      code: "DB_BROWSER_IMPORT_CLEANUP_REQUIRED",
      details: { savedPlanCleanupFailed: true, affectedPlans: 1 },
      cause: expect.objectContaining({
        errors: [inspectionFailure, closeFailure],
      }),
    });
    expect(error.message).toContain("Choose Retry saved reviews");
  });
});
