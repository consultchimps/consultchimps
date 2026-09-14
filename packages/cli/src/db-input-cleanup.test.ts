import { expect, test, vi } from "vitest";

import {
  createDbInputCloser,
  dbInputInitializationCleanupError,
} from "./db-input-cleanup.js";

test("settles workbook, file, and scratch cleanup and preserves each cause", async () => {
  const laterWorkbook = vi.fn(async () => undefined);
  const fileFailure = new Error("Injected file close failure");
  const scratchFailure = new Error("Injected scratch close failure");
  const close = createDbInputCloser({
    workbooks: [
      () => {
        throw undefined;
      },
      laterWorkbook,
    ],
    files: [
      async () => {
        throw fileFailure;
      },
    ],
    scratch: async () => {
      throw scratchFailure;
    },
  });

  let failure: unknown;
  try {
    await close();
  } catch (error) {
    failure = error;
  }

  expect(laterWorkbook).toHaveBeenCalledOnce();
  expect(failure).toMatchObject({
    code: "CLI_DB_INPUT_CLEANUP_REQUIRED",
    details: {
      workbookCloseFailures: 1,
      fileCloseFailures: 1,
      scratchCloseFailures: 1,
    },
  });
  expect(((failure as Error).cause as AggregateError).errors).toEqual([
    undefined,
    fileFailure,
    scratchFailure,
  ]);
});

test("shares concurrent cleanup and does not repeat successful closes", async () => {
  let release!: () => void;
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  const workbook = vi.fn(async () => blocked);
  const file = vi.fn(async () => undefined);
  const scratch = vi.fn(async () => undefined);
  const close = createDbInputCloser({
    workbooks: [workbook],
    files: [file],
    scratch,
  });

  const first = close();
  const second = close();
  release();
  await Promise.all([first, second]);
  await close();

  expect(workbook).toHaveBeenCalledOnce();
  expect(file).toHaveBeenCalledOnce();
  expect(scratch).toHaveBeenCalledOnce();
});

test("retries only resources whose earlier close failed", async () => {
  const workbook = vi
    .fn<() => Promise<void>>()
    .mockRejectedValueOnce(new Error("Injected workbook close failure"))
    .mockResolvedValueOnce(undefined);
  const file = vi.fn(async () => undefined);
  const scratch = vi.fn(async () => undefined);
  const close = createDbInputCloser({
    workbooks: [workbook],
    files: [file],
    scratch,
  });

  await expect(close()).rejects.toMatchObject({
    code: "CLI_DB_INPUT_CLEANUP_REQUIRED",
  });
  await close();

  expect(workbook).toHaveBeenCalledTimes(2);
  expect(file).toHaveBeenCalledOnce();
  expect(scratch).toHaveBeenCalledOnce();
});

test("keeps initialization failure first when cleanup also fails", () => {
  const primary = new Error("Injected workbook open failure");
  const cleanupCause = new Error("Injected scratch close failure");
  const cleanup = new AggregateError([cleanupCause], "cleanup failed");

  const failure = dbInputInitializationCleanupError(primary, cleanup);

  expect(failure.code).toBe("CLI_DB_INPUT_CLEANUP_REQUIRED");
  expect((failure.cause as AggregateError).errors).toEqual([
    primary,
    cleanupCause,
  ]);
});
