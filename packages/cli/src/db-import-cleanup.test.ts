import { ConsultChimpsError } from "@consultchimps/core";
import { expect, test, vi } from "vitest";

import { finishCliImport } from "./db-import-cleanup.js";

test("skips private batch removal after its handle fails to close", async () => {
  const closeInputs = vi.fn(async () => undefined);
  const closeDatabase = vi.fn(async () => undefined);
  const removeTemporary = vi.fn(async () => undefined);

  let failure: unknown;
  try {
    await finishCliImport({
      outcome: { status: "failed", error: undefined },
      temporaryPath: "/synthetic/cc-import-plan-private",
      closePrepared() {
        throw undefined;
      },
      closeInputs,
      closeDatabase,
      removeTemporary,
    });
  } catch (error) {
    failure = error;
  }

  expect(closeInputs).toHaveBeenCalledOnce();
  expect(closeDatabase).toHaveBeenCalledOnce();
  expect(removeTemporary).not.toHaveBeenCalled();
  expect(failure).toMatchObject({
    code: "CLI_DB_IMPORT_CLEANUP_REQUIRED",
    details: {
      temporaryPath: "/synthetic/cc-import-plan-private",
      preparedCloseFailed: true,
      inputCloseFailed: false,
      databaseCloseFailed: false,
      removalFailed: false,
      removalSkipped: true,
      operationCompleted: false,
    },
  });
  expect(((failure as Error).cause as AggregateError).errors).toEqual([
    undefined,
    undefined,
  ]);
});

test("removes a closed private batch even when the database close fails", async () => {
  const databaseFailure = new Error("Injected database close failure");
  const removeTemporary = vi.fn(async () => undefined);

  await expect(
    finishCliImport({
      outcome: { status: "completed" },
      temporaryPath: "/synthetic/cc-import-plan-private",
      closePrepared: async () => undefined,
      closeInputs: async () => undefined,
      async closeDatabase() {
        throw databaseFailure;
      },
      removeTemporary,
    }),
  ).rejects.toMatchObject({
    code: "CLI_DB_IMPORT_CLEANUP_REQUIRED",
    details: {
      preparedCloseFailed: false,
      databaseCloseFailed: true,
      removalFailed: false,
      removalSkipped: false,
      operationCompleted: true,
    },
    cause: expect.objectContaining({ errors: [databaseFailure] }),
  });
  expect(removeTemporary).toHaveBeenCalledOnce();
});

test("reports a discarded batch after checkpoint failure and successful cleanup", async () => {
  const checkpointFailure = new ConsultChimpsError(
    "DB_BATCH_CHECKPOINT_REQUIRED",
    "Injected committed batch checkpoint failure.",
  );

  await expect(
    finishCliImport({
      outcome: { status: "failed", error: checkpointFailure },
      temporaryPath: "/synthetic/cc-import-plan-private",
      closePrepared: async () => undefined,
      closeInputs: async () => undefined,
      closeDatabase: async () => undefined,
      removeTemporary: async () => undefined,
    }),
  ).rejects.toMatchObject({
    code: "DB_IMPORT_PREPARATION_DISCARDED",
    details: { batchDiscarded: true, batchPublished: false },
    cause: checkpointFailure,
  });
});

test("keeps a caller-owned checkpoint failure when there is no private batch", async () => {
  const checkpointFailure = new ConsultChimpsError(
    "DB_BATCH_CHECKPOINT_REQUIRED",
    "Injected committed batch checkpoint failure.",
  );

  await expect(
    finishCliImport({
      outcome: { status: "failed", error: checkpointFailure },
      closeInputs: async () => undefined,
      closeDatabase: async () => undefined,
      removeTemporary: async () => undefined,
    }),
  ).rejects.toBe(checkpointFailure);
});

test("preserves operation and removal failures after confirmed closes", async () => {
  const operationFailure = new ConsultChimpsError(
    "DB_BATCH_CHECKPOINT_REQUIRED",
    "Injected committed batch checkpoint failure.",
  );
  const removalFailure = new Error("Injected directory removal failure");

  let failure: unknown;
  try {
    await finishCliImport({
      outcome: { status: "failed", error: operationFailure },
      temporaryPath: "/synthetic/cc-import-plan-private",
      closePrepared: async () => undefined,
      closeInputs: async () => undefined,
      closeDatabase: async () => undefined,
      async removeTemporary() {
        throw removalFailure;
      },
    });
  } catch (error) {
    failure = error;
  }

  expect(failure).toMatchObject({
    code: "CLI_DB_IMPORT_CLEANUP_REQUIRED",
    details: {
      removalFailed: true,
      removalSkipped: false,
    },
  });
  expect(((failure as Error).cause as AggregateError).errors).toEqual([
    operationFailure,
    removalFailure,
  ]);
});
