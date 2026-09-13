import { expect, test, vi } from "vitest";

const resources = vi.hoisted(() => ({
  createScratchDirectory: vi.fn(),
  openRandomAccessSource: vi.fn(),
  createWorkbookImportSource: vi.fn(),
}));

vi.mock("@consultchimps/files", () => ({
  createScratchDirectory: resources.createScratchDirectory,
  openRandomAccessSource: resources.openRandomAccessSource,
}));

vi.mock("@consultchimps/db", () => ({
  createWorkbookImportSource: resources.createWorkbookImportSource,
}));

import { openDbInputs } from "./db-inputs.js";

test("preserves initialization and independently settled input cleanup failures", async () => {
  const initializationFailure = new Error("Injected workbook open failure");
  const fileFailure = new Error("Injected file close failure");
  const scratchFailure = new Error("Injected scratch close failure");
  const workbookClose = vi.fn(() => {
    throw undefined;
  });
  const firstFileClose = vi.fn(async () => {
    throw fileFailure;
  });
  const secondFileClose = vi.fn(async () => undefined);
  const scratchClose = vi.fn(async () => {
    throw scratchFailure;
  });
  resources.createScratchDirectory.mockResolvedValue({ close: scratchClose });
  resources.openRandomAccessSource
    .mockResolvedValueOnce({
      close: firstFileClose,
      verifyUnchanged: vi.fn(async () => undefined),
    })
    .mockResolvedValueOnce({
      close: secondFileClose,
      verifyUnchanged: vi.fn(async () => undefined),
    });
  resources.createWorkbookImportSource
    .mockResolvedValueOnce({ close: workbookClose, source: { key: "first" } })
    .mockRejectedValueOnce(initializationFailure);

  let failure: unknown;
  try {
    await openDbInputs(
      { input: ["first=first.xlsx", "second=second.xlsx"] },
      {},
    );
  } catch (error) {
    failure = error;
  }

  expect(workbookClose).toHaveBeenCalledOnce();
  expect(firstFileClose).toHaveBeenCalledOnce();
  expect(secondFileClose).toHaveBeenCalledOnce();
  expect(scratchClose).toHaveBeenCalledOnce();
  expect(failure).toMatchObject({
    code: "CLI_DB_INPUT_CLEANUP_REQUIRED",
    details: {
      initializationFailed: true,
      workbookCloseFailures: 1,
      fileCloseFailures: 1,
      scratchCloseFailures: 1,
    },
  });
  expect(((failure as Error).cause as AggregateError).errors).toEqual([
    initializationFailure,
    undefined,
    fileFailure,
    scratchFailure,
  ]);
});
