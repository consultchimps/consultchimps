import { beforeEach, expect, test, vi } from "vitest";

const xlsx = vi.hoisted(() => ({
  inspect: vi.fn(),
  open: vi.fn(),
}));

vi.mock("@consultchimps/xlsx/stream", () => ({
  inspectWorkbookStream: xlsx.inspect,
  openWorkbookStream: xlsx.open,
}));

import { createWorkbookImportSource } from "../src/workbook.js";

beforeEach(() => {
  vi.clearAllMocks();
});

const inspection = {
  sheets: [{ name: "Inventory", visibility: "visible" as const }],
  tables: [],
  namedRanges: [],
};

function options() {
  return {
    key: "inventory.xlsx",
    bytes: {
      name: "inventory.xlsx",
      size: 0,
      async readAt() {
        return new Uint8Array();
      },
    },
    scratch: {
      async create(): Promise<never> {
        throw new Error("Scratch creation is not expected");
      },
    },
  };
}

test("a rejected lazy workbook open remains the read failure and needs no cleanup retry", async () => {
  const readFailure = new AggregateError(
    [
      new Error("Injected workbook initialization failure"),
      new Error("Injected workbook initialization cleanup failure"),
    ],
    "Workbook initialization and cleanup failed",
  );
  xlsx.inspect.mockResolvedValueOnce(inspection);
  xlsx.open.mockRejectedValueOnce(readFailure);
  const workbook = await createWorkbookImportSource(options());
  const selection = workbook.source.selections[0];
  if (selection === undefined) throw new Error("Expected workbook selection");

  await expect(selection.open({})).rejects.toBe(readFailure);
  await expect(workbook.close()).resolves.toBeUndefined();
  await expect(workbook.close()).resolves.toBeUndefined();
  expect(xlsx.open).toHaveBeenCalledOnce();
});

test("concurrent closes share a failed acquired-stream attempt and retry once", async () => {
  const closeFailure = new Error("Injected workbook close failure");
  let closeStarted!: () => void;
  const started = new Promise<void>((resolve) => {
    closeStarted = resolve;
  });
  let rejectClose!: (error: unknown) => void;
  const close = vi
    .fn<() => Promise<void>>()
    .mockImplementationOnce(
      () =>
        new Promise<void>((_resolve, reject) => {
          rejectClose = reject;
          closeStarted();
        }),
    )
    .mockResolvedValueOnce(undefined);
  const readerClose = vi.fn(async () => undefined);
  xlsx.inspect.mockResolvedValueOnce(inspection);
  xlsx.open.mockResolvedValueOnce({
    inspection,
    openRegion: vi.fn(async () => ({
      region: {
        sheet: "Inventory",
        origin: { kind: "declared-header" as const },
        headerRow: 1,
        firstRow: 1,
        lastRow: 1,
        startColumn: 0,
        endColumn: 0,
        columns: [],
      },
      async *batches() {},
      close: readerClose,
    })),
    close,
  });
  const workbook = await createWorkbookImportSource(options());
  const selection = workbook.source.selections[0];
  if (selection === undefined) throw new Error("Expected workbook selection");
  const reader = await selection.open({});
  await reader.close();

  const first = workbook.close();
  const concurrent = workbook.close();
  await started;
  expect(close).toHaveBeenCalledOnce();
  rejectClose(closeFailure);
  await expect(first).rejects.toBe(closeFailure);
  await expect(concurrent).rejects.toBe(closeFailure);

  await expect(workbook.close()).resolves.toBeUndefined();
  await expect(workbook.close()).resolves.toBeUndefined();
  expect(close).toHaveBeenCalledTimes(2);
  expect(xlsx.open).toHaveBeenCalledOnce();
});

test("closing during lazy initialization releases the acquired stream without opening a region", async () => {
  let resolveOpen!: (stream: {
    readonly inspection: typeof inspection;
    openRegion(): Promise<never>;
    close(): Promise<void>;
  }) => void;
  const pending = new Promise<Parameters<typeof resolveOpen>[0]>((resolve) => {
    resolveOpen = resolve;
  });
  const close = vi.fn(async () => undefined);
  const openRegion = vi.fn(async (): Promise<never> => {
    throw new Error("A closed source must not open a region");
  });
  xlsx.inspect.mockResolvedValueOnce(inspection);
  xlsx.open.mockReturnValueOnce(pending);
  const workbook = await createWorkbookImportSource(options());
  const selection = workbook.source.selections[0];
  if (selection === undefined) throw new Error("Expected workbook selection");

  const opening = selection.open({});
  const closing = workbook.close();
  resolveOpen({ inspection, openRegion, close });

  await expect(opening).rejects.toMatchObject({ code: "DB_SOURCE_CLOSED" });
  await expect(closing).resolves.toBeUndefined();
  expect(openRegion).not.toHaveBeenCalled();
  expect(close).toHaveBeenCalledOnce();
});

test("closing during region initialization releases the new reader without returning it", async () => {
  let regionStarted!: () => void;
  const started = new Promise<void>((resolve) => {
    regionStarted = resolve;
  });
  let resolveRegion!: (reader: {
    readonly region: {
      readonly sheet: string;
      readonly origin: { readonly kind: "declared-header" };
      readonly headerRow: number;
      readonly firstRow: number;
      readonly lastRow: number;
      readonly startColumn: number;
      readonly endColumn: number;
      readonly columns: readonly [];
    };
    batches(): AsyncIterable<readonly []>;
    close(): Promise<void>;
  }) => void;
  const pendingRegion = new Promise<Parameters<typeof resolveRegion>[0]>(
    (resolve) => {
      resolveRegion = resolve;
    },
  );
  const readerClose = vi.fn(async () => undefined);
  const sessionClose = vi.fn(async () => undefined);
  const openRegion = vi.fn(() => {
    regionStarted();
    return pendingRegion;
  });
  xlsx.inspect.mockResolvedValueOnce(inspection);
  xlsx.open.mockResolvedValueOnce({
    inspection,
    openRegion,
    close: sessionClose,
  });
  const workbook = await createWorkbookImportSource(options());
  const selection = workbook.source.selections[0];
  if (selection === undefined) throw new Error("Expected workbook selection");

  const opening = selection.open({});
  await started;
  await workbook.close();
  resolveRegion({
    region: {
      sheet: "Inventory",
      origin: { kind: "declared-header" },
      headerRow: 1,
      firstRow: 1,
      lastRow: 1,
      startColumn: 0,
      endColumn: 0,
      columns: [],
    },
    async *batches() {},
    close: readerClose,
  });

  await expect(opening).rejects.toMatchObject({ code: "DB_SOURCE_CLOSED" });
  expect(readerClose).toHaveBeenCalledOnce();
  expect(sessionClose).toHaveBeenCalledOnce();
});
