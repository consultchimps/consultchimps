import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  applyImport: vi.fn(),
  checkpointDatabaseWrite: vi.fn(),
  createWorkbookImportSource: vi.fn(),
  discardImportBatch: vi.fn(async () => undefined),
  preparedClose: vi.fn(async () => undefined),
  sourceClose: vi.fn<() => Promise<void>>(),
  inspectDatabase: vi.fn(),
  inspectImport: vi.fn(),
  prepareImport: vi.fn(),
}));

vi.mock("@consultchimps/db", async (importOriginal) => {
  const prepared = {
    id: "plan-1",
    planRevision: "plan-revision-1",
    reviewFingerprint: "review-fingerprint-1",
    state: "ready",
  };
  const inspection = {
    application: { state: "pending" },
    captureIds: ["capture-1"],
    capturedRows: 1n,
    conflicts: [
      {
        kind: "inferred-schema",
        source: "source-1",
        selection: '{"sheet":"Data","headerRow":1}',
        schema: {
          name: "Data",
          recordId: { prefix: "DAT", padding: 6 },
          columns: [{ name: "Value", type: "text" }],
        },
      },
    ],
    examples: [],
    prepared,
    previewWarnings: [],
    reviewRows: 1n,
    routeCount: 1n,
    routes: [
      {
        applicationState: "not-applied",
        captureId: "capture-1",
        columns: [],
        destination: {
          kind: "new-table-infer",
          name: "Data",
          recordId: { prefix: "DAT", padding: 6 },
        },
        destinationColumns: [{ name: "Value", type: "text" }],
        displayName: "source.xlsx",
        inferredColumns: [{ name: "Value", type: "text" }],
        label: "Data",
        rowCount: 1n,
        source: "source-1",
        selection: '{"sheet":"Data","headerRow":1}',
        suggestedDestination: {
          kind: "new-table-infer",
          name: "Data",
          recordId: { prefix: "DAT", padding: 6 },
        },
      },
    ],
    targetRevision: 0n,
  };
  return {
    ...(await importOriginal<typeof import("@consultchimps/db")>()),
    applyImport: mocks.applyImport,
    checkpointDatabaseWrite: mocks.checkpointDatabaseWrite,
    createWorkbookImportSource: mocks.createWorkbookImportSource,
    draftImportProfile: vi.fn(async () => ({
      version: 1,
      routes: [],
    })),
    inspectDatabase: mocks.inspectDatabase,
    inspectImport: mocks.inspectImport,
    prepareImport: mocks.prepareImport.mockImplementation(async () => ({
      prepared,
      inspection,
    })),
  };
});

vi.mock("@consultchimps/db/browser", () => ({
  configureBrowserDatabaseRuntime: vi.fn(async () => ({
    createDatabase: vi.fn(async () => ({
      database: {
        close: vi.fn(async () => undefined),
        format: "sqlite",
        id: "database-1",
      },
    })),
    createImportBatch: vi.fn(async () => ({
      close: mocks.preparedClose,
      id: "plan-1",
    })),
    discardImportBatch: mocks.discardImportBatch,
  })),
}));

vi.mock("@/lib/shared", () => ({ basePath: "" }));
vi.mock("@/lib/workspace-files", () => ({
  BrowserBlobSource: class BrowserBlobSource {
    constructor(
      readonly name: string,
      readonly file: File,
    ) {}
  },
  BrowserOpfsFile: class BrowserOpfsFile {},
  browserScratchFactory: {},
  createBrowserExportName: vi.fn(),
  removeOpfsFile: vi.fn(),
  withBrowserExportLease: vi.fn(),
}));
vi.mock("@/lib/workspace-source", () => ({
  workspaceSourceDescription: (value: string) => value,
  workspaceSourceFileName: (value: string) => value,
  workspaceSourceKey: () => "source-1",
}));
vi.mock(
  "@/lib/workspace-replacement",
  async () => import("../lib/workspace-replacement"),
);
vi.mock("@/lib/workspace-naming", () => ({
  workspaceWorkingCopyName: (value: string) => value,
}));

interface WorkerMessage {
  readonly id: number;
  readonly type: string;
  readonly code?: string;
}

function importInspection() {
  return {
    application: { state: "pending" },
    captureIds: ["capture-1"],
    capturedRows: 1n,
    conflicts: [],
    examples: [],
    prepared: {
      id: "plan-1",
      planRevision: "plan-revision-1",
      reviewFingerprint: "review-fingerprint-1",
      state: "ready",
    },
    previewWarnings: [],
    reviewRows: 1n,
    routeCount: 0n,
    routes: [],
    targetRevision: 0n,
  };
}

describe("workspace worker import cleanup", () => {
  beforeEach(() => {
    vi.resetModules();
    vi.clearAllMocks();
    mocks.inspectDatabase.mockResolvedValue({
      completedImports: 0n,
      recordedBatches: 0n,
      format: "sqlite",
      formatVersion: 1,
      id: "database-1",
      revision: 0n,
      tables: [],
    });
    mocks.inspectImport.mockResolvedValue(importInspection());
  });

  it("reports cleanup failure instead of success and retries retained owners", async () => {
    const cleanupFailure = new Error("Injected workbook source close failure");
    mocks.sourceClose
      .mockRejectedValueOnce(cleanupFailure)
      .mockResolvedValue(undefined);
    mocks.createWorkbookImportSource.mockImplementation(async () => ({
      close: mocks.sourceClose,
      inspection: { namedRanges: [], sheets: [], tables: [] },
      source: {
        key: "source-1",
        selections: [
          {
            key: '{"sheet":"Data","headerRow":1}',
            label: "Data",
          },
        ],
      },
    }));

    let receive!: (event: MessageEvent<Record<string, unknown>>) => void;
    const waiters = new Map<number, (message: WorkerMessage) => void>();
    const messages: WorkerMessage[] = [];
    vi.stubGlobal("self", {
      addEventListener(
        _type: "message",
        listener: (event: MessageEvent<Record<string, unknown>>) => void,
      ) {
        receive = listener;
      },
      postMessage(message: WorkerMessage) {
        messages.push(message);
        if (message.type !== "progress") waiters.get(message.id)?.(message);
      },
    });
    await import("./workspace.worker");
    const request = (command: Record<string, unknown>) =>
      new Promise<WorkerMessage>((resolve) => {
        waiters.set(command["id"] as number, resolve);
        receive({ data: command } as MessageEvent<Record<string, unknown>>);
      });

    await expect(
      request({ id: 1, type: "create", name: "test.sqlite", format: "sqlite" }),
    ).resolves.toMatchObject({ id: 1, type: "ready" });
    const source = { file: new File(["synthetic"], "source.xlsx") };
    await expect(
      request({ id: 2, type: "prepareImport", sources: [source] }),
    ).resolves.toEqual({
      id: 2,
      type: "error",
      code: "DB_BROWSER_IMPORT_CLEANUP_REQUIRED",
      message: expect.stringContaining("could not finish cleanup"),
    });
    expect(mocks.preparedClose).toHaveBeenCalledOnce();
    expect(mocks.discardImportBatch).toHaveBeenCalledOnce();
    await Promise.resolve();
    expect(messages.filter((message) => message.id === 2)).toEqual([
      expect.objectContaining({
        code: "DB_BROWSER_IMPORT_CLEANUP_REQUIRED",
        type: "error",
      }),
    ]);

    await expect(
      request({ id: 3, type: "prepareImport", sources: [source] }),
    ).resolves.toMatchObject({ id: 3, type: "importPrepared" });
    expect(mocks.sourceClose).toHaveBeenCalledTimes(3);
    expect(mocks.createWorkbookImportSource).toHaveBeenCalledTimes(2);
    expect(mocks.sourceClose.mock.invocationCallOrder[1]).toBeLessThan(
      mocks.createWorkbookImportSource.mock.invocationCallOrder[1]!,
    );
  });

  it("reports a discarded private batch when its committed preparation cannot checkpoint", async () => {
    const { ConsultChimpsError } = await import("@consultchimps/core");
    mocks.sourceClose.mockResolvedValue(undefined);
    mocks.createWorkbookImportSource.mockResolvedValue({
      close: mocks.sourceClose,
      inspection: { namedRanges: [], sheets: [], tables: [] },
      source: {
        key: "source-1",
        selections: [{ key: '{"sheet":"Data","headerRow":1}', label: "Data" }],
      },
    });
    const checkpointFailure = new ConsultChimpsError(
      "DB_BATCH_CHECKPOINT_REQUIRED",
      "Injected committed batch checkpoint failure.",
      {
        details: {
          batchUpdated: true,
          checkpointRequired: true,
          batchId: "plan-1",
          planRevision: "2",
        },
      },
    );
    mocks.prepareImport.mockRejectedValueOnce(checkpointFailure);

    let receive!: (event: MessageEvent<Record<string, unknown>>) => void;
    const waiters = new Map<number, (message: WorkerMessage) => void>();
    vi.stubGlobal("self", {
      addEventListener(
        _type: "message",
        listener: (event: MessageEvent<Record<string, unknown>>) => void,
      ) {
        receive = listener;
      },
      postMessage(message: WorkerMessage) {
        if (message.type !== "progress") waiters.get(message.id)?.(message);
      },
    });
    await import("./workspace.worker");
    const request = (command: Record<string, unknown>) =>
      new Promise<WorkerMessage>((resolve) => {
        waiters.set(command["id"] as number, resolve);
        receive({ data: command } as MessageEvent<Record<string, unknown>>);
      });

    await request({
      id: 1,
      type: "create",
      name: "test.sqlite",
      format: "sqlite",
    });
    await expect(
      request({
        id: 2,
        type: "prepareImport",
        sources: [{ file: new File(["synthetic"], "source.xlsx") }],
      }),
    ).resolves.toMatchObject({
      type: "error",
      code: "DB_IMPORT_PREPARATION_DISCARDED",
      message: expect.stringContaining("working database was not changed"),
    });
    expect(mocks.preparedClose).toHaveBeenCalledOnce();
    expect(mocks.discardImportBatch).toHaveBeenCalledOnce();
  });

  it("reports a committed apply when summary refresh fails and still checkpoints", async () => {
    mocks.createWorkbookImportSource.mockResolvedValue({
      close: mocks.sourceClose,
      inspection: { namedRanges: [], sheets: [], tables: [] },
      source: {
        key: "source-1",
        selections: [{ key: '{"sheet":"Data","headerRow":1}', label: "Data" }],
      },
    });
    mocks.applyImport.mockResolvedValue({
      operation: "db.import.apply",
      databaseWrite: "committed",
      importIds: ["import-1"],
      batchId: "batch-1",
      captureIds: ["capture-1"],
      metrics: {
        batchesRecorded: 1,
        rowsImported: 1,
        rowsReused: 0,
        tablesCreated: 1,
      },
    });
    mocks.checkpointDatabaseWrite.mockResolvedValue({
      checkpoint: { state: "checkpoint-completed" },
    });

    let receive!: (event: MessageEvent<Record<string, unknown>>) => void;
    const waiters = new Map<number, (message: WorkerMessage) => void>();
    vi.stubGlobal("self", {
      addEventListener(
        _type: "message",
        listener: (event: MessageEvent<Record<string, unknown>>) => void,
      ) {
        receive = listener;
      },
      postMessage(message: WorkerMessage) {
        if (message.type !== "progress") waiters.get(message.id)?.(message);
      },
    });
    await import("./workspace.worker");
    const request = (command: Record<string, unknown>) =>
      new Promise<WorkerMessage>((resolve) => {
        waiters.set(command["id"] as number, resolve);
        receive({ data: command } as MessageEvent<Record<string, unknown>>);
      });

    await request({
      id: 1,
      type: "create",
      name: "test.sqlite",
      format: "sqlite",
    });
    const source = { file: new File(["synthetic"], "source.xlsx") };
    await expect(
      request({ id: 2, type: "prepareImport", sources: [source] }),
    ).resolves.toMatchObject({ type: "importPrepared" });
    mocks.inspectDatabase.mockRejectedValueOnce(
      new Error("Injected summary refresh failure"),
    );

    await expect(
      request({
        id: 3,
        type: "applyImport",
        planId: "plan-1",
        reviewFingerprint: "review-fingerprint-1",
        batch: {
          requestId: "request-1",
          vendor: "Vendor",
          entity: "Entity",
          phase: "Phase",
          coverage: "full",
          effectiveDate: null,
          receivedDate: null,
          note: "",
        },
      }),
    ).resolves.toMatchObject({
      type: "importApplied",
      result: {
        checkpoint: { state: "saved" },
        summary: {
          state: "refresh-required",
          code: "DB_BROWSER_SUMMARY_REFRESH_REQUIRED",
        },
      },
    });
    expect(mocks.applyImport).toHaveBeenCalledOnce();
    expect(mocks.checkpointDatabaseWrite).toHaveBeenCalledOnce();
  });

  it("rejects a changed displayed review before applying it", async () => {
    mocks.createWorkbookImportSource.mockResolvedValue({
      close: mocks.sourceClose,
      inspection: { namedRanges: [], sheets: [], tables: [] },
      source: {
        key: "source-1",
        selections: [{ key: '{"sheet":"Data","headerRow":1}', label: "Data" }],
      },
    });

    let receive!: (event: MessageEvent<Record<string, unknown>>) => void;
    const waiters = new Map<number, (message: WorkerMessage) => void>();
    vi.stubGlobal("self", {
      addEventListener(
        _type: "message",
        listener: (event: MessageEvent<Record<string, unknown>>) => void,
      ) {
        receive = listener;
      },
      postMessage(message: WorkerMessage) {
        if (message.type !== "progress") waiters.get(message.id)?.(message);
      },
    });
    await import("./workspace.worker");
    const request = (command: Record<string, unknown>) =>
      new Promise<WorkerMessage>((resolve) => {
        waiters.set(command["id"] as number, resolve);
        receive({ data: command } as MessageEvent<Record<string, unknown>>);
      });

    await request({
      id: 1,
      type: "create",
      name: "test.sqlite",
      format: "sqlite",
    });
    await request({
      id: 2,
      type: "prepareImport",
      sources: [{ file: new File(["synthetic"], "source.xlsx") }],
    });
    mocks.inspectImport.mockResolvedValue({
      ...importInspection(),
      prepared: {
        ...importInspection().prepared,
        reviewFingerprint: "changed-review-fingerprint",
      },
    });

    await expect(
      request({
        id: 3,
        type: "applyImport",
        planId: "plan-1",
        reviewFingerprint: "review-fingerprint-1",
        batch: {
          requestId: "request-1",
          vendor: "Vendor",
          entity: "Entity",
          phase: "Phase",
          coverage: "full",
          effectiveDate: null,
          receivedDate: null,
          note: "",
        },
      }),
    ).resolves.toMatchObject({ type: "error", code: "DB_STALE_IMPORT_PLAN" });
    expect(mocks.applyImport).not.toHaveBeenCalled();
    expect(mocks.checkpointDatabaseWrite).not.toHaveBeenCalled();
  });
});
