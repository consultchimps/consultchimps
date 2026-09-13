import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  createWorkbookImportSource: vi.fn(),
  discardPreparedImport: vi.fn(async () => undefined),
  preparedClose: vi.fn(async () => undefined),
  sourceClose: vi.fn<() => Promise<void>>(),
}));

vi.mock("@consultchimps/db", async (importOriginal) => ({
  ...(await importOriginal<typeof import("@consultchimps/db")>()),
  createWorkbookImportSource: mocks.createWorkbookImportSource,
  draftImportRecipe: vi.fn(async () => ({ version: 1, routes: [] })),
  inspectDatabase: vi.fn(async () => ({
    completedImports: 0n,
    deliveries: 0n,
    format: "sqlite",
    formatVersion: 1,
    id: "database-1",
    revision: "revision-1",
    tables: [],
  })),
  inspectImport: vi.fn(async () => ({
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
    prepared: {
      id: "plan-1",
      planRevision: "plan-revision-1",
      state: "needs-review",
    },
    routes: [
      {
        applicationState: "new",
        captureId: "capture-1",
        destination: null,
        label: "Data",
        rowCount: 1n,
        source: "source-1",
        selection: '{"sheet":"Data","headerRow":1}',
      },
    ],
  })),
  prepareImport: vi.fn(async () => ({
    prepared: {
      id: "plan-1",
      planRevision: "plan-revision-1",
      state: "needs-review",
    },
  })),
}));

vi.mock("@consultchimps/db/browser", () => ({
  configureBrowserDatabaseRuntime: vi.fn(async () => ({
    createDatabase: vi.fn(async () => ({
      database: {
        close: vi.fn(async () => undefined),
        format: "sqlite",
        id: "database-1",
      },
    })),
    createPreparedImport: vi.fn(async () => ({
      close: mocks.preparedClose,
      id: "plan-1",
    })),
    discardPreparedImport: mocks.discardPreparedImport,
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
  workspaceRecordIdPrefix: () => "DAT",
  workspaceWorkingCopyName: (value: string) => value,
}));

interface WorkerMessage {
  readonly id: number;
  readonly type: string;
  readonly code?: string;
}

describe("workspace worker import cleanup", () => {
  beforeEach(() => {
    vi.resetModules();
    vi.clearAllMocks();
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
      message: expect.stringContaining("success was not reported"),
    });
    expect(mocks.preparedClose).toHaveBeenCalledOnce();
    expect(mocks.discardPreparedImport).toHaveBeenCalledOnce();
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
});
