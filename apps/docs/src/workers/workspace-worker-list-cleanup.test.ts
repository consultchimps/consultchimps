import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  databaseClose: vi.fn(async () => undefined),
  firstClose: vi.fn<() => Promise<void>>(),
  inspectAppliedImportBatch: vi.fn(async () => null),
  inspectImport: vi.fn(),
  listImportBatches: vi.fn(),
  openImportBatch: vi.fn(),
  secondClose: vi.fn<() => Promise<void>>(),
}));

vi.mock("@consultchimps/db", async (importOriginal) => ({
  ...(await importOriginal<typeof import("@consultchimps/db")>()),
  inspectAppliedImportBatch: mocks.inspectAppliedImportBatch,
  inspectDatabase: vi.fn(async () => ({
    completedImports: 0n,
    recordedBatches: 0n,
    format: "sqlite",
    formatVersion: 1,
    id: "database-1",
    revision: 0n,
    tables: [],
  })),
  inspectImport: mocks.inspectImport,
}));

vi.mock("@consultchimps/db/browser", () => ({
  configureBrowserDatabaseRuntime: vi.fn(async () => ({
    createDatabase: vi.fn(async () => ({
      database: {
        close: mocks.databaseClose,
        format: "sqlite",
        id: "database-1",
      },
    })),
    listImportBatches: mocks.listImportBatches,
    openImportBatch: mocks.openImportBatch,
  })),
}));

vi.mock("@/lib/shared", () => ({ basePath: "" }));
vi.mock("@/lib/workspace-files", () => ({
  BrowserBlobSource: class BrowserBlobSource {},
  BrowserOpfsFile: class BrowserOpfsFile {},
  browserScratchFactory: {},
  createBrowserExportName: vi.fn(),
  removeOpfsFile: vi.fn(),
  withBrowserExportLease: vi.fn(),
}));
vi.mock("@/lib/workspace-source", () => ({
  workspaceSourceDescription: (value: string) => value,
  workspaceSourceFileName: (value: string) => value,
  workspaceSourceKey: (value: string) => value,
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
  readonly plans?: readonly unknown[];
}

function inspection(planId: string) {
  return {
    application: { state: "pending" },
    captureIds: [],
    capturedRows: 0n,
    conflicts: [],
    examples: [],
    nextRouteCursor: undefined,
    prepared: {
      id: planId,
      planRevision: 1n,
      reviewFingerprint: `fingerprint-${planId}`,
      state: "ready",
    },
    previewWarnings: [],
    reviewRows: 0n,
    routeCount: 0n,
    routes: [],
    targetRevision: 0n,
  };
}

async function workerHarness() {
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
  return {
    messages,
    request(command: Record<string, unknown>) {
      return new Promise<WorkerMessage>((resolve) => {
        waiters.set(command["id"] as number, resolve);
        receive({ data: command } as MessageEvent<Record<string, unknown>>);
      });
    },
  };
}

describe("workspace worker saved-plan cleanup", () => {
  beforeEach(() => {
    vi.resetModules();
    vi.clearAllMocks();
    mocks.firstClose
      .mockRejectedValueOnce(new Error("Injected first plan close failure"))
      .mockResolvedValue(undefined);
    mocks.secondClose
      .mockRejectedValueOnce(new Error("Injected second plan close failure"))
      .mockResolvedValue(undefined);
    mocks.openImportBatch.mockImplementation(async ({ name }) =>
      name === "first.ccplan"
        ? { id: "first-plan", close: mocks.firstClose }
        : { id: "second-plan", close: mocks.secondClose },
    );
  });

  it("withholds a successful listing and retries every retained plan close", async () => {
    mocks.listImportBatches
      .mockResolvedValueOnce({
        imports: [
          {
            name: "first.ccplan",
            databaseId: "database-1",
            application: "applied",
          },
          {
            name: "second.ccplan",
            databaseId: "database-1",
            application: "pending",
          },
        ],
        ignored: [],
      })
      .mockResolvedValueOnce({ imports: [], ignored: [] });
    mocks.inspectImport.mockImplementation(async ({ prepared }) => {
      if (prepared.id === "first-plan") return inspection("first-plan");
      throw new Error("Injected second plan inspection failure");
    });
    const { messages, request } = await workerHarness();

    await expect(
      request({ id: 1, type: "create", name: "test.sqlite", format: "sqlite" }),
    ).resolves.toMatchObject({ type: "ready" });
    await expect(
      request({ id: 2, type: "listImports" }),
    ).resolves.toMatchObject({
      type: "error",
      code: "DB_BROWSER_IMPORT_CLEANUP_REQUIRED",
    });
    expect(mocks.firstClose).toHaveBeenCalledOnce();
    expect(mocks.secondClose).toHaveBeenCalledOnce();
    expect(
      messages.filter(
        (message) => message.id === 2 && message.type === "importsListed",
      ),
    ).toEqual([]);

    await expect(
      request({ id: 3, type: "listImports" }),
    ).resolves.toMatchObject({ type: "importsListed", plans: [] });
    expect(mocks.firstClose).toHaveBeenCalledTimes(2);
    expect(mocks.secondClose).toHaveBeenCalledTimes(2);
    expect(mocks.firstClose.mock.invocationCallOrder[1]).toBeLessThan(
      mocks.listImportBatches.mock.invocationCallOrder[1]!,
    );
    expect(mocks.secondClose.mock.invocationCallOrder[1]).toBeLessThan(
      mocks.listImportBatches.mock.invocationCallOrder[1]!,
    );
  });

  it("drains a retained saved-plan handle during worker close", async () => {
    mocks.listImportBatches.mockResolvedValue({
      imports: [
        {
          name: "second.ccplan",
          databaseId: "database-1",
          application: "pending",
        },
      ],
      ignored: [],
    });
    mocks.inspectImport.mockRejectedValue(
      new Error("Injected plan inspection failure"),
    );
    const { request } = await workerHarness();

    await request({
      id: 1,
      type: "create",
      name: "test.sqlite",
      format: "sqlite",
    });
    await expect(
      request({ id: 2, type: "listImports" }),
    ).resolves.toMatchObject({
      type: "error",
      code: "DB_BROWSER_IMPORT_CLEANUP_REQUIRED",
    });
    await expect(request({ id: 3, type: "close" })).resolves.toMatchObject({
      type: "closed",
    });
    expect(mocks.secondClose).toHaveBeenCalledTimes(2);
    expect(mocks.databaseClose).toHaveBeenCalledOnce();
  });

  it("loads each route page from the saved batch instead of cached worker state", async () => {
    mocks.firstClose.mockReset().mockResolvedValue(undefined);
    mocks.listImportBatches.mockResolvedValue({
      imports: [
        {
          name: "first.ccplan",
          databaseId: "database-1",
          application: "pending",
        },
      ],
      ignored: [],
    });
    mocks.inspectImport.mockImplementation(async ({ routePage }) => {
      const next = routePage?.cursor === "route-page-2";
      return {
        ...inspection("first-plan"),
        nextRouteCursor: next ? undefined : "route-page-2",
        reviewRows: 51n,
        routeCount: 51n,
        routes: [
          {
            applicationState: "not-applied",
            captureId: next ? "capture-51" : "capture-1",
            columns: [],
            destination: null,
            destinationColumns: [],
            displayName: "source.xlsx",
            inferredColumns: [{ name: "Value", type: "text" }],
            label: next ? "Sheet 51" : "Sheet 1",
            rowCount: 1n,
            source: "source-1",
            selection: next ? "selection-51" : "selection-1",
            suggestedDestination: {
              kind: "new-table-infer",
              name: next ? "Sheet_51" : "Sheet_1",
              recordId: { prefix: "SHEET", padding: 6 },
            },
          },
        ],
      };
    });
    const { request } = await workerHarness();

    await request({
      id: 1,
      type: "create",
      name: "test.sqlite",
      format: "sqlite",
    });
    await expect(
      request({ id: 2, type: "listImports" }),
    ).resolves.toMatchObject({
      type: "importsListed",
      plans: [
        {
          routeCursor: null,
          nextRouteCursor: "route-page-2",
          regions: [{ label: "Sheet 1" }],
        },
      ],
    });
    await expect(
      request({
        id: 3,
        type: "inspectImport",
        planId: "first-plan",
        routeCursor: "route-page-2",
      }),
    ).resolves.toMatchObject({
      type: "importInspected",
      plan: {
        routeCursor: "route-page-2",
        nextRouteCursor: null,
        routeCount: 51,
        regions: [{ label: "Sheet 51" }],
      },
    });
    expect(mocks.inspectImport).toHaveBeenLastCalledWith(
      expect.objectContaining({
        routePage: { limit: 50, cursor: "route-page-2" },
      }),
    );
  });
});
