import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  listDatabases: vi.fn(),
  removeDatabase: vi.fn(),
}));

vi.mock("@consultchimps/db/browser", () => ({
  configureBrowserDatabaseRuntime: vi.fn(async () => ({
    listDatabases: mocks.listDatabases,
    removeDatabase: mocks.removeDatabase,
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
  readonly message?: string;
  readonly listing?: unknown;
}

async function workerHarness() {
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
  return {
    request(command: Record<string, unknown>) {
      return new Promise<WorkerMessage>((resolve) => {
        waiters.set(command["id"] as number, resolve);
        receive({ data: command } as MessageEvent<Record<string, unknown>>);
      });
    },
  };
}

describe("workspace worker working copies", () => {
  beforeEach(() => {
    vi.resetModules();
    mocks.listDatabases.mockReset();
    mocks.removeDatabase.mockReset();
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("returns the stored listing without an open database", async () => {
    const listing = {
      databases: [{ name: "alpha.sqlite", format: "sqlite" }],
      ignored: [],
    };
    mocks.listDatabases.mockResolvedValue(listing);
    const { request } = await workerHarness();
    await expect(request({ id: 1, type: "listDatabases" })).resolves.toEqual({
      type: "databases",
      id: 1,
      listing,
    });
  });

  it("removes a copy by name and reports a coded refusal unchanged", async () => {
    mocks.removeDatabase
      .mockResolvedValueOnce(undefined)
      .mockImplementationOnce(async () => {
        const core = await import("@consultchimps/core");
        throw new core.ConsultChimpsError(
          "DB_BROWSER_DATABASE_BUSY",
          "Close the open browser working copy before replacing it.",
        );
      });
    const { request } = await workerHarness();
    await expect(
      request({ id: 1, type: "removeDatabase", name: "alpha.sqlite" }),
    ).resolves.toEqual({ type: "databaseRemoved", id: 1 });
    expect(mocks.removeDatabase).toHaveBeenCalledWith({
      name: "alpha.sqlite",
    });
    await expect(
      request({ id: 2, type: "removeDatabase", name: "open.sqlite" }),
    ).resolves.toMatchObject({
      type: "error",
      code: "DB_BROWSER_DATABASE_BUSY",
    });
  });
});
