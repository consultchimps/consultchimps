import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  databaseClose: vi.fn(async () => undefined),
  inspectDatabase: vi.fn(),
  readTableRows: vi.fn(),
}));

vi.mock("@consultchimps/db", async (importOriginal) => ({
  ...(await importOriginal<typeof import("@consultchimps/db")>()),
  inspectDatabase: mocks.inspectDatabase,
  readTableRows: mocks.readTableRows,
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
  readonly page?: unknown;
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

describe("workspace worker row pages", () => {
  beforeEach(() => {
    vi.resetModules();
    mocks.readTableRows.mockReset();
    mocks.inspectDatabase.mockReset().mockResolvedValue({
      completedImports: 0n,
      recordedBatches: 0n,
      format: "sqlite",
      formatVersion: 1,
      id: "database-1",
      revision: 0n,
      tables: [],
    });
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("maps a page with its cursors and refuses without an open database", async () => {
    const { request } = await workerHarness();
    await expect(
      request({ id: 1, type: "readRows", table: "t", cursor: null, limit: 50 }),
    ).resolves.toMatchObject({
      type: "error",
      message: expect.stringContaining("Create or open a database"),
    });

    await request({
      id: 2,
      type: "create",
      name: "rows.sqlite",
      format: "sqlite",
    });
    mocks.readTableRows.mockResolvedValueOnce({
      table: "Readings",
      columns: [{ name: "record_id", type: "text" }],
      rows: [["READ-0001"]],
      nextCursor: '["readings","READ-0001"]',
    });
    await expect(
      request({
        id: 3,
        type: "readRows",
        table: "readings",
        cursor: null,
        limit: 50,
      }),
    ).resolves.toEqual({
      type: "rows",
      id: 3,
      page: {
        table: "Readings",
        columns: [{ name: "record_id", type: "text" }],
        rows: [["READ-0001"]],
        cursor: null,
        nextCursor: '["readings","READ-0001"]',
      },
    });
    expect(mocks.readTableRows).toHaveBeenCalledWith(
      expect.objectContaining({
        table: "readings",
        page: { limit: 50 },
      }),
    );

    mocks.readTableRows.mockResolvedValueOnce({
      table: "Readings",
      columns: [{ name: "record_id", type: "text" }],
      rows: [],
    });
    await expect(
      request({
        id: 4,
        type: "readRows",
        table: "readings",
        cursor: '["readings","READ-0001"]',
        limit: 50,
      }),
    ).resolves.toMatchObject({
      type: "rows",
      page: { cursor: '["readings","READ-0001"]', nextCursor: null },
    });
    expect(mocks.readTableRows).toHaveBeenLastCalledWith(
      expect.objectContaining({
        page: { limit: 50, cursor: '["readings","READ-0001"]' },
      }),
    );
  });

  it("passes a coded refusal through unchanged", async () => {
    const { request } = await workerHarness();
    await request({
      id: 1,
      type: "create",
      name: "rows.sqlite",
      format: "sqlite",
    });
    mocks.readTableRows.mockImplementationOnce(async () => {
      const core = await import("@consultchimps/core");
      throw new core.ConsultChimpsError(
        "DB_TABLE_NOT_FOUND",
        "The table does not exist in this database.",
      );
    });
    await expect(
      request({
        id: 2,
        type: "readRows",
        table: "gone",
        cursor: null,
        limit: 50,
      }),
    ).resolves.toMatchObject({ type: "error", code: "DB_TABLE_NOT_FOUND" });
  });
});
