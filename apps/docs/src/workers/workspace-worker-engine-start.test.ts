import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  configureBrowserDatabaseRuntime: vi.fn(),
  databaseClose: vi.fn(async () => undefined),
  inspectDatabase: vi.fn(),
}));

vi.mock("@consultchimps/db", async (importOriginal) => ({
  ...(await importOriginal<typeof import("@consultchimps/db")>()),
  inspectDatabase: mocks.inspectDatabase,
}));

vi.mock("@consultchimps/db/browser", () => ({
  configureBrowserDatabaseRuntime: mocks.configureBrowserDatabaseRuntime,
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
  readonly progress?: { readonly message: string };
}

function runtimeStub() {
  return {
    createDatabase: vi.fn(async () => ({
      database: {
        close: mocks.databaseClose,
        format: "sqlite",
        id: "database-1",
      },
    })),
  };
}

function busyHandles(): DOMException {
  return new DOMException(
    "Access Handles cannot be created if there is another open Access Handle",
    "NoModificationAllowedError",
  );
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

const CREATE = { type: "create", name: "retry.sqlite", format: "sqlite" };

describe("workspace worker engine start", () => {
  beforeEach(() => {
    vi.resetModules();
    vi.clearAllMocks();
    mocks.configureBrowserDatabaseRuntime.mockReset();
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
    vi.useRealTimers();
    vi.unstubAllGlobals();
  });

  it("does not report a wait when the cancellation lands during the attempt", async () => {
    let failAttempt!: (error: unknown) => void;
    mocks.configureBrowserDatabaseRuntime.mockImplementationOnce(
      () =>
        new Promise((_resolve, reject) => {
          failAttempt = reject;
        }),
    );
    const { messages, request } = await workerHarness();

    const pending = request({ id: 1, ...CREATE });
    await request({ id: 2, type: "cancel", targetId: 1 });
    failAttempt(busyHandles());
    await expect(pending).resolves.toMatchObject({
      type: "error",
      code: "OPERATION_ABORTED",
    });
    expect(
      messages.filter((message) => message.type === "progress"),
    ).toHaveLength(0);

    mocks.configureBrowserDatabaseRuntime.mockResolvedValue(runtimeStub());
    await expect(request({ id: 3, ...CREATE })).resolves.toMatchObject({
      type: "ready",
    });
  });

  it("reports a raw start failure with a stable code and starts again on the next command", async () => {
    mocks.configureBrowserDatabaseRuntime
      .mockRejectedValueOnce(new Error("Aborted(NetworkError: fetch failed)"))
      .mockResolvedValue(runtimeStub());
    const { request } = await workerHarness();

    const failed = await request({ id: 1, ...CREATE });
    expect(failed).toMatchObject({
      type: "error",
      code: "DB_BROWSER_ENGINE_UNAVAILABLE",
    });
    expect(failed.message).toContain("could not start");
    expect(failed.message).toContain("NetworkError");

    const recovered = await request({ id: 2, ...CREATE });
    expect(recovered).toMatchObject({ type: "ready" });
    expect(mocks.configureBrowserDatabaseRuntime).toHaveBeenCalledTimes(2);
  });

  it("waits for busy file handles with progress and then starts", async () => {
    vi.useFakeTimers();
    mocks.configureBrowserDatabaseRuntime
      .mockRejectedValueOnce(busyHandles())
      .mockRejectedValueOnce(busyHandles())
      .mockResolvedValue(runtimeStub());
    const { messages, request } = await workerHarness();

    const pending = request({ id: 1, ...CREATE });
    await vi.advanceTimersByTimeAsync(2_500);
    await expect(pending).resolves.toMatchObject({ type: "ready" });
    expect(mocks.configureBrowserDatabaseRuntime).toHaveBeenCalledTimes(3);
    const waits = messages.filter((message) => message.type === "progress");
    expect(waits).toHaveLength(2);
    expect(waits[0]?.progress?.message).toContain("Waiting for another tab");
  });

  it("gives up on busy file handles after the wait limit with a busy message", async () => {
    vi.useFakeTimers();
    mocks.configureBrowserDatabaseRuntime.mockImplementation(async () => {
      throw busyHandles();
    });
    const { request } = await workerHarness();

    const pending = request({ id: 1, ...CREATE });
    await vi.advanceTimersByTimeAsync(60_000);
    const failed = await pending;
    expect(failed).toMatchObject({
      type: "error",
      code: "DB_BROWSER_ENGINE_UNAVAILABLE",
    });
    expect(failed.message).toContain("still holds the working database files");

    mocks.configureBrowserDatabaseRuntime.mockResolvedValue(runtimeStub());
    await expect(request({ id: 2, ...CREATE })).resolves.toMatchObject({
      type: "ready",
    });
  });

  it("stops waiting when the command is cancelled", async () => {
    vi.useFakeTimers();
    mocks.configureBrowserDatabaseRuntime.mockImplementation(async () => {
      throw busyHandles();
    });
    const { request } = await workerHarness();

    const pending = request({ id: 1, ...CREATE });
    await vi.advanceTimersByTimeAsync(500);
    await request({ id: 2, type: "cancel", targetId: 1 });
    await expect(pending).resolves.toMatchObject({
      type: "error",
      code: "OPERATION_ABORTED",
    });

    mocks.configureBrowserDatabaseRuntime.mockResolvedValue(runtimeStub());
    await expect(request({ id: 3, ...CREATE })).resolves.toMatchObject({
      type: "ready",
    });
  });

  it("keeps a coded start failure unchanged and still retries", async () => {
    mocks.configureBrowserDatabaseRuntime
      .mockImplementationOnce(async () => {
        // Build the error from the module registry the worker imported after
        // vi.resetModules, so the worker's instanceof brand check sees it.
        const core = await import("@consultchimps/core");
        throw new core.ConsultChimpsError(
          "DB_BROWSER_STORAGE_UNAVAILABLE",
          "This browser does not provide origin-private file storage.",
        );
      })
      .mockResolvedValue(runtimeStub());
    const { request } = await workerHarness();

    await expect(request({ id: 1, ...CREATE })).resolves.toMatchObject({
      type: "error",
      code: "DB_BROWSER_STORAGE_UNAVAILABLE",
      message: "This browser does not provide origin-private file storage.",
    });
    await expect(request({ id: 2, ...CREATE })).resolves.toMatchObject({
      type: "ready",
    });
  });

  it("starts the engine once for successive commands", async () => {
    mocks.configureBrowserDatabaseRuntime.mockResolvedValue(runtimeStub());
    const { request } = await workerHarness();

    await request({ id: 1, ...CREATE });
    await request({ id: 2, ...CREATE });
    expect(mocks.configureBrowserDatabaseRuntime).toHaveBeenCalledOnce();
  });
});
