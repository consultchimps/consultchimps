import { describe, expect, it, vi } from "vitest";

import {
  closeTrackedResources,
  replaceActiveWorkspace,
  WorkspaceReplacementCleanupError,
} from "./workspace-replacement";

function workspace(close: () => Promise<void> = async () => undefined) {
  return { close };
}

describe("workspace replacement", () => {
  it("activates the next workspace only after dependencies and the previous workspace close", async () => {
    const calls: string[] = [];
    const previous = workspace(async () => {
      calls.push("previous");
    });
    const nextClose = vi.fn(async () => undefined);
    const next = workspace(nextClose);

    await expect(
      replaceActiveWorkspace({
        previous,
        next,
        async prepare() {
          calls.push("prepare");
          return "ready";
        },
        async closeDependencies() {
          calls.push("dependencies");
        },
        close: (value) => value.close(),
        activate(received) {
          expect(received).toBe(next);
          calls.push("activate");
        },
      }),
    ).resolves.toBe("ready");

    expect(calls).toEqual(["prepare", "dependencies", "previous", "activate"]);
    expect(nextClose).not.toHaveBeenCalled();
  });

  it("keeps the previous workspace active when its close fails", async () => {
    const closeFailure = new Error("Injected previous close failure");
    const previous = workspace(vi.fn(async () => Promise.reject(closeFailure)));
    const nextClose = vi.fn(async () => undefined);
    const next = workspace(nextClose);
    let active = previous;

    await expect(
      replaceActiveWorkspace({
        previous,
        next,
        async prepare() {
          return "ready";
        },
        async closeDependencies() {},
        close: (value) => value.close(),
        activate(received) {
          active = received;
        },
      }),
    ).rejects.toBe(closeFailure);
    expect(active).toBe(previous);
    expect(nextClose).toHaveBeenCalledOnce();
  });

  it("closes the next workspace when dependency cleanup fails", async () => {
    const dependencyFailure = new Error("Injected import close failure");
    const previousClose = vi.fn(async () => undefined);
    const previous = workspace(previousClose);
    const nextClose = vi.fn(async () => undefined);
    const next = workspace(nextClose);
    const activate = vi.fn();

    await expect(
      replaceActiveWorkspace({
        previous,
        next,
        async prepare() {
          return "ready";
        },
        async closeDependencies() {
          throw dependencyFailure;
        },
        close: (value) => value.close(),
        activate,
      }),
    ).rejects.toBe(dependencyFailure);
    expect(previousClose).not.toHaveBeenCalled();
    expect(nextClose).toHaveBeenCalledOnce();
    expect(activate).not.toHaveBeenCalled();
  });

  it("preserves both the replacement and next cleanup failures", async () => {
    const replacementFailure = new Error("Injected import close failure");
    const nextCleanupFailure = new Error("Injected next close failure");
    const next = workspace(
      vi.fn(async () => Promise.reject(nextCleanupFailure)),
    );

    let rejection: unknown;
    try {
      await replaceActiveWorkspace({
        previous: workspace(),
        next,
        async prepare() {
          return "ready";
        },
        async closeDependencies() {
          throw replacementFailure;
        },
        close: (value) => value.close(),
        activate: vi.fn(),
      });
    } catch (error) {
      rejection = error;
    }
    expect(rejection).toBeInstanceOf(WorkspaceReplacementCleanupError);
    expect(rejection).toMatchObject({
      replacementCause: replacementFailure,
      nextCleanupCause: nextCleanupFailure,
    });
    expect((rejection as Error).cause).toBeInstanceOf(AggregateError);
  });

  it("closes the next workspace when preparing its ready state fails", async () => {
    const summaryFailure = new Error("Injected summary failure");
    const previousClose = vi.fn(async () => undefined);
    const nextClose = vi.fn(async () => undefined);
    const activate = vi.fn();

    await expect(
      replaceActiveWorkspace({
        previous: workspace(previousClose),
        next: workspace(nextClose),
        async prepare() {
          throw summaryFailure;
        },
        closeDependencies: vi.fn(async () => undefined),
        close: (value) => value.close(),
        activate,
      }),
    ).rejects.toBe(summaryFailure);
    expect(previousClose).not.toHaveBeenCalled();
    expect(nextClose).toHaveBeenCalledOnce();
    expect(activate).not.toHaveBeenCalled();
  });
});

describe("tracked resource cleanup", () => {
  it("removes successful closes and retains failed closes for retry", async () => {
    const successful = { name: "successful" };
    const failed = { name: "failed" };
    const resources = new Map([
      ["successful", successful],
      ["failed", failed],
    ]);
    const failure = new Error("Injected tracked close failure");

    await expect(
      closeTrackedResources({
        resources,
        async close(resource) {
          if (resource === failed) throw failure;
        },
      }),
    ).rejects.toBe(failure);
    expect(resources).toEqual(new Map([["failed", failed]]));
  });
});
