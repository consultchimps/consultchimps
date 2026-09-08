import { isConsultChimpsError } from "@consultchimps/core";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import {
  WORKSPACE_WORKER_UNAVAILABLE,
  WorkspaceClient,
} from "./workspace-worker";

/**
 * A worker that accepts commands and never answers, so every command stays
 * pending and the client's queue keeps commands waiting behind it. That is the
 * state a page is in when it unmounts with edits outstanding.
 */
class SilentWorker {
  static created: SilentWorker[] = [];
  terminated = false;

  constructor() {
    SilentWorker.created.push(this);
  }

  addEventListener(): void {}

  postMessage(): void {}

  terminate(): void {
    this.terminated = true;
  }
}

/**
 * The client posts a command from a microtask (its queue is a promise chain),
 * so a worker exists only after the current task yields. A macrotask hop lets
 * every queued microtask run first.
 */
function settle(): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, 0);
  });
}

async function codeOf(run: Promise<unknown>): Promise<string> {
  try {
    await run;
  } catch (error) {
    return isConsultChimpsError(error) ? error.code : "NOT_A_STABLE_ERROR";
  }
  return "NO_ERROR_THROWN";
}

describe("WorkspaceClient.terminate", () => {
  beforeEach(() => {
    SilentWorker.created = [];
    vi.stubGlobal("Worker", SilentWorker);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("fails queued commands without starting a worker again", async () => {
    const client = new WorkspaceClient();
    const pending = client.create();
    const queued = client.listTables();
    await settle();
    expect(SilentWorker.created).toHaveLength(1);

    client.terminate();

    // The pending command is rejected outright; the queued one reaches the
    // client only afterwards, and must be refused rather than served by a
    // fresh worker the page has already left behind.
    expect(await codeOf(pending)).toBe(WORKSPACE_WORKER_UNAVAILABLE);
    expect(await codeOf(queued)).toBe(WORKSPACE_WORKER_UNAVAILABLE);
    expect(SilentWorker.created).toHaveLength(1);
    expect(SilentWorker.created[0]?.terminated).toBe(true);
  });

  it("refuses a command issued after termination", async () => {
    const client = new WorkspaceClient();
    client.terminate();
    expect(await codeOf(client.create())).toBe(WORKSPACE_WORKER_UNAVAILABLE);
    expect(SilentWorker.created).toHaveLength(0);
  });

  it("can be called more than once", async () => {
    const client = new WorkspaceClient();
    void client.create().catch(() => undefined);
    await settle();
    client.terminate();
    expect(() => client.terminate()).not.toThrow();
    expect(SilentWorker.created).toHaveLength(1);
  });
});
