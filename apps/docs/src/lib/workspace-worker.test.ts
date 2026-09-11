import { isConsultChimpsError } from "@consultchimps/core";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import {
  WORKSPACE_GESTURE_TOO_LARGE,
  type UpdateWorkspaceCellsCommand,
  type WorkspaceCommand,
  type WorkspaceEvent,
} from "./workspace-protocol";
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
    const queued = client.readTable("Customer", 1);
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

/**
 * Any worker event without its id, which the stub below fills in from the
 * command it is answering. Written as a mapped type rather than an `Omit` of the
 * union, because omitting from a union of object types collapses it to the
 * members they share.
 */
type ReplyWithoutId = {
  [Kind in WorkspaceEvent["type"]]: Omit<
    Extract<WorkspaceEvent, { type: Kind }>,
    "id"
  >;
}[WorkspaceEvent["type"]];

/**
 * A worker that records every command and answers the ones a test tells it to,
 * so the client's own behaviour can be checked without an engine: how many
 * commands one call posts, what it makes of a reply, and what it does with a
 * failure.
 */
class ScriptedWorker {
  static latest: ScriptedWorker | null = null;
  readonly posted: WorkspaceCommand[] = [];
  #listeners: Array<(event: MessageEvent<WorkspaceEvent>) => void> = [];

  constructor() {
    ScriptedWorker.latest = this;
  }

  addEventListener(
    type: string,
    listener: (event: MessageEvent<WorkspaceEvent>) => void,
  ): void {
    if (type === "message") {
      this.#listeners.push(listener);
    }
  }

  postMessage(command: WorkspaceCommand): void {
    this.posted.push(command);
  }

  terminate(): void {}

  /** Answer the command posted at `index` with this event. */
  reply(event: ReplyWithoutId, index = 0): void {
    const command = this.posted[index];
    if (command === undefined) {
      throw new Error(`No command was posted at ${String(index)}`);
    }
    const full = { ...event, id: command.id } as WorkspaceEvent;
    for (const listener of this.#listeners) {
      listener({ data: full } as MessageEvent<WorkspaceEvent>);
    }
  }
}

describe("WorkspaceClient.updateCells", () => {
  beforeEach(() => {
    ScriptedWorker.latest = null;
    vi.stubGlobal("Worker", ScriptedWorker);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  const gesture = {
    generation: 3,
    table: "Customer",
    writes: [
      { recordId: "CUST-0001", column: "name", value: "Acme" },
      { recordId: "CUST-0001", column: "headcount", value: "12" },
      { recordId: "CUST-0002", column: "name", value: "Globex" },
      { recordId: "CUST-0002", column: "headcount", value: null },
    ],
  };

  it("posts exactly one command for the whole gesture", async () => {
    const client = new WorkspaceClient();
    const pending = client.updateCells(gesture);
    await settle();
    const worker = ScriptedWorker.latest;

    // One command, carrying every cell and the generation it belongs to. A
    // command per cell is what this rule exists to rule out.
    expect(worker?.posted).toHaveLength(1);
    const posted = worker?.posted[0] as UpdateWorkspaceCellsCommand;
    expect(posted.type).toBe("updateCells");
    expect(posted.generation).toBe(3);
    expect(posted.table).toBe("Customer");
    expect(posted.writes).toEqual(gesture.writes);

    worker?.reply({ type: "cellsUpdated", results: [] });
    expect(await pending).toEqual([]);
  });

  it("reports each cell's outcome, rebuilding a refusal into a stable error", async () => {
    const client = new WorkspaceClient();
    const pending = client.updateCells(gesture);
    await settle();
    ScriptedWorker.latest?.reply({
      type: "cellsUpdated",
      results: [
        {
          accepted: true,
          recordId: "CUST-0001",
          column: "name",
          value: "Acme",
        },
        {
          accepted: true,
          recordId: "CUST-0001",
          column: "headcount",
          value: 12,
        },
        {
          accepted: false,
          recordId: "CUST-0002",
          column: "name",
          message: "That value is not allowed.",
          code: "DB_INVALID_NUMBER",
        },
        {
          accepted: false,
          recordId: "CUST-0002",
          column: "headcount",
          message: "Something went wrong.",
        },
      ],
    });

    const outcomes = await pending;
    expect(outcomes.map((outcome) => outcome.accepted)).toEqual([
      true,
      true,
      false,
      false,
    ]);
    expect(outcomes[1]).toEqual({
      accepted: true,
      recordId: "CUST-0001",
      column: "headcount",
      value: 12,
    });
    const refused = outcomes[2];
    expect(refused?.accepted).toBe(false);
    const error = refused?.accepted === false ? refused.error : null;
    expect(isConsultChimpsError(error)).toBe(true);
    expect(isConsultChimpsError(error) && error.code).toBe("DB_INVALID_NUMBER");
    // A failure with no code is still an error, just not a stable one.
    const plain = outcomes[3];
    const withoutCode = plain?.accepted === false ? plain.error : null;
    expect(withoutCode).toBeInstanceOf(Error);
    expect(isConsultChimpsError(withoutCode)).toBe(false);
  });

  it("rejects the whole gesture when the worker refuses it", async () => {
    const client = new WorkspaceClient();
    const pending = client.updateCells(gesture);
    await settle();
    ScriptedWorker.latest?.reply({
      type: "error",
      message: "That change covers too many cells.",
      code: WORKSPACE_GESTURE_TOO_LARGE,
    });

    expect(await codeOf(pending)).toBe(WORKSPACE_GESTURE_TOO_LARGE);
  });

  it("waits for a cell edit already in flight before posting the gesture", async () => {
    const client = new WorkspaceClient();
    const edit = client.updateCell({
      generation: 3,
      table: "Customer",
      recordId: "CUST-0001",
      column: "name",
      value: "Acme",
    });
    const pending = client.updateCells(gesture);
    await settle();
    const worker = ScriptedWorker.latest;

    // The worker holds one database and runs one command at a time, so the
    // gesture is not even posted until the edit before it has been answered.
    expect(worker?.posted).toHaveLength(1);
    expect(worker?.posted[0]?.type).toBe("updateCell");

    worker?.reply({ type: "cellUpdated", value: "Acme" });
    await edit;
    await settle();

    expect(worker?.posted).toHaveLength(2);
    expect(worker?.posted[1]?.type).toBe("updateCells");
    worker?.reply({ type: "cellsUpdated", results: [] }, 1);
    expect(await pending).toEqual([]);
  });
});
