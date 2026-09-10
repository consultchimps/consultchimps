import { isConsultChimpsError } from "@consultchimps/core";
import { describe, expect, it } from "vitest";

import { HeldWorkspace } from "./workspace-generation";
import {
  WORKSPACE_STALE_EDIT,
  WORKSPACE_STALE_READ,
} from "./workspace-protocol";

/** A workspace that only records whether it was closed, which is all this owns. */
function fake(): { closed: boolean; close: () => void } {
  const workspace = {
    closed: false,
    close: (): void => {
      workspace.closed = true;
    },
  };
  return workspace;
}

function codeOf(run: () => unknown): string {
  try {
    run();
  } catch (error) {
    return isConsultChimpsError(error) ? error.code : "NOT_A_STABLE_ERROR";
  }
  return "NO_ERROR_THROWN";
}

describe("HeldWorkspace", () => {
  it("holds nothing until a workspace is opened", () => {
    const held = new HeldWorkspace<ReturnType<typeof fake>>();

    expect(held.generation).toBe(0);
    expect(codeOf(() => held.require())).toBe("WORKSPACE_NONE_OPEN");
    // Nothing has ever handed out a generation to quote back, so nothing but
    // the sentinel can match.
    expect(codeOf(() => held.assertCurrent(0, "edit"))).toBe("NO_ERROR_THROWN");
    expect(codeOf(() => held.assertCurrent(1, "edit"))).toBe(
      WORKSPACE_STALE_EDIT,
    );
  });

  it("accepts a command that names the workspace now held", () => {
    const held = new HeldWorkspace<ReturnType<typeof fake>>();
    const first = fake();
    held.replace(first);

    expect(held.generation).toBe(1);
    expect(held.require()).toBe(first);
    expect(codeOf(() => held.assertCurrent(1, "edit"))).toBe("NO_ERROR_THROWN");
    expect(codeOf(() => held.assertCurrent(1, "read"))).toBe("NO_ERROR_THROWN");
  });

  it("closes the previous workspace when another replaces it", () => {
    const held = new HeldWorkspace<ReturnType<typeof fake>>();
    const first = fake();
    const second = fake();

    held.replace(first);
    held.replace(second);

    expect(first.closed).toBe(true);
    expect(second.closed).toBe(false);
    expect(held.require()).toBe(second);
  });

  it("refuses a command made against a workspace that has been replaced", () => {
    const held = new HeldWorkspace<ReturnType<typeof fake>>();
    held.replace(fake());
    held.replace(fake());

    // The case the counter exists for: a command posted before the open and
    // delivered after it would otherwise act on the new database, where the
    // same Record ID is very likely a different record.
    expect(codeOf(() => held.assertCurrent(1, "edit"))).toBe(
      WORKSPACE_STALE_EDIT,
    );
    expect(codeOf(() => held.assertCurrent(1, "read"))).toBe(
      WORKSPACE_STALE_READ,
    );
  });

  it("refuses a command for a workspace that has been released", () => {
    const held = new HeldWorkspace<ReturnType<typeof fake>>();
    const only = fake();
    held.replace(only);
    held.release();

    expect(only.closed).toBe(true);
    expect(codeOf(() => held.require())).toBe("WORKSPACE_NONE_OPEN");
    expect(codeOf(() => held.assertCurrent(1, "edit"))).toBe(
      WORKSPACE_STALE_EDIT,
    );
  });

  it("refuses a command made before the workspace itself changed", () => {
    const held = new HeldWorkspace<ReturnType<typeof fake>>();
    const only = fake();
    held.replace(only);
    const beforeImport = held.generation;

    // An import adds tables and rows to the database already open rather than
    // replacing it, so the object is the same one and only the generation says
    // that a snapshot taken before it knows neither the new tables nor the new
    // rows.
    held.changed();

    expect(held.require()).toBe(only);
    expect(only.closed).toBe(false);
    expect(held.generation).not.toBe(beforeImport);
    expect(codeOf(() => held.assertCurrent(beforeImport, "edit"))).toBe(
      WORKSPACE_STALE_EDIT,
    );
    expect(codeOf(() => held.assertCurrent(held.generation, "edit"))).toBe(
      "NO_ERROR_THROWN",
    );
  });

  it("moves the generation on every way the workspace can change", () => {
    const held = new HeldWorkspace<ReturnType<typeof fake>>();
    const seen = new Set<number>([held.generation]);

    // There is no way to reach the held workspace that leaves the generation
    // where it was, which is what stops a handler forgetting to move it.
    held.replace(fake());
    seen.add(held.generation);
    held.changed();
    seen.add(held.generation);
    held.replace(fake());
    seen.add(held.generation);
    held.release();
    seen.add(held.generation);

    expect(seen.size).toBe(5);
  });

  it("says what happened to the edit, without quoting bookkeeping numbers", () => {
    const held = new HeldWorkspace<ReturnType<typeof fake>>();
    held.replace(fake());
    held.replace(fake());

    let message = "";
    try {
      held.assertCurrent(1, "edit");
    } catch (error) {
      message = error instanceof Error ? error.message : "";
    }
    expect(message).toContain("no longer open");
    expect(message).toContain("was not applied");
    expect(message).not.toMatch(/\d/u);
  });
});
