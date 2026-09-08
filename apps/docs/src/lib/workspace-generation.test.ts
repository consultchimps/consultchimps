import { isConsultChimpsError } from "@consultchimps/core";
import { describe, expect, it } from "vitest";

import { WorkspaceGenerations } from "./workspace-generation";
import {
  WORKSPACE_STALE_EDIT,
  WORKSPACE_STALE_READ,
} from "./workspace-protocol";

function codeOf(run: () => unknown): string {
  try {
    run();
  } catch (error) {
    return isConsultChimpsError(error) ? error.code : "NOT_A_STABLE_ERROR";
  }
  return "NO_ERROR_THROWN";
}

describe("WorkspaceGenerations", () => {
  it("names no workspace before one is opened", () => {
    const generations = new WorkspaceGenerations();
    expect(generations.current).toBe(0);
    // A command that arrives before any workspace exists cannot match, because
    // no reply has ever handed out a generation to quote back.
    expect(codeOf(() => generations.assertCurrent(0, "edit"))).toBe(
      "NO_ERROR_THROWN",
    );
    expect(codeOf(() => generations.assertCurrent(1, "edit"))).toBe(
      WORKSPACE_STALE_EDIT,
    );
  });

  it("accepts a command that names the workspace now held", () => {
    const generations = new WorkspaceGenerations();
    const opened = generations.replaced();

    expect(opened).toBe(1);
    expect(generations.current).toBe(1);
    expect(codeOf(() => generations.assertCurrent(opened, "edit"))).toBe(
      "NO_ERROR_THROWN",
    );
    expect(codeOf(() => generations.assertCurrent(opened, "read"))).toBe(
      "NO_ERROR_THROWN",
    );
  });

  it("refuses an edit made against a workspace that has been replaced", () => {
    const generations = new WorkspaceGenerations();
    const first = generations.replaced();
    generations.replaced();

    // This is the case the counter exists for: an edit posted before the open
    // and delivered after it would otherwise be applied to the new database,
    // where the same Record ID is very likely a different record.
    expect(codeOf(() => generations.assertCurrent(first, "edit"))).toBe(
      WORKSPACE_STALE_EDIT,
    );
    expect(codeOf(() => generations.assertCurrent(first, "read"))).toBe(
      WORKSPACE_STALE_READ,
    );
  });

  it("refuses a command for a workspace that has been closed", () => {
    const generations = new WorkspaceGenerations();
    const opened = generations.replaced();
    // Closing releases the database, so what the grid holds is stale too.
    generations.replaced();

    expect(codeOf(() => generations.assertCurrent(opened, "edit"))).toBe(
      WORKSPACE_STALE_EDIT,
    );
  });

  it("says what happened to the edit, without quoting bookkeeping numbers", () => {
    const generations = new WorkspaceGenerations();
    generations.replaced();
    generations.replaced();

    let message = "";
    try {
      generations.assertCurrent(1, "edit");
    } catch (error) {
      message = error instanceof Error ? error.message : "";
    }
    expect(message).toContain("no longer open");
    expect(message).toContain("was not applied");
    expect(message).not.toMatch(/\d/u);
  });
});
