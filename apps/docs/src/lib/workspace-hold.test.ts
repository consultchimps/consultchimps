import { describe, expect, it } from "vitest";

import { mustHoldWorkspace, type WorkspaceHoldState } from "./workspace-hold";

const nothingAtStake: WorkspaceHoldState = {
  unsavedChanges: false,
  importing: false,
  editsInFlight: 0,
};

describe("mustHoldWorkspace", () => {
  it("lets a clean workspace go without a word", () => {
    expect(mustHoldWorkspace(nothingAtStake)).toBe(false);
  });

  it("holds for changes no file has", () => {
    expect(mustHoldWorkspace({ ...nothingAtStake, unsavedChanges: true })).toBe(
      true,
    );
  });

  it("holds for an import that has not come back", () => {
    expect(mustHoldWorkspace({ ...nothingAtStake, importing: true })).toBe(
      true,
    );
  });

  it("holds for a cell edit that has not come back", () => {
    // The window this closes: the workspace still reads as clean, because an
    // edit is counted as unsaved only once the worker accepts it, and a visitor
    // who commits an edit and reaches straight for New would otherwise replace
    // the database the edit is still travelling to.
    expect(mustHoldWorkspace({ ...nothingAtStake, editsInFlight: 1 })).toBe(
      true,
    );
    expect(mustHoldWorkspace({ ...nothingAtStake, editsInFlight: 4 })).toBe(
      true,
    );
  });

  it("stops holding once the last edit has settled", () => {
    // A refused edit changes nothing, so there is nothing left to ask about.
    expect(mustHoldWorkspace({ ...nothingAtStake, editsInFlight: 0 })).toBe(
      false,
    );
  });

  it("holds while any one reason stands, whatever the others say", () => {
    const reasons: ReadonlyArray<keyof WorkspaceHoldState> = [
      "unsavedChanges",
      "importing",
      "editsInFlight",
    ];
    for (const reason of reasons) {
      const state: WorkspaceHoldState = {
        ...nothingAtStake,
        [reason]: reason === "editsInFlight" ? 1 : true,
      };
      expect(mustHoldWorkspace(state)).toBe(true);
    }
    expect(
      mustHoldWorkspace({
        unsavedChanges: true,
        importing: true,
        editsInFlight: 2,
      }),
    ).toBe(true);
  });
});
