import { describe, expect, it } from "vitest";

import {
  mustHoldWorkspace,
  workspaceHoldHeading,
  workspaceHoldReasons,
  workspaceHoldSentence,
  type WorkspaceHoldState,
} from "./workspace-hold";

const nothingAtStake: WorkspaceHoldState = {
  unsavedChanges: false,
  importing: false,
  editsInFlight: 0,
  editorOpen: false,
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

  it("holds for a cell that is open for editing", () => {
    // What is typed and not committed exists only in an input element, so
    // nothing else knows there is anything to lose. This is held from the
    // moment the editor opens, which cannot miss a keystroke.
    expect(mustHoldWorkspace({ ...nothingAtStake, editorOpen: true })).toBe(
      true,
    );
  });

  it("holds while any one reason stands, whatever the others say", () => {
    const reasons: ReadonlyArray<keyof WorkspaceHoldState> = [
      "unsavedChanges",
      "importing",
      "editsInFlight",
      "editorOpen",
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
        editorOpen: true,
      }),
    ).toBe(true);
  });
});

/** Every state the hold can be in, named by the reasons that make it up. */
const CASES: ReadonlyArray<
  readonly [string, WorkspaceHoldState, readonly string[], string]
> = [
  ["nothing at stake", nothingAtStake, [], ""],
  [
    "a cell open for editing alone",
    { ...nothingAtStake, editorOpen: true },
    ["editorOpen"],
    "A cell is still open for editing.",
  ],
  [
    "a cell open for editing over unsaved changes",
    { ...nothingAtStake, unsavedChanges: true, editorOpen: true },
    ["unsavedChanges", "editorOpen"],
    "This workspace has changes that have not been saved to a file and a cell is still open for editing.",
  ],
  [
    "a cell open for editing while another edit is in flight",
    { ...nothingAtStake, editsInFlight: 1, editorOpen: true },
    ["editInFlight", "editorOpen"],
    "An edit is still being applied and a cell is still open for editing.",
  ],
  [
    "all four at once",
    {
      unsavedChanges: true,
      importing: true,
      editsInFlight: 3,
      editorOpen: true,
    },
    ["unsavedChanges", "importing", "editInFlight", "editorOpen"],
    "This workspace has changes that have not been saved to a file, an import is still running, an edit is still being applied, and a cell is still open for editing.",
  ],
  [
    "unsaved changes alone",
    { ...nothingAtStake, unsavedChanges: true },
    ["unsavedChanges"],
    "This workspace has changes that have not been saved to a file.",
  ],
  [
    "an import alone",
    { ...nothingAtStake, importing: true },
    ["importing"],
    "An import is still running.",
  ],
  [
    "an edit alone",
    { ...nothingAtStake, editsInFlight: 1 },
    ["editInFlight"],
    "An edit is still being applied.",
  ],
  [
    "unsaved changes and an import",
    { ...nothingAtStake, unsavedChanges: true, importing: true },
    ["unsavedChanges", "importing"],
    "This workspace has changes that have not been saved to a file and an import is still running.",
  ],
  [
    "unsaved changes and an edit",
    { ...nothingAtStake, unsavedChanges: true, editsInFlight: 2 },
    ["unsavedChanges", "editInFlight"],
    "This workspace has changes that have not been saved to a file and an edit is still being applied.",
  ],
  [
    "an import and an edit",
    { ...nothingAtStake, importing: true, editsInFlight: 1 },
    ["importing", "editInFlight"],
    "An import is still running and an edit is still being applied.",
  ],
  [
    "all three of the sent kinds at once",
    {
      ...nothingAtStake,
      unsavedChanges: true,
      importing: true,
      editsInFlight: 3,
    },
    ["unsavedChanges", "importing", "editInFlight"],
    "This workspace has changes that have not been saved to a file, an import is still running, and an edit is still being applied.",
  ],
];

describe("workspaceHoldReasons", () => {
  it.each(CASES)("names what is at stake with %s", (_, state, reasons) => {
    expect(workspaceHoldReasons(state)).toEqual(reasons);
  });

  it("agrees with the decision, because the decision is made from it", () => {
    for (const [, state] of CASES) {
      expect(mustHoldWorkspace(state)).toBe(
        workspaceHoldReasons(state).length > 0,
      );
    }
  });
});

describe("workspaceHoldSentence", () => {
  it.each(CASES)("explains %s", (_, state, __, sentence) => {
    expect(workspaceHoldSentence(state)).toBe(sentence);
  });

  it("names every reason that stands, never just the first", () => {
    // The defect this replaced: the wording was worked out from one flag after
    // the decision had been made from three, so a workspace held only by an
    // edit was explained as an import.
    for (const [, state, reasons] of CASES) {
      // Lower cased, because whichever clause comes first opens the sentence.
      const sentence = workspaceHoldSentence(state).toLowerCase();
      if (reasons.includes("unsavedChanges")) {
        expect(sentence).toContain("have not been saved to a file");
      }
      if (reasons.includes("importing")) {
        expect(sentence).toContain("an import is still running");
      }
      if (reasons.includes("editInFlight")) {
        expect(sentence).toContain("an edit is still being applied");
      }
      if (reasons.includes("editorOpen")) {
        expect(sentence).toContain("a cell is still open for editing");
      }
    }
  });

  it("reads as one sentence, whatever it has to say", () => {
    for (const [, state] of CASES) {
      const sentence = workspaceHoldSentence(state);
      if (sentence === "") {
        continue;
      }
      expect(sentence.charAt(0)).toBe(sentence.charAt(0).toUpperCase());
      expect(sentence.endsWith(".")).toBe(true);
      expect(sentence).not.toContain("..");
      // It is followed by the clause naming what would replace the workspace,
      // so it carries no dangling joiner of its own.
      expect(sentence.endsWith(" and.")).toBe(false);
    }
  });
});

describe("workspaceHoldHeading", () => {
  it.each(CASES)("names %s in two words", (_, state, reasons) => {
    expect(workspaceHoldHeading(state)).toBe(
      reasons.includes("unsavedChanges")
        ? "Unsaved changes"
        : "Work not saved yet",
    );
  });

  it("does not claim a change until one has been made", () => {
    // The badge that means "changes no file has" is gated on the same flag, so
    // a heading that said this over a workspace without it would be the two
    // disagreeing on one screen.
    expect(workspaceHoldHeading({ ...nothingAtStake, editsInFlight: 1 })).toBe(
      "Work not saved yet",
    );
    expect(workspaceHoldHeading({ ...nothingAtStake, importing: true })).toBe(
      "Work not saved yet",
    );
  });
});
