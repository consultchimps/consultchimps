import { describe, expect, it } from "vitest";

import type { WorkspaceSummary } from "./workspace-protocol";
import {
  commandsEnabled,
  editingLocked,
  INITIAL_WORKSPACE_STATE,
  questionView,
  shouldArmHistorySpare,
  spinningButton,
  unloadGuardInstalled,
  workspaceBusy,
  workspaceHoldsWork,
  workspaceStep,
  type WorkspaceActivity,
  type WorkspaceEffect,
  type WorkspaceEvent,
  type WorkspaceState,
} from "./workspace-state";

function summary(generation: number, tableCount = 1): WorkspaceSummary {
  return {
    generation,
    tableCount,
    schemaFormatVersion: 1,
    tables:
      tableCount === 0
        ? []
        : [
            {
              name: "Customer",
              rowCount: 2,
              recordIdPrefix: "CUST",
              recordIdPadding: 4,
              columns: [{ name: "name", type: "text" }],
            },
          ],
  };
}

const open: WorkspaceState = {
  ...INITIAL_WORKSPACE_STATE,
  held: {
    summary: summary(1),
    fileName: "workspace.sqlite",
    unsavedChanges: false,
  },
};

const unsaved: WorkspaceState = {
  ...open,
  held: {
    summary: summary(1),
    fileName: "workspace.sqlite",
    unsavedChanges: true,
  },
};

/** Apply a run of events, so a test reads as the sequence it is about. */
function play(
  from: WorkspaceState,
  ...events: readonly WorkspaceEvent[]
): WorkspaceState {
  return events.reduce(
    (state, event) => workspaceStep(state, event).state,
    from,
  );
}

/** What the last event asked the page to do, which is where the effects live. */
function effectOf(
  from: WorkspaceState,
  ...events: readonly WorkspaceEvent[]
): WorkspaceEffect | null {
  const last = events[events.length - 1] as WorkspaceEvent;
  return workspaceStep(play(from, ...events.slice(0, -1)), last).effect;
}

// Every activity and every event, named by the compiler rather than by hand: a
// missing key is a type error, so the totality test below cannot fall behind
// the model it is checking.
const activities: { [K in WorkspaceActivity["kind"]]: WorkspaceActivity } = {
  ready: { kind: "ready" },
  creating: { kind: "creating" },
  opening: { kind: "opening" },
  reading: { kind: "reading" },
  importing: { kind: "importing" },
  saving: { kind: "saving", how: "save" },
  leaving: { kind: "leaving" },
};

const events: {
  [K in WorkspaceEvent["type"]]: Extract<WorkspaceEvent, { type: K }>;
} = {
  newClicked: { type: "newClicked" },
  openClicked: { type: "openClicked" },
  openStarted: { type: "openStarted" },
  saveClicked: { type: "saveClicked", how: "save" },
  linkClicked: { type: "linkClicked", href: "https://example.test/docs" },
  backPressed: { type: "backPressed" },
  confirmClicked: { type: "confirmClicked" },
  cancelClicked: { type: "cancelClicked" },
  importFileChosen: { type: "importFileChosen" },
  importReadFinished: { type: "importReadFinished" },
  importClicked: { type: "importClicked" },
  importSucceeded: {
    type: "importSucceeded",
    summary: summary(2),
    notice: "Imported 1 table with 2 rows",
  },
  importFailed: { type: "importFailed" },
  createSucceeded: { type: "createSucceeded", summary: summary(1, 0) },
  createFailed: { type: "createFailed", message: "No" },
  openSucceeded: {
    type: "openSucceeded",
    summary: summary(1),
    fileName: "workspace.sqlite",
  },
  openFailed: { type: "openFailed", message: "No" },
  openPickerFailed: { type: "openPickerFailed", message: "No" },
  saveSucceeded: { type: "saveSucceeded", destination: "inPlace" },
  saveDismissed: { type: "saveDismissed" },
  saveFailed: { type: "saveFailed", message: "No" },
  editorOpened: { type: "editorOpened" },
  editorClosed: { type: "editorClosed", byTeardown: false },
  editSent: { type: "editSent" },
  editSettled: { type: "editSettled", accepted: true },
  gridDetached: { type: "gridDetached" },
  historySpareArmed: { type: "historySpareArmed" },
  pageRestored: { type: "pageRestored" },
};

const everyEvent = Object.values(events) as readonly WorkspaceEvent[];

/** Freeze deeply, so a reducer that writes into its input throws rather than passes. */
function deepFreeze<T>(value: T): T {
  if (typeof value !== "object" || value === null) {
    return value;
  }
  for (const inner of Object.values(value as Record<string, unknown>)) {
    deepFreeze(inner);
  }
  return Object.freeze(value);
}

describe("workspaceStep totality", () => {
  it("answers every event in every activity", () => {
    // The findings this model replaces all lived in the gaps of a table nobody
    // had written down, so the table being total is the deliverable.
    for (const activity of Object.values(activities)) {
      for (const event of everyEvent) {
        const state = deepFreeze<WorkspaceState>({
          ...unsaved,
          activity,
          editsInFlight: 1,
          editorOpen: true,
          spares: 1,
          question: { intent: { kind: "back" }, stake: ["unsavedChanges"] },
        });
        const step = workspaceStep(state, event);
        expect(step.state).toBeDefined();
        expect(step.state.editsInFlight).toBeGreaterThanOrEqual(0);
        expect(step.state.spares).toBeGreaterThanOrEqual(0);
      }
    }
  });

  it("answers every event on a closed workspace as well", () => {
    for (const event of everyEvent) {
      const step = workspaceStep(deepFreeze(INITIAL_WORKSPACE_STATE), event);
      expect(step.state).toBeDefined();
      expect(step.state.held).toBeNull();
    }
  });

  it("never writes into the state it was given", () => {
    for (const activity of Object.values(activities)) {
      for (const event of everyEvent) {
        const state = deepFreeze<WorkspaceState>({
          ...unsaved,
          activity,
          spares: 1,
        });
        const before = JSON.stringify(state);
        workspaceStep(state, event);
        expect(JSON.stringify(state)).toBe(before);
      }
    }
  });

  it("returns the state it was given for an event it ignores", () => {
    // Identity, not equality: an ignored event that rebuilt the state would
    // re-render the page and re-run every effect keyed on it.
    const busy: WorkspaceState = { ...open, activity: { kind: "importing" } };
    expect(workspaceStep(busy, { type: "newClicked" }).state).toBe(busy);
    expect(workspaceStep(busy, { type: "openClicked" }).state).toBe(busy);
    expect(workspaceStep(open, { type: "backPressed" }).state).toBe(open);
    expect(workspaceStep(open, { type: "cancelClicked" }).state).toBe(open);
  });
});

describe("derived answers", () => {
  const samples: readonly WorkspaceState[] = [
    INITIAL_WORKSPACE_STATE,
    open,
    unsaved,
    { ...unsaved, activity: { kind: "importing" } },
    { ...open, editorOpen: true },
    { ...open, editsInFlight: 2 },
    {
      ...unsaved,
      question: { intent: { kind: "back" }, stake: ["unsavedChanges"] },
    },
  ];

  it("answers the same way every time it is asked", () => {
    for (const state of samples) {
      expect(workspaceHoldsWork(state)).toBe(workspaceHoldsWork(state));
      expect(editingLocked(state)).toBe(editingLocked(state));
      expect(workspaceBusy(state)).toBe(workspaceBusy(state));
      expect(questionView(state)).toStrictEqual(questionView(state));
      expect(shouldArmHistorySpare(state)).toBe(shouldArmHistorySpare(state));
    }
  });

  it("locks editing for a command in flight and for nothing else", () => {
    expect(editingLocked(open)).toBe(false);
    expect(editingLocked({ ...open, activity: { kind: "importing" } })).toBe(
      true,
    );
    expect(
      editingLocked({ ...open, activity: { kind: "saving", how: "save" } }),
    ).toBe(true);
    expect(editingLocked({ ...open, activity: { kind: "reading" } })).toBe(
      true,
    );
  });

  it("does not lock editing while a question is standing", () => {
    // The whole of #174: a question that locked the grid cancelled the editor
    // it had been raised about, and that cancel then dismissed the question.
    const armed = play(
      { ...open, editorOpen: true, spares: 1 },
      { type: "backPressed" },
    );
    expect(armed.question).not.toBeNull();
    expect(editingLocked(armed)).toBe(false);
  });

  it("installs the unload guard exactly while something is at stake", () => {
    expect(unloadGuardInstalled(open)).toBe(false);
    expect(unloadGuardInstalled(unsaved)).toBe(true);
    expect(unloadGuardInstalled({ ...open, editorOpen: true })).toBe(true);
    expect(unloadGuardInstalled({ ...open, editsInFlight: 1 })).toBe(true);
    expect(
      unloadGuardInstalled({ ...open, activity: { kind: "importing" } }),
    ).toBe(true);
  });

  it("enables the commands only when nothing is in flight", () => {
    expect(commandsEnabled(unsaved)).toBe(true);
    expect(
      commandsEnabled({
        ...unsaved,
        activity: { kind: "saving", how: "saveAs" },
      }),
    ).toBe(false);
  });

  it("reports the activity as the busy value the import section reads", () => {
    expect(workspaceBusy(open)).toBeNull();
    expect(workspaceBusy({ ...open, activity: { kind: "reading" } })).toBe(
      "reading",
    );
    expect(
      workspaceBusy({ ...open, activity: { kind: "saving", how: "saveAs" } }),
    ).toBe("saving");
  });
});

describe("starting and opening a workspace", () => {
  it("creates at once when nothing is at stake", () => {
    expect(effectOf(open, { type: "newClicked" })).toStrictEqual({
      kind: "create",
    });
    expect(play(open, { type: "newClicked" }).activity.kind).toBe("creating");
  });

  it("asks first when something is at stake", () => {
    const asked = play(unsaved, { type: "newClicked" });
    expect(asked.activity.kind).toBe("ready");
    expect(asked.question?.intent).toStrictEqual({ kind: "new" });
    expect(asked.question?.stake).toStrictEqual(["unsavedChanges"]);
    expect(effectOf(unsaved, { type: "newClicked" })).toBeNull();
  });

  it("holds New for an edit that has been sent and not answered", () => {
    // Nothing is unsaved yet, because an edit counts as unsaved only once the
    // worker has accepted it. The count is what covers the gap.
    const sending = play(open, { type: "editSent" });
    expect(play(sending, { type: "newClicked" }).question?.stake).toStrictEqual(
      ["editInFlight"],
    );
  });

  it("keeps the workspace unsaved through a create the visitor confirmed", () => {
    // A create that then fails has to leave the old workspace still reading as
    // unsaved, or the next New would replace it without a word.
    const confirmed = play(
      unsaved,
      { type: "newClicked" },
      { type: "confirmClicked" },
    );
    expect(confirmed.held?.unsavedChanges).toBe(true);
    const failed = play(confirmed, { type: "createFailed", message: "No" });
    expect(failed.held?.unsavedChanges).toBe(true);
    expect(failed.announcement).toStrictEqual({ kind: "error", text: "No" });
  });

  it("voids the drafts of the workspace it replaced", () => {
    const busyGrid = play(
      unsaved,
      { type: "editorOpened" },
      { type: "editSent" },
    );
    const replaced = play(
      { ...busyGrid, activity: { kind: "creating" } },
      { type: "createSucceeded", summary: summary(2, 0) },
    );
    expect(replaced.editorOpen).toBe(false);
    expect(replaced.editsInFlight).toBe(0);
    expect(replaced.held?.unsavedChanges).toBe(false);
    expect(replaced.held?.summary.generation).toBe(2);
  });

  it("never counts an edit that settles after its grid is gone below zero", () => {
    const detached = play(
      { ...open, editsInFlight: 1 },
      { type: "gridDetached" },
    );
    expect(
      play(detached, { type: "editSettled", accepted: false }).editsInFlight,
    ).toBe(0);
  });

  it("runs the picker before it counts the open as in flight", () => {
    // The fallback file input reports nothing when a visitor closes it, so a
    // busy activity entered at the click would have no event to leave it.
    expect(effectOf(open, { type: "openClicked" })).toStrictEqual({
      kind: "openFile",
    });
    expect(play(open, { type: "openClicked" }).activity.kind).toBe("ready");
    expect(
      play(open, { type: "openClicked" }, { type: "openStarted" }).activity
        .kind,
    ).toBe("opening");
  });

  it("reports a picker that failed without moving the activity", () => {
    const failed = play(open, { type: "openPickerFailed", message: "No" });
    expect(failed.activity.kind).toBe("ready");
    expect(failed.announcement).toStrictEqual({ kind: "error", text: "No" });
  });

  it("does not let a picker that failed end an open that is running", () => {
    // Two events rather than one, because a picker refusing before anything was
    // read and a file that stopped being readable are different things: reported
    // alike, the first would end the second.
    const opening = play(
      open,
      { type: "openClicked" },
      { type: "openStarted" },
    );
    const during = play(opening, { type: "openPickerFailed", message: "No" });
    expect(during.activity.kind).toBe("opening");
    expect(during.announcement).toBeNull();
    expect(
      play(during, {
        type: "openSucceeded",
        summary: summary(5),
        fileName: "other.sqlite",
      }).held?.summary.generation,
    ).toBe(5);
  });

  it("refuses a second open while the first is still running", () => {
    // The worker would replace its database for a page that had stopped
    // listening, and the summary on screen would name a workspace it no longer
    // holds. The shell reads the refusal off the state it was handed back, so
    // it has to come back by identity from the whole step, closing rule
    // included, and not merely from the table.
    const opening = play(
      unsaved,
      { type: "openClicked" },
      { type: "openStarted" },
    );
    expect(workspaceStep(opening, { type: "openStarted" }).state).toBe(opening);
    const asked: WorkspaceState = {
      ...opening,
      question: { intent: { kind: "back" }, stake: ["unsavedChanges"] },
    };
    expect(workspaceStep(asked, { type: "openStarted" }).state).toBe(asked);
  });

  it("takes the opened file as the source and starts clean", () => {
    const opened = play(
      { ...unsaved, editorOpen: true },
      { type: "openClicked" },
      { type: "openStarted" },
      { type: "openSucceeded", summary: summary(4), fileName: "other.sqlite" },
    );
    expect(opened.held?.fileName).toBe("other.sqlite");
    expect(opened.held?.unsavedChanges).toBe(false);
    expect(opened.editorOpen).toBe(false);
  });
});

describe("saving", () => {
  const saving = play(unsaved, { type: "saveClicked", how: "save" });

  it("clears the unsaved flag only once the write resolved", () => {
    expect(saving.held?.unsavedChanges).toBe(true);
    const saved = play(saving, {
      type: "saveSucceeded",
      destination: "inPlace",
    });
    expect(saved.held?.unsavedChanges).toBe(false);
    expect(saved.announcement?.text).toBe("Saved to the workspace file");
  });

  it("keeps the workspace unsaved when the write failed", () => {
    const failed = play(saving, { type: "saveFailed", message: "No" });
    expect(failed.held?.unsavedChanges).toBe(true);
    expect(failed.activity.kind).toBe("ready");
  });

  it("says nothing when the visitor closed the save picker", () => {
    const dismissed = play(saving, { type: "saveDismissed" });
    expect(dismissed.activity.kind).toBe("ready");
    expect(dismissed.announcement).toBeNull();
    expect(dismissed.held?.unsavedChanges).toBe(true);
  });

  it("takes the name a Save as wrote to", () => {
    const saved = play(
      unsaved,
      { type: "saveClicked", how: "saveAs" },
      { type: "saveSucceeded", destination: "saveAs", fileName: "copy.sqlite" },
    );
    expect(saved.held?.fileName).toBe("copy.sqlite");
    expect(saved.announcement?.text).toBe("Saved to the workspace file");
  });

  it("says a download is a copy, because the page never learns where it landed", () => {
    const saved = play(
      unsaved,
      { type: "saveClicked", how: "saveAs" },
      { type: "saveSucceeded", destination: "download" },
    );
    expect(saved.held?.fileName).toBe("workspace.sqlite");
    expect(saved.announcement?.text).toBe("Downloaded a copy of the workspace");
  });
});

describe("importing", () => {
  it("marks the workspace unsaved and replaces the listing", () => {
    const imported = play(
      open,
      { type: "importClicked" },
      {
        type: "importSucceeded",
        summary: summary(3),
        notice: "Imported 1 table with 2 rows",
      },
    );
    expect(imported.held?.unsavedChanges).toBe(true);
    expect(imported.held?.summary.generation).toBe(3);
    expect(imported.announcement).toStrictEqual({
      kind: "notice",
      text: "Imported 1 table with 2 rows",
    });
  });

  it("leaves the announcement to the import section when it fails", () => {
    const failed = play(
      open,
      { type: "importClicked" },
      { type: "importFailed" },
    );
    expect(failed.activity.kind).toBe("ready");
    expect(failed.announcement).toBeNull();
  });

  it("holds the page while an import is in flight", () => {
    const running = play(open, { type: "importClicked" });
    expect(workspaceHoldsWork(running)).toBe(true);
    expect(unloadGuardInstalled(running)).toBe(true);
  });

  it("does not hold the page for a file it is only describing", () => {
    // Reading a file to describe it touches no workspace, so a question in
    // front of it would be a question with nothing at stake.
    const reading = play(open, { type: "importFileChosen" });
    expect(workspaceHoldsWork(reading)).toBe(false);
    expect(play(reading, { type: "importReadFinished" }).activity.kind).toBe(
      "ready",
    );
  });
});

describe("a standing question", () => {
  const importing = play(open, { type: "importClicked" });
  const heldLink = play(importing, {
    type: "linkClicked",
    href: "https://example.test/docs",
  });

  it("holds a link out of the page and does not follow it", () => {
    expect(heldLink.question?.intent).toStrictEqual({
      kind: "link",
      href: "https://example.test/docs",
    });
    expect(
      effectOf(importing, {
        type: "linkClicked",
        href: "https://example.test/docs",
      }),
    ).toStrictEqual({ kind: "holdNavigation" });
  });

  it("goes when a failed import leaves nothing to ask about", () => {
    // #161: a question must not outlive its reason, or it stands over work that
    // already went and reappears at the next change.
    expect(play(heldLink, { type: "importFailed" }).question).toBeNull();
  });

  it("stands when the import lands, because the workspace is now unsaved", () => {
    // The visitor's click is still unanswered and there is still something to
    // lose, so the question stays and its sentence moves to the live reason.
    const landed = play(heldLink, {
      type: "importSucceeded",
      summary: summary(2),
      notice: "Imported 1 table with 2 rows",
    });
    expect(landed.question).not.toBeNull();
    expect(questionView(landed)?.sentence).toBe(
      "This workspace has changes that have not been saved to a file.",
    );
  });

  it("goes when a save removes what it was about", () => {
    const asked = play(unsaved, { type: "newClicked" });
    const saved = play(
      asked,
      { type: "saveClicked", how: "save" },
      { type: "saveSucceeded", destination: "inPlace" },
    );
    expect(saved.question).toBeNull();
  });

  it("goes when the last edit it was about settles refused", () => {
    // A refused edit changed nothing, so there is nothing left to ask about.
    const asked = play(open, { type: "editSent" }, { type: "newClicked" });
    expect(asked.question).not.toBeNull();
    expect(
      play(asked, { type: "editSettled", accepted: false }).question,
    ).toBeNull();
  });

  it("stands when that edit settles accepted, because the hold never lapses", () => {
    // The acceptance and the count move in one transition: marked first and
    // released second the hold would lapse between them, and released first a
    // New clicked in that instant would replace the database it landed in.
    const asked = play(open, { type: "editSent" }, { type: "newClicked" });
    const landed = play(asked, { type: "editSettled", accepted: true });
    expect(landed.editsInFlight).toBe(0);
    expect(landed.held?.unsavedChanges).toBe(true);
    expect(landed.question).not.toBeNull();
    expect(questionView(landed)?.sentence).toBe(
      "This workspace has changes that have not been saved to a file.",
    );
  });

  it("is replaced, not stacked, when the visitor tries something else", () => {
    // The most recent thing the visitor tried is what they are asked about.
    const held = play(unsaved, {
      type: "linkClicked",
      href: "https://example.test/docs",
    });
    expect(play(held, { type: "newClicked" }).question?.intent).toStrictEqual({
      kind: "new",
    });
  });

  it("is answered by the visitor keeping the workspace, changing nothing else", () => {
    const kept = play(
      { ...unsaved, editorOpen: true },
      { type: "newClicked" },
      { type: "cancelClicked" },
    );
    expect(kept.question).toBeNull();
    expect(kept.editorOpen).toBe(true);
    expect(kept.held?.unsavedChanges).toBe(true);
  });

  it("says what it is standing in front of", () => {
    const view = questionView(heldLink);
    expect(view?.heading).toBe("Work not saved yet");
    expect(view?.sentence).toBe("An import is still running.");
    expect(view?.consequence).toBe("Leaving this page closes the workspace");
    expect(view?.confirmLabel).toBe("Discard the changes and leave");
    expect(questionView(play(unsaved, { type: "newClicked" }))).toStrictEqual({
      heading: "Unsaved changes",
      sentence:
        "This workspace has changes that have not been saved to a file.",
      consequence: "Starting a new workspace replaces this one",
      confirmLabel: "Discard the changes and continue",
    });
    expect(
      questionView(play(unsaved, { type: "openClicked" }))?.consequence,
    ).toBe("Opening another workspace replaces this one");
    expect(questionView(open)).toBeNull();
  });

  it("cannot be answered while a command is in flight", () => {
    // Both answer buttons are disabled then, and confirming would tear the
    // worker down under the command it is still running.
    expect(play(heldLink, { type: "confirmClicked" }).question).not.toBeNull();
  });

  it("starts the work the visitor accepted losing the workspace for", () => {
    expect(
      effectOf(unsaved, { type: "newClicked" }, { type: "confirmClicked" }),
    ).toStrictEqual({ kind: "create" });
    expect(
      effectOf(unsaved, { type: "openClicked" }, { type: "confirmClicked" }),
    ).toStrictEqual({ kind: "openFile" });
    expect(
      effectOf(
        unsaved,
        { type: "linkClicked", href: "https://example.test/docs" },
        { type: "confirmClicked" },
      ),
    ).toStrictEqual({
      kind: "leave",
      href: "https://example.test/docs",
      spares: 0,
    });
  });
});

describe("the #174 rule", () => {
  // Back with a cell open for editing. The draft exists only in an input
  // element, so the question is the only thing between it and a navigation that
  // has already happened.
  const asked = play(
    { ...open, editorOpen: true, spares: 1 },
    { type: "backPressed" },
  );

  it("raises a question that names the draft", () => {
    expect(asked.question?.intent).toStrictEqual({ kind: "back" });
    expect(asked.question?.stake).toStrictEqual(["editorOpen"]);
    expect(questionView(asked)?.sentence).toBe(
      "A cell is still open for editing.",
    );
  });

  it("is not answered by the grid being torn down", () => {
    const detached = play(asked, { type: "gridDetached" });
    expect(detached.question).not.toBeNull();
    // What it was raised about is what it keeps saying: a teardown takes the
    // bookkeeping away without making the work safe.
    expect(questionView(detached)?.sentence).toBe(
      "A cell is still open for editing.",
    );
  });

  it("is not answered by an editor the model itself closed", () => {
    const cancelled = play(
      asked,
      { type: "saveClicked", how: "save" },
      { type: "editorClosed", byTeardown: false },
    );
    expect(cancelled.question).not.toBeNull();
  });

  it("cannot be answered by a later event once a teardown took its reasons", () => {
    const detached = play(asked, { type: "gridDetached" });
    expect(
      play(detached, { type: "editSettled", accepted: false }).question,
    ).not.toBeNull();
    expect(
      play(detached, { type: "historySpareArmed" }).question,
    ).not.toBeNull();
    expect(
      play(detached, { type: "editorClosed", byTeardown: false }).question,
    ).not.toBeNull();
    expect(
      play(
        detached,
        { type: "saveClicked", how: "save" },
        {
          type: "saveSucceeded",
          destination: "inPlace",
        },
      ).question,
    ).not.toBeNull();
  });

  it("is not answered by an editor the grid closed on its way out", () => {
    // The grid's own rebuild destroys its Tabulator instance, and Tabulator
    // reports that as a cancel like any other. Only the grid can tell the two
    // apart, so it says which, and a teardown is not an answer even when the
    // page is idle and nothing was locked.
    const rebuilt = play(asked, { type: "editorClosed", byTeardown: true });
    expect(editingLocked(asked)).toBe(false);
    expect(rebuilt.question).not.toBeNull();
    expect(questionView(rebuilt)?.sentence).toBe(
      "A cell is still open for editing.",
    );
  });

  it("is answered by the visitor closing the editor themselves", () => {
    // Nothing commanded this one, so it is the visitor's own doing and the
    // draft went with it: there is nothing left to ask about.
    const escaped = play(asked, { type: "editorClosed", byTeardown: false });
    expect(escaped.question).toBeNull();
    expect(workspaceHoldsWork(escaped)).toBe(false);
  });

  it("keeps the editor open when the visitor keeps the workspace", () => {
    const kept = play(asked, { type: "cancelClicked" });
    expect(kept.question).toBeNull();
    expect(kept.editorOpen).toBe(true);
    expect(editingLocked(kept)).toBe(false);
  });

  it("releases the whole hold when the visitor confirms the leave", () => {
    // The browser's own warning must not ask a second time, in its own words,
    // for the navigation the visitor has just approved, and that means every
    // reason at once rather than the unsaved flag alone.
    const leaving = play(
      { ...unsaved, editorOpen: true, editsInFlight: 1, spares: 1 },
      { type: "backPressed" },
      { type: "historySpareArmed" },
      { type: "confirmClicked" },
    );
    expect(unloadGuardInstalled(leaving)).toBe(false);
    expect(shouldArmHistorySpare(leaving)).toBe(false);
    // Released, not rewritten: what was true of the workspace is still true, so
    // a page that turns out to be staying holds for it again rather than
    // reading as saved when it is not.
    expect(leaving.editorOpen).toBe(true);
    expect(leaving.editsInFlight).toBe(1);
    expect(leaving.held?.unsavedChanges).toBe(true);
  });

  it("steps back over the spares and this page together", () => {
    // The press spent one entry and the re-arm pushed another, so honouring it
    // means stepping back over both that entry and this page.
    expect(
      effectOf(
        { ...unsaved, spares: 1 },
        { type: "backPressed" },
        { type: "historySpareArmed" },
        { type: "confirmClicked" },
      ),
    ).toStrictEqual({ kind: "leave", href: null, spares: 1 });
  });
});

describe("a page the visitor has answered for", () => {
  // An edit is in flight, the visitor is asked about it, and they choose to
  // leave. The worker's reply is still coming, and a page that acted on it
  // would mark the workspace unsaved again behind a navigation that has
  // already started: the browser's own warning would come back, and the guard
  // would push a spare history entry into a page that is on its way out.
  const leaving = play(
    { ...open, editsInFlight: 1, spares: 1 },
    { type: "backPressed" },
    { type: "historySpareArmed" },
    { type: "confirmClicked" },
  );

  /**
   * Everything the visitor themselves can do, which is what ends `leaving`. An
   * editor closing counts when the grid says the visitor closed it.
   */
  const endsLeaving = (event: WorkspaceEvent): boolean =>
    new Set<WorkspaceEvent["type"]>([
      "newClicked",
      "openClicked",
      "saveClicked",
      "linkClicked",
      "backPressed",
      "confirmClicked",
      "cancelClicked",
      "importFileChosen",
      "importClicked",
      "editorOpened",
      "editSent",
      "pageRestored",
    ]).has(event.type) ||
    (event.type === "editorClosed" && !event.byTeardown);

  /** The same page, over an unsaved workspace. */
  const leavingUnsaved = play(
    { ...unsaved, spares: 1 },
    { type: "backPressed" },
    { type: "historySpareArmed" },
    { type: "confirmClicked" },
  );

  it("asks nothing more, and is otherwise as idle as any other page", () => {
    expect(leaving.activity.kind).toBe("leaving");
    // What it holds is untouched, which is what keeps the reasons the one
    // honest account of what is at stake. The two guards that would ask about
    // them again are the only answers that stand down.
    expect(workspaceHoldsWork(leaving)).toBe(true);
    expect(unloadGuardInstalled(leaving)).toBe(false);
    expect(shouldArmHistorySpare(leaving)).toBe(false);
    // Not disabled. A navigation the router resolves back to this same route
    // never unmounts this page, and a page left disabled there would be a live
    // workspace nobody could save.
    expect(commandsEnabled(leaving)).toBe(true);
    expect(editingLocked(leaving)).toBe(false);
    expect(workspaceBusy(leaving)).toBeNull();
    expect(spinningButton(leaving)).toBeNull();
  });

  it("cannot be put back into holding by a reply that lands during it", () => {
    // The reply is answered, because the edit really did land and a page that
    // turns out to be staying has to know that. What it cannot do is put the
    // browser's warning back or arm an entry behind a navigation already under
    // way, and it cannot, because a page answered for holds nothing.
    const landed = play(leaving, { type: "editSettled", accepted: true });
    expect(landed.activity.kind).toBe("leaving");
    expect(landed.editsInFlight).toBe(0);
    expect(landed.held?.unsavedChanges).toBe(true);
    expect(unloadGuardInstalled(landed)).toBe(false);
    expect(shouldArmHistorySpare(landed)).toBe(false);
    // And the page that stayed knows the truth about it.
    expect(workspaceHoldsWork(play(landed, { type: "editorOpened" }))).toBe(
      true,
    );
  });

  it("stays on its way out for every reply and every report", () => {
    for (const event of everyEvent) {
      if (endsLeaving(event)) {
        continue;
      }
      const after = workspaceStep(leaving, event);
      expect(after.state.activity.kind).toBe("leaving");
      expect(after.effect).toBeNull();
      expect(unloadGuardInstalled(after.state)).toBe(false);
      expect(shouldArmHistorySpare(after.state)).toBe(false);
      // Nothing claims a spare entry the page retired on its way out. The one
      // event that would is the report that the browser was given one, which is
      // never refused: see the case for it.
      expect(after.state.spares).toBe(
        event.type === "historySpareArmed" ? 1 : 0,
      );
    }
  });

  it("is a workspace again the moment the visitor does anything", () => {
    // They are plainly still here, so the page goes back to being a workspace
    // with everything that was true of it still true. What they answered stands
    // against replies, not against themselves.
    const editing = play(leavingUnsaved, { type: "editorOpened" });
    expect(editing.activity.kind).toBe("ready");
    expect(editing.editorOpen).toBe(true);
    expect(workspaceHoldsWork(editing)).toBe(true);
    expect(shouldArmHistorySpare(editing)).toBe(true);
  });

  it("asks again about a workspace it turns out not to be leaving", () => {
    // The unsaved workspace was never rewritten as saved, so a New clicked on a
    // page that stayed is held for it exactly as it was before.
    const stayed = play(leavingUnsaved, { type: "newClicked" });
    expect(stayed.question?.intent).toStrictEqual({ kind: "new" });
    expect(stayed.question?.stake).toStrictEqual(["unsavedChanges"]);
  });

  it("is revived before the table and the closing rule, not inside them", () => {
    // Both halves have to judge the same page. A question cannot stand here as
    // the model is written, but the wiring is what is being pinned: the reasons
    // the closing rule reads are the revived page's, so a reviving event is
    // judged as it would be on any other page rather than against a page that
    // is defined to hold nothing.
    const asked: WorkspaceState = {
      ...leavingUnsaved,
      question: { intent: { kind: "back" }, stake: ["unsavedChanges"] },
    };
    const replaced = play(asked, { type: "newClicked" });
    expect(replaced.activity.kind).toBe("ready");
    expect(replaced.question?.intent).toStrictEqual({ kind: "new" });
  });

  it("leaves a Back press it did not arm to the browser", () => {
    // Its spares were retired on the way out, so the press is not one of ours
    // and the page does not take it over.
    const step = workspaceStep(leaving, { type: "backPressed" });
    expect(step.effect).toBeNull();
    expect(step.state.activity.kind).toBe("ready");
  });

  it("is a workspace again when the browser hands the page back", () => {
    // Restored from the back-forward cache by a Forward press. The tab was
    // frozen rather than torn down, so the worker still holds what it held, and
    // an unsaved workspace is unsaved again rather than reading as saved.
    const restored = play(leavingUnsaved, { type: "pageRestored" });
    expect(restored.activity.kind).toBe("ready");
    expect(restored.held?.summary.generation).toBe(1);
    expect(restored.held?.unsavedChanges).toBe(true);
    expect(unloadGuardInstalled(restored)).toBe(true);
    expect(commandsEnabled(restored)).toBe(true);
  });

  it("is not something a live page can be restored out of", () => {
    expect(workspaceStep(unsaved, { type: "pageRestored" }).state).toBe(
      unsaved,
    );
  });
});

describe("what a section reports", () => {
  // The grid releases what it holds in an unmount cleanup keyed on the
  // identity of these callbacks, so the shell hands it callbacks that depend on
  // nothing, which it can only do while none of these events asks the page to
  // go and do something. Pinned here rather than assumed there.
  const reported: readonly WorkspaceEvent["type"][] = [
    "editorOpened",
    "editorClosed",
    "editSent",
    "editSettled",
    "gridDetached",
    "importFileChosen",
    "importReadFinished",
    "importClicked",
    "importSucceeded",
    "importFailed",
  ];

  it("never asks the page to do anything", () => {
    for (const type of reported) {
      const event = events[type];
      for (const activity of Object.values(activities)) {
        for (const question of [
          null,
          { intent: { kind: "back" } as const, stake: ["editorOpen"] as const },
        ]) {
          const state: WorkspaceState = {
            ...unsaved,
            activity,
            editorOpen: true,
            editsInFlight: 1,
            spares: 1,
            question,
          };
          expect(workspaceStep(state, event).effect).toBeNull();
        }
      }
    }
  });
});

describe("the Back guard's spare history entries", () => {
  it("arms one whenever something is at stake and none is held", () => {
    expect(shouldArmHistorySpare(open)).toBe(false);
    expect(shouldArmHistorySpare(unsaved)).toBe(true);
    expect(shouldArmHistorySpare({ ...unsaved, spares: 1 })).toBe(false);
    expect(shouldArmHistorySpare({ ...open, editorOpen: true })).toBe(true);
  });

  it("ignores a press on an entry it never armed", () => {
    const pressed = workspaceStep(unsaved, { type: "backPressed" });
    expect(pressed.state).toBe(unsaved);
    expect(pressed.effect).toBeNull();
  });

  it("spends one and arms another while the question is up", () => {
    // A second press while the question is showing has to land somewhere
    // harmless too, so the spare is replaced before the asking.
    const step = workspaceStep(
      { ...unsaved, spares: 1 },
      { type: "backPressed" },
    );
    expect(step.effect).toStrictEqual({ kind: "armHistoryEntry" });
    expect(step.state.spares).toBe(0);
    expect(play(step.state, { type: "historySpareArmed" }).spares).toBe(1);
  });

  it("spends the spare and leaves when nothing is at stake", () => {
    // A save clears the hold and cannot remove the entry the browser holds. A
    // press that only spent that entry would be a press that did nothing, which
    // is the dead press the link guard already refuses to leave behind.
    const step = workspaceStep({ ...open, spares: 1 }, { type: "backPressed" });
    expect(step.effect).toStrictEqual({ kind: "leave", href: null, spares: 0 });
    expect(step.state.spares).toBe(0);
    expect(step.state.question).toBeNull();
  });

  it("retires the spare when a clean page follows a link", () => {
    expect(
      effectOf(
        { ...open, spares: 1 },
        {
          type: "linkClicked",
          href: "https://example.test/docs",
        },
      ),
    ).toStrictEqual({
      kind: "leave",
      href: "https://example.test/docs",
      spares: 1,
    });
  });

  it("leaves a clean page with no spare to the router", () => {
    const step = workspaceStep(open, {
      type: "linkClicked",
      href: "https://example.test/docs",
    });
    expect(step.effect).toBeNull();
    expect(step.state).toBe(open);
  });
});
