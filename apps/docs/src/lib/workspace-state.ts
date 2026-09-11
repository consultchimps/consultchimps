/**
 * What the data workspace is doing, as one state and one table.
 *
 * The page has four ways of losing work (New, Open, a link out, and Back), four
 * reasons to hold on to it (`workspace-hold`), one worker that runs one command
 * at a time, and a confirmation that stands in front of all of it. Before this
 * module those were eight separate pieces of React state and three refs, and
 * nearly every review finding across #161 and #160 was the same shape: a guard
 * keyed on one condition when the invariant spanned several, or bookkeeping
 * consumed on one path only. The findings lived in the gaps of a table nobody
 * had written down, so the table is the deliverable.
 *
 * Three rules hold the model together.
 *
 * **One state.** `WorkspaceState` is the whole of what the page knows. Nothing
 * else may keep a copy: the busy flag, the unsaved flag, the open-editor flag,
 * the in-flight count, the pending confirmation and the spare-entry count were
 * all separate before, and every one of them was a chance for two answers to
 * the same question to drift apart. The components below the shell report into
 * this through events and read back derived answers; they decide nothing.
 *
 * **A total transition.** `workspaceStep` answers every event in every
 * activity. Where an event cannot arrive (a worker reply for a command that is
 * not running) or is deliberately refused (a stale click on a disabled button),
 * it returns the state it was given, unchanged and by identity. There is no
 * fall-through and no implicit case.
 *
 * **Every answer derived.** Whether editing is locked, whether the page is
 * holding and why, whether the unload listener is installed, whether a spare
 * history entry should be armed, which buttons are enabled, and what the
 * confirmation says are all functions of the state and of nothing else.
 *
 * ## The rule from #174
 *
 * A pending confirmation is answered by the visitor, or by an event that
 * removes all of its reasons from outside the navigation. It is never answered
 * by a side effect of the navigation that raised it.
 *
 * The bug that rule exists for: a Back press with a cell open for editing
 * raised the confirmation, the confirmation locked the grid, the lock cancelled
 * the editor, the cancel released the hold, and the shell's rule from #161 that
 * a question must not outlive its reason then dismissed the question the
 * navigation itself had raised. The draft went with no warning. Both rules were
 * right on their own; they met only where the navigation removed the reason.
 *
 * It is resolved in two halves, one structural and one explicit.
 *
 * Structurally, `editingLocked` reads the activity alone. A standing question
 * no longer locks the grid, so the model never commands an editor closed while
 * a question is up, and an editor that closes then can only be the visitor's
 * own doing. That removes the cause rather than special-casing Back.
 *
 * Explicitly, `settle` dismisses a question only when the event took the last
 * live reason away and that event was not a teardown. A teardown is the grid
 * unmounting (`gridDetached`, which is why it is its own event rather than the
 * zeroed counts it used to report), an editor the grid closed on its way out
 * (which only the grid can know, so it says so), or an editor closing while
 * editing was locked, which is a close the model asked for. Once a teardown has taken the
 * reasons, nothing later can dismiss the question either, because there is no
 * longer a last reason for a later event to remove: only the visitor can answer
 * it. What the question says then falls back to the reasons it captured when it
 * was raised, since a teardown takes the bookkeeping away without making the
 * work safe.
 *
 * The dismissal reads the live reasons rather than the captured ones on
 * purpose. A link held while an import runs, where the import then lands, has a
 * captured set that is empty and a workspace that is now unsaved: dismissing
 * there would leave the visitor on the page with work at stake and a click that
 * silently did nothing.
 */

import {
  mustHoldWorkspace,
  workspaceHoldHeadingFor,
  workspaceHoldReasons,
  workspaceHoldSentenceFor,
  type WorkspaceHoldReason,
  type WorkspaceHoldState,
} from "./workspace-hold";
import type { WorkspaceSummary } from "./workspace-protocol";

/**
 * The one command in flight, or `ready` while the page is idle.
 *
 * There is a single activity for the whole page, not one per section, because
 * the worker runs one command at a time and every long-running command has the
 * same consequence: nothing else may start, and in particular nothing may
 * replace the workspace. A second notion of busy is exactly how a click on New
 * lands behind a running import and throws its result away.
 *
 * `reading` describes a chosen import file and touches no workspace, so it is
 * the one activity that is not a reason to hold.
 *
 * `leaving` is not a command at all: it is the page on its way out, after the
 * visitor has answered for the workspace. It exists because a reply can arrive
 * during the navigation it was answered for, and an accepted edit landing then
 * would mark the workspace unsaved again, re-installing the browser's own
 * warning and pushing a spare history entry into a page that is already going.
 * So the two guards that would ask again stand down there, which is the whole
 * of what the visitor's answer does: what the workspace holds is untouched and
 * every reply is still answered, so nothing is lost track of and nothing is
 * rewritten.
 *
 * It is idle in every other respect. Every derived answer treats it as `ready`,
 * because a page that turns out not to be going anywhere (a navigation the
 * router resolves back to this same route, a traversal the browser declines)
 * must not be left disabled over a live workspace it can no longer save. And
 * anything the visitor does ends it: they are plainly still here, so the page
 * is a workspace again, holding for exactly what it held for before. What they
 * answered stands against replies, not against themselves.
 */
export type WorkspaceActivity =
  | { readonly kind: "ready" }
  | { readonly kind: "creating" }
  | { readonly kind: "opening" }
  | { readonly kind: "reading" }
  | { readonly kind: "importing" }
  | { readonly kind: "saving"; readonly how: SaveMode }
  | { readonly kind: "leaving" };

/** What the visitor asked for. Which file it reaches is decided as it runs. */
export type SaveMode = "save" | "saveAs";

/**
 * Where a save actually landed, reported back once it has.
 *
 * The page cannot know this at the click: writing in place needs a handle a
 * picker granted, Save as needs a picker the browser may not have, and the
 * download is what is left. So the mode is what the state carries and the
 * destination is what the event brings back.
 */
export type SaveDestination = "inPlace" | "saveAs" | "download";

/**
 * The busy value the import section reads, derived from the activity.
 *
 * Neither of the two idle activities is one of them: a page that is ready is not
 * busy, and neither is a page on its way out, which is why `leaving` is excluded
 * rather than left as a value nothing can ever return and every reader would
 * still have to consider.
 */
export type WorkspaceBusy = Exclude<
  WorkspaceActivity["kind"],
  "ready" | "leaving"
> | null;

/** What the page shows once a workspace is held: where it came from and its shape. */
export interface HeldWorkspace {
  /** What the worker holds, including the generation every command quotes. */
  readonly summary: WorkspaceSummary;
  /** The file it was opened from or last saved as, or null for a fresh one. */
  readonly fileName: string | null;
  /**
   * Whether the workspace has changed since it was last written to a file. Set
   * only by a command the worker accepted, and cleared only by a write that
   * resolved.
   */
  readonly unsavedChanges: boolean;
}

/** What the visitor was trying to do when the question was raised. */
export type LeaveIntent =
  | { readonly kind: "new" }
  | { readonly kind: "open" }
  | { readonly kind: "link"; readonly href: string }
  | { readonly kind: "back" };

/**
 * A confirmation waiting for an answer.
 *
 * It is data on the state rather than an activity of its own because it
 * composes with one: a link clicked while an import runs is held, and the
 * import goes on running underneath the question until it lands.
 */
export interface PendingQuestion {
  readonly intent: LeaveIntent;
  /** What was at stake when it was raised, for the #174 rule and its sentence. */
  readonly stake: readonly WorkspaceHoldReason[];
}

/**
 * The one thing the page has to say, as a notice or a failure.
 *
 * One slot rather than two: every path that set one of the old pair cleared the
 * other first, so they were already the same slot with two names.
 */
export interface WorkspaceAnnouncement {
  readonly kind: "notice" | "error";
  readonly text: string;
}

export interface WorkspaceState {
  readonly activity: WorkspaceActivity;
  /** The workspace held, or null while none is open. */
  readonly held: HeldWorkspace | null;
  /** A cell is open for editing, so a draft exists only in an input element. */
  readonly editorOpen: boolean;
  /** Cell edits sent to the worker that have not come back. */
  readonly editsInFlight: number;
  readonly question: PendingQuestion | null;
  /**
   * Spare history entries the Back guard has pushed and not spent.
   *
   * Real state, not a derived answer: a save clears the hold and cannot remove
   * an entry the browser already holds, so the count outlives its reason. It
   * moves only on something that actually happened to history.
   */
  readonly spares: number;
  readonly announcement: WorkspaceAnnouncement | null;
}

export const INITIAL_WORKSPACE_STATE: WorkspaceState = {
  activity: { kind: "ready" },
  held: null,
  editorOpen: false,
  editsInFlight: 0,
  question: null,
  spares: 0,
  announcement: null,
};

/**
 * Everything that can move the state: what a visitor did, and what came back.
 *
 * Named for what happened rather than for what should follow, so the table below
 * is the only place that decides what a happening means.
 */
export type WorkspaceEvent =
  // The visitor
  | { readonly type: "newClicked" }
  | { readonly type: "openClicked" }
  | { readonly type: "saveClicked"; readonly how: SaveMode }
  | { readonly type: "linkClicked"; readonly href: string }
  | { readonly type: "backPressed" }
  | { readonly type: "confirmClicked" }
  | { readonly type: "cancelClicked" }
  | { readonly type: "importFileChosen" }
  | { readonly type: "importClicked" }
  | { readonly type: "editorOpened" }
  // Whether the grid closed this editor on its way out. Tabulator reports a
  // cancel whenever an editor goes, including when the grid it is in is being
  // destroyed, and the two mean opposite things to a standing question: a
  // visitor's Escape answers it, a teardown may not. The grid is the only thing
  // that knows which happened, so it says.
  | { readonly type: "editorClosed"; readonly byTeardown: boolean }
  | { readonly type: "editSent" }
  // The grid went away, taking whatever it was holding with it. Its own event,
  // because the zeroed counts it used to report could not be told apart from a
  // visitor pressing Escape, and the difference is the whole of #174.
  | { readonly type: "gridDetached" }
  // The worker, the browser, and the effects reporting back
  | { readonly type: "openStarted" }
  | { readonly type: "createSucceeded"; readonly summary: WorkspaceSummary }
  | { readonly type: "createFailed"; readonly message: string }
  | {
      readonly type: "openSucceeded";
      readonly summary: WorkspaceSummary;
      readonly fileName: string;
    }
  | { readonly type: "openFailed"; readonly message: string }
  // The picker itself refused, before any workspace was read. Its own event
  // because it happens where an open has not started: reported as an open that
  // failed, it would end an open that is genuinely running.
  | { readonly type: "openPickerFailed"; readonly message: string }
  | {
      readonly type: "saveSucceeded";
      readonly destination: SaveDestination;
      /** The name written to, when a picker chose one. */
      readonly fileName?: string;
    }
  | { readonly type: "saveDismissed" }
  | { readonly type: "saveFailed"; readonly message: string }
  | { readonly type: "importReadFinished" }
  | {
      readonly type: "importSucceeded";
      readonly summary: WorkspaceSummary;
      readonly notice: string;
    }
  | { readonly type: "importFailed" }
  // One reply, carrying what became of the edit. The acceptance and the count
  // move together because they are two halves of one moment: marked first and
  // released second, the hold would lapse between them; released first, a New
  // clicked in that instant would replace the database the edit had just
  // landed in.
  | { readonly type: "editSettled"; readonly accepted: boolean }
  | { readonly type: "historySpareArmed" }
  // The browser handed this page back from its back-forward cache, so a page
  // that had been left is on screen again and is a workspace once more.
  | { readonly type: "pageRestored" };

/** Everything that can only be the visitor themselves acting on the page. */
const VISITOR_EVENTS: ReadonlySet<WorkspaceEvent["type"]> = new Set([
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
]);

/**
 * Whether this event ends `leaving`: everything the visitor themselves does,
 * and the browser handing the page back from its cache.
 *
 * Deliberately not a reply. Those are answered where they are, so the state
 * stays true, but they say nothing about whether this page is still the
 * visitor's: an edit settling, an import landing and a grid going away all
 * happen just as readily to a page that is on its way out.
 *
 * An editor closing is the one that could be either, which is why the grid says
 * which: an Escape is the visitor, a close the grid made on its way out is not.
 */
function endsLeaving(event: WorkspaceEvent): boolean {
  return (
    VISITOR_EVENTS.has(event.type) ||
    (event.type === "editorClosed" && !event.byTeardown)
  );
}

/**
 * What the page must go and do, described rather than done.
 *
 * The model stays pure, and the effects the navigation guards depend on are
 * visible in the table and pinned by tests rather than buried in a handler.
 */
export type WorkspaceEffect =
  | { readonly kind: "create" }
  | { readonly kind: "openFile" }
  | { readonly kind: "save"; readonly how: SaveMode }
  /** Leave, retiring the spare entries. A null href honours a Back press. */
  | {
      readonly kind: "leave";
      readonly href: string | null;
      readonly spares: number;
    }
  | { readonly kind: "armHistoryEntry" }
  /** The click was taken over by the question, so the browser must not follow it. */
  | { readonly kind: "holdNavigation" };

export interface WorkspaceStep {
  readonly state: WorkspaceState;
  readonly effect: WorkspaceEffect | null;
}

/** The workspace's hold state, which is the same four reasons everything reads. */
export function workspaceHoldStateOf(
  state: WorkspaceState,
): WorkspaceHoldState {
  return {
    unsavedChanges: state.held?.unsavedChanges === true,
    importing: state.activity.kind === "importing",
    editsInFlight: state.editsInFlight,
    editorOpen: state.editorOpen,
  };
}

/** Everything at stake, in the order it is explained. */
export function workspaceHoldReasonsOf(
  state: WorkspaceState,
): readonly WorkspaceHoldReason[] {
  return workspaceHoldReasons(workspaceHoldStateOf(state));
}

/** Whether there is anything to lose by leaving or replacing the workspace. */
export function workspaceHoldsWork(state: WorkspaceState): boolean {
  return mustHoldWorkspace(workspaceHoldStateOf(state));
}

/**
 * Whether the grid must open no editor and accept no edit.
 *
 * The activity alone. A standing question deliberately does not lock: it is an
 * inline question, not a modal, and a lock would cancel the very editor the
 * question is there to protect. See the #174 rule above.
 */
export function editingLocked(state: WorkspaceState): boolean {
  return commandInFlight(state);
}

/**
 * Whether the worker is running something for this page.
 *
 * The one question behind the lock, the enabled buttons and the busy value, so
 * the three cannot disagree. `leaving` is not one: the page is on its way out,
 * not busy, and a page that turns out to be staying has to still work.
 */
function commandInFlight(state: WorkspaceState): boolean {
  return state.activity.kind !== "ready" && state.activity.kind !== "leaving";
}

/**
 * Whether the visitor has already answered for this workspace.
 *
 * The two guards below stand down for a page that is leaving, and they are the
 * only answers that do. What the workspace holds is not touched: it is still
 * unsaved if it was, so a page that turns out to be staying, or one the browser
 * hands back, holds for exactly what it held for before rather than reading as
 * saved when it is not. Suppressing the guards where they are asked, rather
 * than in the reasons they are asked about, is what keeps the reasons the one
 * honest account of what is at stake.
 */
function answeredFor(state: WorkspaceState): boolean {
  return state.activity.kind === "leaving";
}

/** Whether the browser's own warning belongs on this page right now. */
export function unloadGuardInstalled(state: WorkspaceState): boolean {
  // Never for the navigation the visitor has just approved: it would ask them
  // a second time, in the browser's own words, for what they already answered.
  return !answeredFor(state) && workspaceHoldsWork(state);
}

/**
 * Whether the Back guard needs a spare history entry pushed.
 *
 * Arm whenever there is something to lose and no spare is held, which covers
 * both the first change and a change made after an earlier spare was spent.
 */
export function shouldArmHistorySpare(state: WorkspaceState): boolean {
  // Not behind a navigation already under way: the entry would outlive the page
  // that pushed it, which is the phantom this guard exists to avoid.
  return !answeredFor(state) && workspaceHoldsWork(state) && state.spares === 0;
}

/** Whether New, Open, Save, Save as, and the two answer buttons are live. */
export function commandsEnabled(state: WorkspaceState): boolean {
  return !commandInFlight(state);
}

/** The busy value handed to the import section, which keeps none of its own. */
export function workspaceBusy(state: WorkspaceState): WorkspaceBusy {
  const { kind } = state.activity;
  return kind === "ready" || kind === "leaving" ? null : kind;
}

/** Which of the page's own buttons is showing that its command is running. */
export type SpinningButton = "new" | "open" | "save" | "saveAs" | null;

/**
 * The one button that spins, or null.
 *
 * A save carries the mode it was asked for, so Save as spins its own button
 * rather than the one beside it: the state knows which was pressed, and a
 * spinner that reads the activity without its mode throws that away.
 */
export function spinningButton(state: WorkspaceState): SpinningButton {
  switch (state.activity.kind) {
    case "creating":
      return "new";
    case "opening":
      return "open";
    case "saving":
      return state.activity.how === "save" ? "save" : "saveAs";
    default:
      // Reading a file and running an import spin the import section's own
      // button, and `leaving` and `ready` spin nothing.
      return null;
  }
}

/** What the confirmation says, or null when there is nothing being asked. */
export interface QuestionView {
  readonly heading: string;
  readonly sentence: string;
  readonly consequence: string;
  readonly confirmLabel: string;
}

const CONSEQUENCES: Readonly<Record<LeaveIntent["kind"], string>> = {
  new: "Starting a new workspace replaces this one",
  open: "Opening another workspace replaces this one",
  link: "Leaving this page closes the workspace",
  back: "Leaving this page closes the workspace",
};

export function questionView(state: WorkspaceState): QuestionView | null {
  const question = state.question;
  if (question === null) {
    return null;
  }
  // The live reasons while there are any, so the sentence describes what is
  // true now rather than what was true when it was raised. The captured stake
  // is the fallback for the one case that empties them without answering: a
  // teardown, which takes the bookkeeping and leaves the work at stake.
  const live = workspaceHoldReasonsOf(state);
  const reasons = live.length > 0 ? live : question.stake;
  const leaving =
    question.intent.kind === "link" || question.intent.kind === "back";
  return {
    heading: workspaceHoldHeadingFor(reasons),
    sentence: workspaceHoldSentenceFor(reasons),
    consequence: CONSEQUENCES[question.intent.kind],
    confirmLabel: leaving
      ? "Discard the changes and leave"
      : "Discard the changes and continue",
  };
}

/** The one place a state is moved, and the only place that decides anything. */
export function workspaceStep(
  state: WorkspaceState,
  event: WorkspaceEvent,
): WorkspaceStep {
  // A page the visitor has answered for is a workspace again the moment they do
  // anything, and that happens here rather than inside the table: both halves
  // below have to judge the same state. The closing rule reads the reasons an
  // event arrived at, and a page in `leaving` is defined to have none, so
  // judging the state the revive replaced would switch that rule off for
  // exactly the events that revive.
  const from = revived(state, event);
  return settle(from, event, transition(from, event));
}

/**
 * The page as the event finds it: itself, or a page that has stopped leaving
 * because the visitor is plainly still here.
 *
 * Their answer stands against replies, not against themselves. A reply is
 * answered where it lands, so nothing is lost track of; what the visitor does
 * has to be answered by a page that is staying, or a New clicked on a page that
 * is not going anywhere after all would be replacing a workspace whose guards
 * are still standing down.
 *
 * What it cannot know is whether the navigation is still on its way: nothing the
 * browser or the router offers says "that transition did not take you
 * anywhere". So a visitor who interacts during a transition that then completes
 * revives a page that really was going, and the guard may arm a spare entry the
 * page then leaves behind: one phantom entry in the history of a page they have
 * left. That is the trade, taken deliberately in this direction, because the
 * other way round is a page that stayed and quietly stopped holding the work
 * done on it.
 */
function revived(state: WorkspaceState, event: WorkspaceEvent): WorkspaceState {
  return state.activity.kind === "leaving" && endsLeaving(event)
    ? { ...state, activity: { kind: "ready" } }
    : state;
}

function stay(state: WorkspaceState): WorkspaceStep {
  return { state, effect: null };
}

/** A workspace whose drafts belong to a database that is no longer held. */
function withoutDrafts(state: WorkspaceState): WorkspaceState {
  return { ...state, editorOpen: false, editsInFlight: 0 };
}

/**
 * The visitor has accepted losing the workspace and the page is on its way out.
 *
 * Nothing about the workspace is rewritten. What changes is that the two guards
 * that would ask again stand down (`answeredFor`), which is what stops the
 * browser's own warning asking a second time, in its own words, for the
 * navigation that was just approved, and stops a spare entry being armed behind
 * it. Rewriting the unsaved flag instead would be a lie the page could be
 * handed back carrying.
 */
function leavingNow(state: WorkspaceState): WorkspaceState {
  return {
    ...state,
    activity: { kind: "leaving" },
    // The effect that goes with this retires them, so the page holds none.
    spares: 0,
  };
}

function asking(state: WorkspaceState, intent: LeaveIntent): WorkspaceState {
  return {
    ...state,
    question: { intent, stake: workspaceHoldReasonsOf(state) },
  };
}

function replacing(state: WorkspaceState, kind: "new" | "open"): WorkspaceStep {
  // A stale click on a disabled button, or a click that raced a command into
  // flight. The worker runs one at a time, so a second command is refused here
  // rather than queued behind a workspace it was never meant for.
  if (state.activity.kind !== "ready") {
    return stay(state);
  }
  if (workspaceHoldsWork(state)) {
    return stay(asking(state, { kind }));
  }
  return begin(state, kind);
}

/**
 * Start the replacement itself.
 *
 * A create is in flight from the moment it is asked for. An open is not: the
 * picker comes first, and the fallback file input reports nothing at all when a
 * visitor closes it, so an activity entered at the click would have no event
 * that could leave it. `openStarted` is dispatched once there are bytes to read.
 */
function begin(state: WorkspaceState, kind: "new" | "open"): WorkspaceStep {
  return kind === "new"
    ? {
        state: { ...state, activity: { kind: "creating" }, announcement: null },
        effect: { kind: "create" },
      }
    : { state, effect: { kind: "openFile" } };
}

const SAVE_NOTICES: Readonly<Record<SaveDestination, string>> = {
  inPlace: "Saved to the workspace file",
  saveAs: "Saved to the workspace file",
  // A download hands the bytes to the browser and the page never learns where
  // they landed, so this is the strongest thing this surface can say.
  download: "Downloaded a copy of the workspace",
};

function transition(
  state: WorkspaceState,
  event: WorkspaceEvent,
): WorkspaceStep {
  // Nothing here special-cases `leaving`. What the visitor does arrives here
  // already revived (`revived` above), and everything else is answered as it
  // would be anywhere, because the release lives in one place: a page the
  // visitor has answered for has both guards standing down (`answeredFor`), so a
  // reply landing during the navigation cannot put the browser's warning back or
  // arm an entry behind it, whatever it carries. Answering it rather than dropping
  // it is what keeps the facts true: an edit that was accepted really did land,
  // and a page that turns out to be staying has to know that rather than believe
  // it was saved. The commands each refuse themselves there anyway, because none
  // of their activities is `leaving`.
  switch (event.type) {
    case "pageRestored":
      // Only ever meaningful while leaving, which the branch above has already
      // taken: the page is a workspace again by the time this is reached.
      return stay(state);

    case "newClicked":
      return replacing(state, "new");

    case "openClicked":
      return replacing(state, "open");

    case "openStarted":
      return state.activity.kind === "ready"
        ? stay({
            ...state,
            activity: { kind: "opening" },
            announcement: null,
          })
        : stay(state);

    case "saveClicked":
      // Saving is never held: it is the way out of holding.
      if (state.activity.kind !== "ready" || state.held === null) {
        return stay(state);
      }
      return {
        state: {
          ...state,
          activity: { kind: "saving", how: event.how },
          announcement: null,
        },
        effect: { kind: "save", how: event.how },
      };

    case "linkClicked":
      if (workspaceHoldsWork(state)) {
        return {
          state: asking(state, { kind: "link", href: event.href }),
          effect: { kind: "holdNavigation" },
        };
      }
      if (state.spares === 0) {
        // Nothing at stake and nothing to retire: the router's own handling is
        // exactly right, so it is left alone.
        return stay(state);
      }
      // Nothing at stake, but a spare from an earlier change is still in the
      // history. A save clears the hold and cannot remove that entry, so the
      // clean way out has to retire it or inherit the phantom.
      return {
        state: { ...state, spares: 0 },
        effect: { kind: "leave", href: event.href, spares: state.spares },
      };

    case "backPressed": {
      if (state.spares === 0) {
        // Not one of ours: the visitor is leaving a page we never armed.
        return stay(state);
      }
      const remaining = state.spares - 1;
      if (workspaceHoldsWork(state)) {
        // Re-arm before asking, so a second press while the question is up
        // lands somewhere harmless too.
        return {
          state: { ...asking(state, { kind: "back" }), spares: remaining },
          effect: { kind: "armHistoryEntry" },
        };
      }
      // Nothing at stake, and the press has already happened. Spending the
      // entry and stopping there would be a press that did nothing, which is
      // the dead press the link guard above already refuses to leave behind.
      return {
        state: { ...state, spares: 0 },
        effect: { kind: "leave", href: null, spares: remaining },
      };
    }

    case "confirmClicked": {
      const question = state.question;
      // Both answer buttons are disabled while a command runs, and confirming
      // a leave then would tear the worker down under it.
      if (question === null || state.activity.kind !== "ready") {
        return stay(state);
      }
      const answered: WorkspaceState = { ...state, question: null };
      switch (question.intent.kind) {
        case "new":
        case "open":
          // The unsaved flag is deliberately left alone: a create or an open
          // that then fails leaves this workspace held, and it has to still
          // read as unsaved or the next New would replace it without a word.
          return begin(answered, question.intent.kind);
        case "link":
          return {
            state: leavingNow(answered),
            effect: {
              kind: "leave",
              href: question.intent.href,
              spares: state.spares,
            },
          };
        case "back":
          return {
            state: leavingNow(answered),
            effect: { kind: "leave", href: null, spares: state.spares },
          };
      }
    }

    case "cancelClicked":
      // The question goes and nothing else moves. In particular an open editor
      // stays open, which is what keeping the workspace means.
      return state.question === null
        ? stay(state)
        : stay({ ...state, question: null });

    case "importFileChosen":
      return state.activity.kind === "ready"
        ? stay({ ...state, activity: { kind: "reading" } })
        : stay(state);

    case "importReadFinished":
      return state.activity.kind === "reading"
        ? stay({ ...state, activity: { kind: "ready" } })
        : stay(state);

    case "importClicked":
      return state.activity.kind === "ready"
        ? stay({
            ...state,
            activity: { kind: "importing" },
            announcement: null,
          })
        : stay(state);

    case "importSucceeded":
      if (state.activity.kind !== "importing" || state.held === null) {
        return stay(state);
      }
      // The summary is replaced wholesale, so the listing always reflects what
      // the worker now holds rather than a count kept in step by hand.
      return stay({
        ...withoutDrafts(state),
        activity: { kind: "ready" },
        held: { ...state.held, summary: event.summary, unsavedChanges: true },
        announcement: { kind: "notice", text: event.notice },
      });

    case "importFailed":
      // The import section explains its own failure against the form it
      // belongs to, so the page says nothing here.
      return state.activity.kind === "importing"
        ? stay({ ...state, activity: { kind: "ready" } })
        : stay(state);

    case "createSucceeded":
      return state.activity.kind === "creating"
        ? stay({
            ...withoutDrafts(state),
            activity: { kind: "ready" },
            held: {
              summary: event.summary,
              fileName: null,
              unsavedChanges: false,
            },
            announcement: {
              kind: "notice",
              text: "Started a new empty workspace",
            },
          })
        : stay(state);

    case "createFailed":
      return state.activity.kind === "creating"
        ? stay({
            ...state,
            activity: { kind: "ready" },
            announcement: { kind: "error", text: event.message },
          })
        : stay(state);

    case "openSucceeded":
      return state.activity.kind === "opening"
        ? stay({
            ...withoutDrafts(state),
            activity: { kind: "ready" },
            held: {
              summary: event.summary,
              fileName: event.fileName,
              unsavedChanges: false,
            },
            announcement: { kind: "notice", text: "Opened the workspace" },
          })
        : stay(state);

    case "openFailed":
      // A file that stopped being readable after it was chosen, or a database
      // the worker rejected. The held workspace is left untouched.
      return state.activity.kind === "opening"
        ? stay({
            ...state,
            activity: { kind: "ready" },
            announcement: { kind: "error", text: event.message },
          })
        : stay(state);

    case "openPickerFailed":
      // Nothing had started, so nothing ends: this only has something to say.
      // While another command is running it says nothing at all, because that
      // command's own outcome is what the visitor is waiting to read.
      return state.activity.kind === "ready"
        ? stay({
            ...state,
            announcement: { kind: "error", text: event.message },
          })
        : stay(state);

    case "saveSucceeded":
      if (state.activity.kind !== "saving" || state.held === null) {
        return stay(state);
      }
      // Cleared here rather than beside the serialize: the bytes only reach the
      // file once the write resolves, and a write that throws has to leave the
      // workspace unsaved.
      return stay({
        ...state,
        activity: { kind: "ready" },
        held: {
          ...state.held,
          fileName: event.fileName ?? state.held.fileName,
          unsavedChanges: false,
        },
        announcement: {
          kind: "notice",
          text: SAVE_NOTICES[event.destination],
        },
      });

    case "saveDismissed":
      // The visitor closed the picker: leave the page exactly as it was.
      return state.activity.kind === "saving"
        ? stay({ ...state, activity: { kind: "ready" } })
        : stay(state);

    case "saveFailed":
      return state.activity.kind === "saving"
        ? stay({
            ...state,
            activity: { kind: "ready" },
            announcement: { kind: "error", text: event.message },
          })
        : stay(state);

    case "editorOpened":
      return state.editorOpen
        ? stay(state)
        : stay({ ...state, editorOpen: true });

    case "editorClosed":
      return state.editorOpen
        ? stay({ ...state, editorOpen: false })
        : stay(state);

    case "editSent":
      return stay({ ...state, editsInFlight: state.editsInFlight + 1 });

    case "editSettled": {
      // Never below zero, so an edit that outlives the grid it was made in
      // cannot release a hold something else is waiting on.
      const settled = {
        ...state,
        editsInFlight: Math.max(0, state.editsInFlight - 1),
      };
      // The worker took it, so the workspace differs from its file. A refused
      // edit changed nothing and marks nothing.
      return stay(
        event.accepted && state.held !== null
          ? { ...settled, held: { ...state.held, unsavedChanges: true } }
          : settled,
      );
    }

    case "gridDetached":
      return state.editorOpen || state.editsInFlight > 0
        ? stay(withoutDrafts(state))
        : stay(state);

    case "historySpareArmed":
      // Never refused, in any activity. It reports something that has already
      // happened to the history, and what the page believes about history has
      // to come from what happened to it: a count that argued with the browser
      // would spend a press the browser never gave it.
      return stay({ ...state, spares: state.spares + 1 });
  }
}

/**
 * Whether the event was the navigation's own doing rather than an answer.
 *
 * The grid unmounting takes the bookkeeping with it, and an editor closing
 * while editing was locked is a close the model asked for. Neither means the
 * work is safe, so neither may answer a question. Everything else that can take
 * a reason away is a worker reply, a completed save, or the visitor closing an
 * editor by hand, and each of those is a real answer.
 */
function isTeardown(before: WorkspaceState, event: WorkspaceEvent): boolean {
  return (
    event.type === "gridDetached" ||
    (event.type === "editorClosed" &&
      // Either the grid closed it on its way out, which only the grid knows,
      // or the model asked for it by locking, which only the model knows.
      (event.byTeardown || editingLocked(before)))
  );
}

/**
 * The closing rule: a question exists only while its reasons do.
 *
 * A save made while it is on screen, or an import that fails after a link was
 * held, answers it by removing what it was about; left standing it would
 * reappear at the next change, asking about something that already happened.
 *
 * It fires only on the event that took the last reason away, which is what
 * stops it firing again later on an event that removed nothing: once a teardown
 * has emptied the reasons, no later event can take a last one away, so only the
 * visitor can answer.
 */
function settle(
  before: WorkspaceState,
  event: WorkspaceEvent,
  step: WorkspaceStep,
): WorkspaceStep {
  if (step.state.question === null) {
    return step;
  }
  if (workspaceHoldReasonsOf(before).length === 0) {
    return step;
  }
  if (workspaceHoldReasonsOf(step.state).length > 0) {
    return step;
  }
  if (isTeardown(before, event)) {
    return step;
  }
  return { state: { ...step.state, question: null }, effect: step.effect };
}
