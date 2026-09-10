/**
 * Whether there is anything to lose by leaving or replacing the workspace, and
 * what it is.
 *
 * Every guard the shell has asks this one question: New, Open, a link out of
 * the page, the Back button, and the browser's own unload warning. Any of them
 * keyed on part of the answer is a hole, so the answer lives here, once, and
 * they all read it.
 *
 * The sentence the visitor reads is derived from the same answer, for the same
 * reason. A confirmation that works out for itself which state it is explaining
 * is a second reading of the state, and two readings drift: the wording said "an
 * import is still running" for a workspace held only by an edit on its way to
 * the worker, because it inferred the reason from one flag after the decision
 * had been made from three. What is being asked about and why are one thing.
 *
 * It is a module rather than an expression inside the component because the
 * conditions are what the guarding is, and each one arrived after a way of
 * losing work was found. Keeping them here means a new one can be added with
 * the case that motivated it, its wording beside it, and every guard and every
 * sentence inherits it by doing nothing.
 */

/** What the shell knows about work the file does not hold. */
export interface WorkspaceHoldState {
  /** The workspace has changes no file has. */
  readonly unsavedChanges: boolean;
  /** An import has been sent to the worker and has not come back. */
  readonly importing: boolean;
  /** Cell edits sent to the worker that have not come back. */
  readonly editsInFlight: number;
  /**
   * A cell is open for editing.
   *
   * Held from the moment the editor opens rather than from the first keystroke,
   * which is a deliberate choice of the coarser rule. Holding on open cannot
   * miss a keystroke: there is no event to observe, no editor type to get
   * wrong, and no editor added later that fires something unexpected. Holding
   * on the first change would be finer, and being wrong about it loses a draft
   * with no warning, which is the failure this exists to remove.
   *
   * It costs almost nothing, because an editor commits or cancels when it loses
   * focus, and every click inside the page moves focus before the click is
   * handled. What it covers is the ways of leaving that never blur anything:
   * the Back button, and closing or reloading the tab, where the browser's own
   * warning has to be installed before the leaving starts.
   */
  readonly editorOpen: boolean;
}

/** One thing the workspace holds that no file does. */
export type WorkspaceHoldReason =
  "unsavedChanges" | "importing" | "editInFlight" | "editorOpen";

/**
 * How each reason is worded, once, as a clause that reads inside a sentence.
 *
 * Lower case and without a stop, so the joining below decides where the
 * sentence starts and ends rather than every clause carrying its own idea.
 */
const REASON_CLAUSES: Readonly<Record<WorkspaceHoldReason, string>> = {
  // The obvious one: changes that exist only in this tab's memory.
  unsavedChanges:
    "this workspace has changes that have not been saved to a file",
  // An import that has not come back is the one command whose result exists
  // nowhere else, so leaving mid-import destroys work that was never anywhere
  // but this tab.
  importing: "an import is still running",
  // A cell edit that has not come back is the same thing at a smaller scale,
  // and it is the harder one to see. The workspace still looks clean, because
  // an edit counts as unsaved only once the worker has accepted it, which is
  // the only honest moment to count it. Between the two, a visitor who commits
  // an edit and reaches straight for New would be replacing a database the edit
  // is still on its way to, and the commands run in the order they were sent:
  // the edit would land, and the workspace holding it would be discarded a
  // moment later without a word.
  editInFlight: "an edit is still being applied",
  // An edit that has been typed and not committed exists only in an input
  // element. Nothing has been sent, so nothing else knows about it at all.
  editorOpen: "a cell is still open for editing",
};

/**
 * Everything at stake, in the order it is explained.
 *
 * Deliberately not held, and so never a reason: creating and opening, which
 * have nothing to lose yet; reading a file to describe it, which touches no
 * workspace; and saving, which is already covered because the unsaved flag
 * stays set until its write resolves. Holding those would put a warning in
 * front of a visitor with nothing at stake, which is how a warning stops being
 * read.
 */
export function workspaceHoldReasons(
  state: WorkspaceHoldState,
): readonly WorkspaceHoldReason[] {
  const reasons: WorkspaceHoldReason[] = [];
  if (state.unsavedChanges) {
    reasons.push("unsavedChanges");
  }
  if (state.importing) {
    reasons.push("importing");
  }
  if (state.editsInFlight > 0) {
    reasons.push("editInFlight");
  }
  // Last, because it is the least far along: the others are work the workspace
  // or the worker already has, and this is work that is still only on screen.
  if (state.editorOpen) {
    reasons.push("editorOpen");
  }
  return reasons;
}

export function mustHoldWorkspace(state: WorkspaceHoldState): boolean {
  return workspaceHoldReasons(state).length > 0;
}

/**
 * What to call what is at stake, in two words.
 *
 * Derived rather than fixed for the same reason the sentence is: a heading that
 * always read "Unsaved changes" said so over a workspace whose only reason for
 * being held was a command still on its way, where nothing has been changed
 * yet and the badge that means exactly that is not showing.
 */
export function workspaceHoldHeading(state: WorkspaceHoldState): string {
  return workspaceHoldHeadingFor(workspaceHoldReasons(state));
}

/**
 * The same heading, over reasons that are already in hand.
 *
 * A standing confirmation keeps the reasons it was raised about, so that a
 * teardown which takes the bookkeeping away cannot leave the question with
 * nothing to say. See `questionView` in `workspace-state`.
 */
export function workspaceHoldHeadingFor(
  reasons: readonly WorkspaceHoldReason[],
): string {
  return reasons.includes("unsavedChanges")
    ? "Unsaved changes"
    : "Work not saved yet";
}

/**
 * What the visitor is being asked about, as one sentence, or the empty string
 * when there is nothing to ask about. Every reason that stands is named: an
 * unsaved workspace can have an edit on its way as well, and a warning that
 * mentioned one of them would be describing half of what is at stake.
 */
export function workspaceHoldSentence(state: WorkspaceHoldState): string {
  return workspaceHoldSentenceFor(workspaceHoldReasons(state));
}

/** The same sentence, over reasons that are already in hand. */
export function workspaceHoldSentenceFor(
  reasons: readonly WorkspaceHoldReason[],
): string {
  const clauses = reasons.map((reason) => REASON_CLAUSES[reason]);
  const [first, ...rest] = clauses;
  if (first === undefined) {
    return "";
  }
  const joined =
    rest.length === 0
      ? first
      : rest.length === 1
        ? `${first} and ${rest[0] as string}`
        : `${[first, ...rest.slice(0, -1)].join(", ")}, and ${rest[rest.length - 1] as string}`;
  return `${joined.charAt(0).toUpperCase()}${joined.slice(1)}.`;
}
