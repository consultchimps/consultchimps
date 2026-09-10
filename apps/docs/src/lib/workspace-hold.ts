/**
 * Whether there is anything to lose by leaving or replacing the workspace.
 *
 * Every guard the shell has asks this one question: New, Open, a link out of
 * the page, the Back button, and the browser's own unload warning. Any of them
 * keyed on part of the answer is a hole, so the answer lives here, once, and
 * they all read it.
 *
 * It is a function rather than an expression inside the component because the
 * conditions are what the guarding is, and each one arrived after a way of
 * losing work was found. Keeping them here means a new one can be added with
 * the case that motivated it, and every guard inherits it by doing nothing.
 */

/** What the shell knows about work the file does not hold. */
export interface WorkspaceHoldState {
  /** The workspace has changes no file has. */
  readonly unsavedChanges: boolean;
  /** An import has been sent to the worker and has not come back. */
  readonly importing: boolean;
  /** Cell edits sent to the worker that have not come back. */
  readonly editsInFlight: number;
}

export function mustHoldWorkspace(state: WorkspaceHoldState): boolean {
  // The obvious half: changes that exist only in this tab's memory.
  if (state.unsavedChanges) {
    return true;
  }
  // An import that has not come back is the one command whose result exists
  // nowhere else, so leaving mid-import destroys work that was never anywhere
  // but this tab.
  if (state.importing) {
    return true;
  }
  // A cell edit that has not come back is the same thing at a smaller scale,
  // and it is the harder one to see. The workspace still looks clean, because
  // an edit is only counted as unsaved once the worker has accepted it, which
  // is the only honest moment to count it. Between the two, a visitor who
  // commits an edit and reaches straight for New would be replacing a database
  // the edit is still on its way to, and the commands run in the order they
  // were sent: the edit would land, and the workspace holding it would be
  // discarded a moment later without a word. Holding here closes that window
  // without giving the grid a second idea of what unsaved means.
  return state.editsInFlight > 0;

  // Deliberately not held: creating and opening, which have nothing to lose
  // yet; reading a file to describe it, which touches no workspace; and saving,
  // which is already covered because the unsaved flag stays set until its write
  // resolves. Holding those would put a warning in front of a visitor with
  // nothing at stake, which is how a warning stops being read.
}
