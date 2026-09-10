"use client";

/**
 * The browser front end for a local data workspace.
 *
 * This is the shell only: it starts an empty workspace, opens an existing
 * `.sqlite` file, and saves the workspace back out. The database itself lives in
 * the workspace Web Worker (`workers/workspace.worker.ts`), which owns the one
 * `@consultchimps/db` instance; this component holds view state and drives the
 * worker through `WorkspaceClient`.
 *
 * Data import lives in `workspace-import.tsx` and the record grid in
 * `workspace-grid.tsx`, both mounted below the summary. Neither keeps state the
 * shell also keeps: the grid is handed the summary it may show tables from, the
 * busy flag that locks editing, and `markChanged` to call when an edit lands.
 * The table listing in the summary stays, as the compact description of what a
 * workspace holds: each table's name, row count, Record ID prefix, and the
 * column types import inferred.
 *
 * The workspace is one in-memory database, so replacing it or leaving the page
 * is the whole of losing it. The shell therefore owns the unsaved-changes flag
 * rather than each mutating feature: `markChanged` is the single place it is
 * set, and a write that actually resolved is the only place it is cleared. A
 * feature that changes the workspace inherits the guard by calling that one
 * marker.
 *
 * Leaving is guarded in every form it takes, not just the one the browser fires
 * an event for. Replacing the workspace (New, Open), following a link out of
 * the page, going back, and closing the tab all end the same way, so all four
 * are held by the same flag and all but the last ask through the same inline
 * confirmation. A guard attached to `beforeunload` alone would miss a
 * client-side transition entirely, because that never unloads anything: it just
 * unmounts this component, and the cleanup below then terminates the worker.
 *
 * Saving prefers the File System Access API so a repeat save writes back to the
 * same file in place. Where that API is missing, saving falls back to a plain
 * download, which is why the page never promises an in-place save it cannot
 * deliver. See `tool-kit.tsx` for the shell and the byte saver reused here.
 */

import {
  describeFailure,
  primaryButtonClass,
  readUploads,
  saveBinaryFile,
  secondaryButtonClass,
  sectionClass,
  ToolShell,
} from "@/components/tool-kit";
import { WorkspaceGrid } from "@/components/workspace-grid";
import { WorkspaceImport } from "@/components/workspace-import";
import { WORKSPACE_FILES } from "@/lib/accepted-files";
import {
  mustHoldWorkspace,
  workspaceHoldHeading,
  workspaceHoldSentence,
} from "@/lib/workspace-hold";
import type { WorkspaceSummary } from "@/lib/workspace-protocol";
import { WorkspaceClient } from "@/lib/workspace-worker";
import {
  Database,
  Download,
  FilePlus,
  FolderOpen,
  LoaderCircle,
  Save,
  TriangleAlert,
} from "lucide-react";
import { useRouter } from "next/navigation";
import { useCallback, useEffect, useRef, useState } from "react";

/** The media type and default name a saved workspace carries. */
const WORKSPACE_MEDIA_TYPE = "application/vnd.sqlite3";
const DEFAULT_WORKSPACE_NAME = "workspace.sqlite";

/**
 * The minimal File System Access API surface this page uses, declared locally so
 * the page does not depend on the DOM lib shipping these still-evolving types. A
 * handle grants in-place writes to one file; the pickers hand one back.
 */
interface WorkspaceWritable {
  write(data: Uint8Array): Promise<void>;
  close(): Promise<void>;
}
interface WorkspaceFileHandle {
  readonly name: string;
  getFile(): Promise<File>;
  createWritable(): Promise<WorkspaceWritable>;
}
interface FilePickerAcceptType {
  readonly description?: string;
  readonly accept: Record<string, string[]>;
}
interface FileSystemWindow {
  showOpenFilePicker?: (options?: {
    readonly multiple?: boolean;
    readonly types?: FilePickerAcceptType[];
  }) => Promise<WorkspaceFileHandle[]>;
  showSaveFilePicker?: (options?: {
    readonly suggestedName?: string;
    readonly types?: FilePickerAcceptType[];
  }) => Promise<WorkspaceFileHandle>;
}

const PICKER_TYPES: FilePickerAcceptType[] = [
  {
    description: "SQLite workspace",
    accept: { [WORKSPACE_MEDIA_TYPE]: [...WORKSPACE_FILES.extensions] },
  },
];

function fileSystemWindow(): FileSystemWindow {
  return window as unknown as FileSystemWindow;
}

/** A cancelled picker or writable rejects with an AbortError; that is not a failure. */
function isAbort(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    (error as { name?: unknown }).name === "AbortError"
  );
}

async function writeToHandle(
  handle: WorkspaceFileHandle,
  bytes: Uint8Array,
): Promise<void> {
  const writable = await handle.createWritable();
  try {
    await writable.write(bytes);
  } finally {
    await writable.close();
  }
}

/** What the page shows once a workspace is held: where it came from and its shape. */
interface OpenWorkspace {
  readonly summary: WorkspaceSummary;
  /** The file it was opened from or last saved as, or null for a fresh one. */
  readonly fileName: string | null;
  /**
   * Whether the workspace has changed since it was last written to a file.
   *
   * The workspace is one in-memory database and nothing else: replacing it, or
   * leaving the page, is the whole of losing it. So the shell tracks this
   * itself rather than leaving each mutating feature to remember, which is what
   * keeps the guard true for a mutation the shell has never heard of.
   */
  readonly unsavedChanges: boolean;
}

/**
 * What a confirmation is standing in front of, or null for none.
 *
 * Every way of losing the workspace routes through this one question rather
 * than growing a dialog each: replacing it, and leaving the page for somewhere
 * else in the app. `href` is where the visitor was going, or null when they
 * pressed Back and the way to honour that is to go back again.
 */
type PendingAction =
  | { readonly kind: "new" }
  | { readonly kind: "open" }
  | { readonly kind: "leave"; readonly href: string | null }
  | null;

/**
 * The shell's record-a-change callback, handed to any section that mutates the
 * workspace. Import receives it folded into `onImported`; the record grid takes
 * it directly when its cell edits land. Passing a summary replaces the table
 * listing at the same time, and passing nothing marks the workspace changed
 * without one.
 */
export type MarkWorkspaceChanged = (summary?: WorkspaceSummary) => void;

/**
 * Report how many of a feature's mutating commands are in flight, so the shell
 * can hold for them exactly as it holds for an import.
 *
 * A command that has been sent and not answered is work this tab holds and no
 * file does, even though nothing is marked unsaved yet: marking before the
 * answer would claim a change the worker may refuse. The count is what covers
 * the gap between the two.
 */
export type ReportPendingEdits = (count: number) => void;

/**
 * Report that a cell is open for editing, from the moment it opens until it
 * commits or cancels.
 *
 * A draft in an input element is work no one else knows about: nothing has been
 * sent, so nothing has an answer to wait for. The shell holds for it so that
 * the ways of leaving that never blur an editor, the Back button and closing
 * the tab, are not the ways that lose it.
 */
export type ReportOpenEditor = (open: boolean) => void;

/**
 * The one command in flight, or null while the page is idle.
 *
 * There is a single busy state for the whole page, not one per section, because
 * the worker runs one command at a time and every long-running command has the
 * same consequence: nothing else may start, and in particular nothing may
 * replace the workspace. A section that runs its own command reports through
 * `onBusy` rather than keeping a flag of its own, so the shell can always see
 * that something is in flight. A second notion of busy is exactly how a click
 * on New lands behind a running import and throws its result away.
 */
export type WorkspaceBusy =
  "creating" | "opening" | "saving" | "reading" | "importing" | null;

export function WorkspaceTool() {
  const clientRef = useRef<WorkspaceClient | null>(null);
  // The handle for an in-place save, held only when a picker granted one.
  const handleRef = useRef<WorkspaceFileHandle | null>(null);
  const fallbackInputRef = useRef<HTMLInputElement | null>(null);
  // How many spare history entries the page is holding for the Back guard, and
  // the condition that guard reads. Both are refs because the popstate listener
  // is installed once and has to see the current answer, not the first one.
  const spareEntriesRef = useRef(0);
  const mustHoldRef = useRef(false);
  /**
   * The same count as `editsInFlight`, kept where a handler can read it in the
   * tick it changed.
   *
   * The grid commits a cell edit when its editor loses focus, so clicking New
   * or Open is itself what commits the edit being replaced: the edit is sent
   * and the click handler runs in the same tick, and state read during that
   * handler would still be the state from before the edit. React catches up a
   * microtask later, which is soon enough to render and far too late to decide.
   */
  const editsInFlightRef = useRef(0);

  const [workspace, setWorkspace] = useState<OpenWorkspace | null>(null);
  const [busy, setBusy] = useState<WorkspaceBusy>(null);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [pending, setPending] = useState<PendingAction>(null);
  const [editsInFlight, setEditsInFlight] = useState(0);
  const [editorOpen, setEditorOpen] = useState(false);
  const router = useRouter();

  const client = useCallback((): WorkspaceClient => {
    clientRef.current ??= new WorkspaceClient();
    return clientRef.current;
  }, []);

  // Tear the worker down with the page so its database and wasm are released.
  useEffect(
    () => () => {
      clientRef.current?.terminate();
      clientRef.current = null;
    },
    [],
  );

  const startNew = useCallback(async () => {
    setBusy("creating");
    setError(null);
    setNotice(null);
    try {
      const summary = await client().create();
      handleRef.current = null;
      setWorkspace({ summary, fileName: null, unsavedChanges: false });
      setNotice("Started a new empty workspace");
    } catch (caught) {
      setError(describeFailure(caught));
    } finally {
      setBusy(null);
    }
  }, [client]);

  // Shared open path for every way a file arrives. The read itself runs inside
  // the guarded section, so a file that stops being readable after it was
  // chosen (moved, deleted, locked) reports through the same error path as a
  // database the worker rejects, whichever entry point chose it. The handle is
  // remembered when a picker supplied one so a later save writes back in place.
  const openWorkspace = useCallback(
    async (
      read: () => Promise<{
        readonly name: string;
        readonly bytes: Uint8Array;
      }>,
      handle: WorkspaceFileHandle | null,
    ) => {
      setBusy("opening");
      setError(null);
      setNotice(null);
      try {
        const { name, bytes } = await read();
        const summary = await client().open(bytes);
        handleRef.current = handle;
        setWorkspace({ summary, fileName: name, unsavedChanges: false });
        setNotice("Opened the workspace");
      } catch (caught) {
        setError(describeFailure(caught));
      } finally {
        setBusy(null);
      }
    },
    [client],
  );

  // The work an Open does once it is allowed to. The hidden file input below is
  // reachable only from here, so guarding this entry point guards every way a
  // visitor can replace the workspace with a file.
  const startOpen = useCallback(async () => {
    const picker = fileSystemWindow().showOpenFilePicker;
    if (picker === undefined) {
      // No File System Access API: fall back to the file input, whose change
      // handler continues the open.
      fallbackInputRef.current?.click();
      return;
    }
    let handle: WorkspaceFileHandle | undefined;
    try {
      [handle] = await picker({ multiple: false, types: PICKER_TYPES });
    } catch (caught) {
      // The visitor closed the picker: leave the page exactly as it was.
      if (isAbort(caught)) {
        return;
      }
      setError(describeFailure(caught));
      return;
    }
    if (handle === undefined) {
      return;
    }
    const chosen = handle;
    // A picker grants a writable handle, so a later Save writes back in place.
    await openWorkspace(async () => {
      const file = await chosen.getFile();
      return {
        name: file.name,
        bytes: new Uint8Array(await file.arrayBuffer()),
      };
    }, chosen);
  }, [openWorkspace]);

  const onFallbackFiles = useCallback(
    (files: readonly File[]) => {
      // A file input alone grants no write handle, so a later Save downloads a
      // copy, or offers Save as where the browser supports it.
      void openWorkspace(async () => {
        const [first] = await readUploads(files, WORKSPACE_FILES.accepts);
        if (first === undefined) {
          throw new Error(
            `That file is not ${WORKSPACE_FILES.description}, so nothing was opened`,
          );
        }
        return first;
      }, null);
    },
    [openWorkspace],
  );

  const hasUnsavedChanges = workspace?.unsavedChanges === true;

  // Whether there is anything to lose by leaving or replacing the workspace.
  // The conditions and the reasoning behind each live in `lib/workspace-hold`,
  // so every guard here asks one question with one answer.
  const holdState = {
    unsavedChanges: hasUnsavedChanges,
    importing: busy === "importing",
    editsInFlight,
    editorOpen,
  };
  const mustHold = mustHoldWorkspace(holdState);

  /**
   * The same question, asked from a handler that may be running in the same
   * tick as the edit that changed the answer. Everything but an in-flight edit
   * settles through a render long before a click reaches a button, so the last
   * rendered answer covers those, and the count covers the one that does not.
   *
   * Measured, not assumed: the editor commits on blur, and blur runs on the
   * mousedown that precedes the click, so React does in fact re-render in time
   * and reading state here would pass the tests below today. It reads the count
   * anyway, because that is a fact about React's scheduling rather than about
   * this page, and the cost of being wrong about it is a silently discarded
   * edit.
   *
   * An open editor needs no such treatment. Opening one is its own click, so a
   * render always separates it from the click that navigates, and the closing
   * happens on the blur that click causes: being a moment late to release a
   * hold only ever holds, which is the safe direction, and the sentence shown
   * is derived from the rendered state and so still describes what is true.
   */
  const holdsNow = useCallback(
    (): boolean => mustHoldRef.current || editsInFlightRef.current > 0,
    [],
  );

  /**
   * Count a feature's mutating commands. Recorded in the ref first, because the
   * click that raises the guard can be the same click that commits the edit.
   */
  const reportPendingEdits = useCallback<ReportPendingEdits>((count) => {
    editsInFlightRef.current = count;
    setEditsInFlight(count);
  }, []);

  /**
   * Count a cell open for editing. State alone, no ref: this has to reach the
   * rendered hold state, because the guard it exists for is the `beforeunload`
   * listener below, which can only be installed while the editor is open. A
   * listener that is decided on at unload time is one that was never there.
   */
  const reportOpenEditor = useCallback<ReportOpenEditor>((open) => {
    setEditorOpen(open);
  }, []);

  // The question exists only while its reason does. A save made while it is on
  // screen, or an import that fails after a link was held, answers it by
  // removing what it was about; left standing it would reappear at the next
  // change, asking about something that already happened. Adjusted during
  // render, which is React's own pattern for state that follows another value:
  // an effect would show the stale question for a frame first.
  const [heldFor, setHeldFor] = useState(mustHold);
  if (heldFor !== mustHold) {
    setHeldFor(mustHold);
    if (!mustHold && pending !== null) {
      setPending(null);
    }
  }

  // Both entry points that replace the held workspace go through here, so a
  // mutating feature added later inherits the guard by doing nothing.
  const replaceWorkspace = useCallback(
    (kind: "new" | "open") => {
      if (holdsNow()) {
        setPending({ kind });
        return;
      }
      void (kind === "new" ? startNew() : startOpen());
    },
    [holdsNow, startNew, startOpen],
  );

  /**
   * Do the thing the visitor has just accepted losing the workspace for.
   *
   * Leaving clears the flag first, because the visitor has answered the
   * question: keeping it set would arm `beforeunload` and ask them a second
   * time, in the browser's own words, for the navigation they just approved.
   */
  /**
   * Leave this page, retiring the spare history entries the Back guard pushed.
   *
   * Every way of leaving goes through here, because retiring them on only one
   * way out is how they became phantoms: a link followed forward pushed the
   * destination on top of the spare, so Back reached a second copy of this page
   * before reaching the page before it. Taking the spare's place instead of
   * stacking on it leaves exactly one entry for this page behind, whichever way
   * it was left and whether or not the guard had ever armed.
   *
   * `href` is null when the visitor pressed Back, which is honoured by stepping
   * back over the spares and this page together.
   */
  const leaveFor = useCallback(
    (href: string | null) => {
      const spares = spareEntriesRef.current;
      spareEntriesRef.current = 0;
      if (href === null) {
        window.history.go(-(spares + 1));
        return;
      }
      // The guard holds at most one spare, pushed only when none is held and
      // re-armed in the same breath by a Back that spends one, and the page is
      // sitting on it: nothing here pushes an entry of its own, and a jump
      // within this same page is left to the router untouched. So replacing the
      // current entry retires the spare exactly.
      if (spares > 0) {
        router.replace(href);
        return;
      }
      router.push(href);
    },
    [router],
  );

  const onConfirmAction = useCallback(() => {
    const action = pending;
    setPending(null);
    if (action === null) {
      return;
    }
    if (action.kind === "new") {
      void startNew();
      return;
    }
    if (action.kind === "open") {
      void startOpen();
      return;
    }
    setWorkspace((previous) =>
      previous === null ? previous : { ...previous, unsavedChanges: false },
    );
    leaveFor(action.href);
  }, [leaveFor, pending, startNew, startOpen]);

  /**
   * Record that the held workspace no longer matches its file.
   *
   * Every command that changes the workspace reports through here: import calls
   * it below with the summary it got back, and the record grid's cell edits call
   * it the same way when they land, with or without a summary. One flag, one
   * place that sets it, so a second mutating feature cannot arrive with a second
   * idea of what unsaved means.
   */
  const markChanged = useCallback<MarkWorkspaceChanged>((summary) => {
    setWorkspace((previous) =>
      previous === null
        ? previous
        : {
            ...previous,
            summary: summary ?? previous.summary,
            unsavedChanges: true,
          },
    );
  }, []);

  // Warn before the tab closes, reloads, or leaves for another site. This is
  // the browser's own dialog and the only guard available for those, but it
  // covers none of the ways of leaving that stay inside the app.
  //
  // Attached on the rendered answer rather than read from `holdsNow` at fire
  // time, which is deliberate. Unload cannot be reached in the same tick as a
  // cell edit, because closing or reloading is a gesture on the browser's own
  // chrome rather than a click on this page, and any later task sees the count
  // already rendered. Keeping the listener off a clean workspace is worth more
  // than covering a case that cannot happen: a page that always carries one
  // gives up the browser's back-forward cache.
  useEffect(() => {
    if (!mustHold) {
      return;
    }
    const warn = (event: BeforeUnloadEvent): void => {
      // Browsers show their own wording; both spellings of "yes, warn" are set
      // because they disagree about which one they honour.
      event.preventDefault();
      event.returnValue = "";
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [mustHold]);

  /**
   * Hold a link out of the page until the visitor has answered for the
   * workspace.
   *
   * A client-side transition unloads nothing, so `beforeunload` never fires; it
   * unmounts this component, and the cleanup above then terminates the worker
   * and the database with it. Catching the click in the capture phase is what
   * makes the guard run before any of that can start: the router's own handler
   * never sees the event, so there is nothing to undo afterward. Only a plain
   * left click on a same-origin link that actually leaves this page is held; a
   * modified click, a new tab, a download, and a jump to an anchor on this page
   * are all left alone, because none of them lose the workspace.
   *
   * The listener lives for the whole page rather than only while the guard is
   * armed, for the reason the Back observer does: a link followed after a save
   * still has to retire the spare entry that change left in the history, and a
   * listener that was not there cannot.
   */
  useEffect(() => {
    const hold = (event: MouseEvent): void => {
      if (
        event.defaultPrevented ||
        event.button !== 0 ||
        event.metaKey ||
        event.ctrlKey ||
        event.shiftKey ||
        event.altKey
      ) {
        return;
      }
      const anchor = (event.target as Element | null)?.closest?.(
        "a[href]",
      ) as HTMLAnchorElement | null;
      if (
        !anchor ||
        anchor.hasAttribute("download") ||
        (anchor.target !== "" && anchor.target !== "_self")
      ) {
        return;
      }
      const destination = new URL(anchor.href, window.location.href);
      if (destination.origin !== window.location.origin) {
        return;
      }
      if (
        destination.pathname === window.location.pathname &&
        destination.search === window.location.search
      ) {
        return;
      }
      if (holdsNow()) {
        event.preventDefault();
        event.stopPropagation();
        setPending({ kind: "leave", href: destination.href });
        return;
      }
      if (spareEntriesRef.current === 0) {
        // Nothing at stake and nothing to retire: the router's own handling is
        // exactly right, so it is left alone.
        return;
      }
      // Nothing at stake, but a spare from an earlier change is still in the
      // history. A save clears the flag and cannot remove that entry, so the
      // clean way out has to retire it or inherit the phantom.
      event.preventDefault();
      event.stopPropagation();
      leaveFor(destination.href);
    };
    document.addEventListener("click", hold, true);
    return () => document.removeEventListener("click", hold, true);
  }, [leaveFor]);

  /**
   * Hold the Back button the same way.
   *
   * Back cannot be cancelled once it has happened, so the only way to catch it
   * is to have somewhere harmless for it to land: a spare history entry for
   * this same page, which the first Back press consumes without changing the
   * route. The trade is one Back press that appears to do nothing, in a session
   * that had something at stake, against losing that work outright.
   *
   * Two things keep the bookkeeping honest. The listener is installed for the
   * life of the page rather than while the guard is armed, so a press that
   * happens between arming and re-arming is still counted: what the page
   * believes about history comes from what happened to history, never from what
   * the guard was doing at the time. And it counts entries rather than holding a
   * flag, so a press while the question is already showing is caught too, and
   * answering it steps back over exactly the spares that were pushed.
   */
  useEffect(() => {
    const observed = (): void => {
      if (spareEntriesRef.current === 0) {
        // Not one of ours: the visitor is leaving a page we never armed.
        return;
      }
      spareEntriesRef.current -= 1;
      if (!holdsNow()) {
        // Nothing at stake, so the press was simply spent. Whether the page
        // stays or goes from here is the browser's business.
        return;
      }
      // Re-arm before asking, so a second press while the question is up lands
      // somewhere harmless too.
      window.history.pushState(window.history.state, "", window.location.href);
      spareEntriesRef.current += 1;
      setPending({ kind: "leave", href: null });
    };
    window.addEventListener("popstate", observed);
    return () => window.removeEventListener("popstate", observed);
  }, []);

  // The observer above outlives every render, so it reads the condition from a
  // ref rather than closing over a stale copy of it.
  useEffect(() => {
    mustHoldRef.current = mustHold;
  }, [mustHold]);

  // Arm whenever there is something to lose and no spare entry is held, which
  // covers both the first change and a change made after an earlier spare was
  // spent.
  useEffect(() => {
    if (!mustHold || spareEntriesRef.current > 0) {
      return;
    }
    // Nothing below this page means Back cannot leave it, so there is nothing
    // to guard against. Arming anyway would make a dead button look live, and
    // then offer to leave for somewhere that does not exist.
    if (window.history.length <= 1) {
      return;
    }
    // Next's router keeps its own state on the entry, so the copy carries it
    // rather than a null that the router would not recognise on the way back.
    window.history.pushState(window.history.state, "", window.location.href);
    spareEntriesRef.current += 1;
  }, [mustHold]);

  // Serialize once, then route the bytes to the right destination.
  const saveWith = useCallback(
    async (mode: "save" | "saveAs") => {
      setBusy("saving");
      setError(null);
      setNotice(null);
      try {
        const bytes = await client().serialize();
        const existing = handleRef.current;
        const saveAs = fileSystemWindow().showSaveFilePicker;

        // Save writes back in place when a handle is already held.
        if (mode === "save" && existing !== null) {
          await writeToHandle(existing, bytes);
          // Cleared here rather than beside the serialize above: the bytes only
          // reach the file once the write resolves, and a write that throws has
          // to leave the workspace unsaved.
          setWorkspace((previous) =>
            previous === null
              ? previous
              : { ...previous, unsavedChanges: false },
          );
          setNotice("Saved to the workspace file");
          return;
        }

        // Otherwise ask for a destination where the API is available: this both
        // handles Save as and gives a first Save somewhere to write.
        if (saveAs !== undefined) {
          let handle: WorkspaceFileHandle;
          try {
            handle = await saveAs({
              suggestedName: workspace?.fileName ?? DEFAULT_WORKSPACE_NAME,
              types: PICKER_TYPES,
            });
          } catch (caught) {
            if (isAbort(caught)) {
              return;
            }
            throw caught;
          }
          await writeToHandle(handle, bytes);
          handleRef.current = handle;
          setWorkspace((previous) =>
            previous === null
              ? previous
              : { ...previous, fileName: handle.name, unsavedChanges: false },
          );
          setNotice("Saved to the workspace file");
          return;
        }

        // No File System Access API anywhere: download a copy.
        saveBinaryFile(
          bytes,
          workspace?.fileName ?? DEFAULT_WORKSPACE_NAME,
          WORKSPACE_MEDIA_TYPE,
        );
        // A download hands the bytes to the browser and the page never learns
        // where they landed, so this is the strongest signal this surface
        // offers. Treating it as unsaved forever would make the guard fire on
        // every New in a browser without the File System Access API, which
        // teaches people to dismiss it.
        setWorkspace((previous) =>
          previous === null ? previous : { ...previous, unsavedChanges: false },
        );
        setNotice("Downloaded a copy of the workspace");
      } catch (caught) {
        setError(describeFailure(caught));
      } finally {
        setBusy(null);
      }
    },
    [client, workspace],
  );

  // Import replaces the summary wholesale, so the table listing above always
  // reflects what the worker now holds rather than a count kept in step by hand,
  // and it marks the workspace changed through the shared marker.
  const onImported = useCallback(
    (summary: WorkspaceSummary, imported: string) => {
      markChanged(summary);
      setError(null);
      setNotice(imported);
    },
    [markChanged],
  );

  const isBusy = busy !== null;
  const hasWorkspace = workspace !== null;
  // Derived rather than stored, so a save made while the question is on screen
  // answers it: the reason to ask is gone, so the asking goes with it.
  // `pending` is cleared the moment its reason goes, so holding one is the
  // whole condition rather than half of it.
  const confirming = pending !== null;

  return (
    <ToolShell
      description="Start a workspace in this tab, or open one you saved before, import a worksheet or a .csv file into it, then save it back to a single file. The workspace is an in-memory database that never leaves your browser"
      guideHref="/docs/libraries#build-a-local-database"
      guideLabel="Read about the local database"
      kicker="Online tool · Data workspace"
      title="Data workspace"
    >
      <section className={sectionClass} data-testid="workspace-actions">
        <h2 className="text-xl font-bold tracking-[-0.03em]">
          Start or open a workspace
        </h2>
        <p className="mt-3 text-sm text-fd-muted-foreground">
          A new workspace is empty. Saving writes back to the same file where
          your browser supports it, and downloads a copy everywhere else.
          Starting or opening another workspace replaces the one in this tab, so
          changes that have not been saved are confirmed first
        </p>
        <div className="mt-5 flex flex-wrap gap-3">
          <button
            className={primaryButtonClass}
            data-testid="workspace-new"
            disabled={isBusy}
            onClick={() => replaceWorkspace("new")}
            type="button"
          >
            {busy === "creating" ? (
              <LoaderCircle
                aria-hidden="true"
                className="size-4 animate-spin"
              />
            ) : (
              <FilePlus aria-hidden="true" className="size-4" />
            )}
            New workspace
          </button>
          <button
            className={secondaryButtonClass}
            data-testid="workspace-open"
            disabled={isBusy}
            onClick={() => replaceWorkspace("open")}
            type="button"
          >
            {busy === "opening" ? (
              <LoaderCircle
                aria-hidden="true"
                className="size-4 animate-spin"
              />
            ) : (
              <FolderOpen aria-hidden="true" className="size-4" />
            )}
            Open a workspace file
          </button>
          {/* The fallback picker for browsers without the File System Access
              API. Hidden, but Playwright and assistive tech reach it by test id
              and label. */}
          <input
            accept={WORKSPACE_FILES.accept}
            aria-label="Open a workspace file"
            className="hidden"
            data-testid="file-input"
            onChange={(event) => {
              const files = [...(event.target.files ?? [])];
              event.target.value = "";
              onFallbackFiles(files);
            }}
            ref={fallbackInputRef}
            type="file"
          />
        </div>
      </section>

      {confirming ? (
        <section
          aria-live="assertive"
          className={`${sectionClass} border-fd-primary/60`}
          data-testid="workspace-confirm"
        >
          <div className="flex items-center gap-2">
            <TriangleAlert
              aria-hidden="true"
              className="size-5 shrink-0 text-fd-primary"
            />
            <h2 className="text-xl font-bold tracking-[-0.03em]">
              {workspaceHoldHeading(holdState)}
            </h2>
          </div>
          <p className="mt-3 text-sm text-fd-muted-foreground">
            {/* The same state the hold decision was made from, so the
                sentence cannot describe a different one. */}
            {workspaceHoldSentence(holdState)}{" "}
            {pending?.kind === "new"
              ? "Starting a new workspace replaces this one"
              : pending?.kind === "open"
                ? "Opening another workspace replaces this one"
                : "Leaving this page closes the workspace"}
            , and that work is gone
          </p>
          <div className="mt-5 flex flex-wrap gap-3">
            <button
              className={secondaryButtonClass}
              data-testid="workspace-confirm-discard"
              disabled={isBusy}
              onClick={onConfirmAction}
              type="button"
            >
              {pending?.kind === "leave"
                ? "Discard the changes and leave"
                : "Discard the changes and continue"}
            </button>
            <button
              className={primaryButtonClass}
              data-testid="workspace-confirm-cancel"
              disabled={isBusy}
              onClick={() => setPending(null)}
              type="button"
            >
              Keep this workspace
            </button>
          </div>
        </section>
      ) : null}

      {hasWorkspace ? (
        <section className={sectionClass} data-testid="workspace-summary">
          <div className="flex items-center gap-2">
            <Database
              aria-hidden="true"
              className="size-5 shrink-0 text-fd-primary"
            />
            <h2 className="text-xl font-bold tracking-[-0.03em]">
              Workspace open
            </h2>
            {workspace.unsavedChanges ? (
              <span
                className="rounded-full border border-fd-primary/50 bg-fd-accent/40 px-2.5 py-0.5 font-mono text-xs font-semibold uppercase tracking-[0.12em] text-fd-accent-foreground"
                data-testid="workspace-unsaved"
              >
                Unsaved changes
              </span>
            ) : null}
          </div>
          <dl className="mt-4 grid grid-cols-2 gap-4 sm:grid-cols-3">
            <div>
              <dt className="font-mono text-xs uppercase tracking-[0.12em] text-fd-muted-foreground">
                Source
              </dt>
              <dd
                className="mt-1 truncate font-mono text-sm"
                data-testid="workspace-file-name"
              >
                {workspace.fileName ?? "New workspace"}
              </dd>
            </div>
            <div>
              <dt className="font-mono text-xs uppercase tracking-[0.12em] text-fd-muted-foreground">
                Tables
              </dt>
              <dd
                className="mt-1 text-2xl font-bold"
                data-testid="workspace-table-count"
              >
                {workspace.summary.tableCount}
              </dd>
            </div>
            <div>
              <dt className="font-mono text-xs uppercase tracking-[0.12em] text-fd-muted-foreground">
                Schema version
              </dt>
              <dd className="mt-1 text-2xl font-bold">
                {workspace.summary.schemaFormatVersion}
              </dd>
            </div>
          </dl>

          {workspace.summary.tables.length > 0 ? (
            <ul className="mt-6 space-y-3" data-testid="workspace-tables">
              {workspace.summary.tables.map((table) => (
                <li
                  className="rounded-lg border bg-fd-background/60 px-4 py-3"
                  data-testid="workspace-table"
                  key={table.name}
                >
                  <div className="flex flex-wrap items-baseline gap-x-3 gap-y-1">
                    <span
                      className="font-mono text-sm font-semibold"
                      data-testid="workspace-table-name"
                    >
                      {table.name}
                    </span>
                    <span
                      className="text-sm text-fd-muted-foreground"
                      data-testid="workspace-table-rows"
                    >
                      {table.rowCount === 1
                        ? "1 row"
                        : `${table.rowCount} rows`}
                    </span>
                    <span
                      className="font-mono text-xs text-fd-muted-foreground"
                      data-testid="workspace-table-prefix"
                    >
                      {table.recordIdPrefix}
                    </span>
                  </div>
                  <p
                    className="mt-1 font-mono text-xs text-fd-muted-foreground"
                    data-testid="workspace-table-columns"
                  >
                    {table.columns
                      .map((column) => `${column.name} (${column.type})`)
                      .join(", ")}
                  </p>
                </li>
              ))}
            </ul>
          ) : (
            <p
              className="mt-6 text-sm text-fd-muted-foreground"
              data-testid="workspace-tables-empty"
            >
              This workspace holds no tables yet. Import a worksheet or a .csv
              file to add one
            </p>
          )}

          <div className="mt-6 flex flex-wrap gap-3">
            <button
              className={primaryButtonClass}
              data-testid="workspace-save"
              disabled={isBusy}
              onClick={() => void saveWith("save")}
              type="button"
            >
              {busy === "saving" ? (
                <LoaderCircle
                  aria-hidden="true"
                  className="size-4 animate-spin"
                />
              ) : (
                <Save aria-hidden="true" className="size-4" />
              )}
              Save
            </button>
            <button
              className={secondaryButtonClass}
              data-testid="workspace-save-as"
              disabled={isBusy}
              onClick={() => void saveWith("saveAs")}
              type="button"
            >
              <Download aria-hidden="true" className="size-4" />
              Save as
            </button>
          </div>
        </section>
      ) : (
        <section className={sectionClass} data-testid="workspace-empty">
          <p className="text-sm text-fd-muted-foreground">
            No workspace is open yet. Start a new one or open a saved file to
            begin
          </p>
        </section>
      )}

      {/* Import mounts as its own section so the two files stay independent:
          everything the flow needs lives in workspace-import.tsx and only the
          new workspace summary comes back here. */}
      {hasWorkspace ? (
        <WorkspaceImport
          busy={busy}
          client={client}
          existingTableNames={workspace.summary.tables.map(
            (table) => table.name,
          )}
          onBusy={setBusy}
          onImported={onImported}
        />
      ) : null}

      {/* The grid keeps no flag the shell already keeps: which tables exist and
          which database they belong to are both read from the summary, editing
          is locked by the shell's own busy and confirming state, and an edit
          that lands reports it through the one marker. It does keep its own
          refusals, because those belong to the cells they name rather than to
          the page, and a slot shared with this component's own failures is one
          each would clear from under the other. */}
      {hasWorkspace ? (
        <WorkspaceGrid
          getClient={client}
          locked={isBusy || confirming}
          markChanged={markChanged}
          onEditsPending={reportPendingEdits}
          onEditorOpen={reportOpenEditor}
          summary={workspace.summary}
        />
      ) : null}

      {notice ? (
        <p
          aria-live="polite"
          className="rounded-lg border border-fd-primary/40 bg-fd-accent/30 px-4 py-3 text-sm text-fd-accent-foreground"
          data-testid="workspace-notice"
          role="status"
        >
          {notice}
        </p>
      ) : null}

      {error ? (
        <pre
          aria-live="polite"
          className="overflow-x-auto whitespace-pre-wrap rounded-lg border border-fd-primary/40 bg-fd-accent/30 px-4 py-3 text-xs leading-6 text-fd-accent-foreground"
          data-testid="workspace-error"
        >
          {error}
        </pre>
      ) : null}
    </ToolShell>
  );
}
