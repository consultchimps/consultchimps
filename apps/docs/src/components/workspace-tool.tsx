"use client";

/**
 * The browser front end for a local data workspace.
 *
 * This is the shell only: it starts an empty workspace, opens an existing
 * `.sqlite` file, and saves the workspace back out. The database itself lives in
 * the workspace Web Worker (`workers/workspace.worker.ts`), which owns the one
 * `@consultchimps/db` instance; this component drives the worker through
 * `WorkspaceClient` and holds no view state of its own.
 *
 * What the page is doing lives in `lib/workspace-state`: one state, one total
 * transition table, and every answer derived from it. This file is the part
 * that cannot be pure. It turns clicks, worker replies and browser events into
 * events for that model, performs the effects the model asks for, and renders
 * the answers it derives. It decides nothing: a condition worked out here would
 * be a second reading of a state that already has one, and two readings drift.
 *
 * Data import lives in `workspace-import.tsx` and the record grid in
 * `workspace-grid.tsx`, both mounted below the summary. Neither keeps state the
 * shell also keeps: the grid is handed the summary it may show tables from and
 * the lock that closes its editors, and reports its edits and its own teardown
 * back as events. The table listing in the summary stays, as the compact
 * description of what a workspace holds: each table's name, row count, Record
 * ID prefix, and the column types import inferred.
 *
 * Leaving is guarded in every form it takes, not just the one the browser fires
 * an event for. Replacing the workspace (New, Open), following a link out of
 * the page, going back, and closing the tab all end the same way, so all four
 * ask the model the one question and all but the last ask the visitor through
 * the same inline confirmation. A guard attached to `beforeunload` alone would
 * miss a client-side transition entirely, because that never unloads anything:
 * it just unmounts this component, and the cleanup below then terminates the
 * worker.
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
import type { WorkspaceSummary } from "@/lib/workspace-protocol";
import {
  commandsEnabled,
  editingLocked,
  INITIAL_WORKSPACE_STATE,
  questionView,
  shouldArmHistorySpare,
  spinningButton,
  unloadGuardInstalled,
  workspaceBusy,
  workspaceStep,
  type SaveMode,
  type WorkspaceEffect,
  type WorkspaceEvent,
  type WorkspaceState,
} from "@/lib/workspace-state";
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

/**
 * An edit is on its way to the worker.
 *
 * A command that has been sent and not answered is work this tab holds and no
 * file does, even though nothing is marked unsaved yet: marking before the
 * answer would claim a change the worker may refuse. Reporting it is what
 * covers the gap between the two.
 */
export type ReportEditSent = () => void;

/**
 * That edit came back, and whether the worker took it. One report rather than
 * two, so the workspace is marked unsaved and the edit's hold released in the
 * same move.
 */
export type ReportEditSettled = (accepted: boolean) => void;

/**
 * Report that a cell is open for editing, from the moment it opens until it
 * commits or cancels.
 *
 * A draft in an input element is work no one else knows about: nothing has been
 * sent, so nothing has an answer to wait for. The shell holds for it so that
 * the ways of leaving that never blur an editor, the Back button and closing
 * the tab, are not the ways that lose it.
 */
export type ReportEditorOpened = () => void;

/**
 * That editor closed, and whether the grid closed it on its way out. A
 * visitor's Escape answers a standing confirmation; a teardown may not, and
 * only the grid knows which of the two Tabulator's cancel was.
 */
export type ReportEditorClosed = (byTeardown: boolean) => void;

/**
 * Report that the grid has gone, taking whatever it held with it. Distinct from
 * an editor closing because the model has to tell a teardown from a visitor's
 * own Escape: see the #174 rule in `lib/workspace-state`.
 */
export type ReportGridDetached = () => void;

export function WorkspaceTool() {
  const clientRef = useRef<WorkspaceClient | null>(null);
  // The handle for an in-place save, held only when a picker granted one. It
  // stays a ref rather than joining the state: it is an opaque browser object,
  // nothing is derived from it, and what a save reached is reported back.
  const handleRef = useRef<WorkspaceFileHandle | null>(null);
  const fallbackInputRef = useRef<HTMLInputElement | null>(null);
  const router = useRouter();

  /**
   * The one state, and the same value where a handler must read it inside the
   * event that changed it.
   *
   * The ref is not a second copy: `send` is the only writer, and it moves both
   * in the one call, so the ref is the state as of the last event and the React
   * copy is the state as of the last render. The distinction matters exactly
   * once. The grid commits a cell edit when its editor loses focus, so clicking
   * New or Open is itself what commits the edit being replaced: the edit is
   * reported and the click handler runs in the same tick, and rendered state
   * read during that handler would still be the state from before the edit.
   */
  const [state, setState] = useState<WorkspaceState>(INITIAL_WORKSPACE_STATE);
  const stateRef = useRef<WorkspaceState>(INITIAL_WORKSPACE_STATE);

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

  /**
   * Move the state, and hand back what the model asked the page to do.
   *
   * It does not perform the effect: a handler that has to decide whether to
   * take a browser event over needs the answer before it acts on it, and one
   * function that both decides and acts could not give it one.
   */
  const send = useCallback((event: WorkspaceEvent): WorkspaceEffect | null => {
    const step = workspaceStep(stateRef.current, event);
    stateRef.current = step.state;
    setState(step.state);
    return step.effect;
  }, []);

  const runCreate = useCallback(async () => {
    try {
      const summary = await client().create();
      handleRef.current = null;
      send({ type: "createSucceeded", summary });
    } catch (caught) {
      send({ type: "createFailed", message: describeFailure(caught) });
    }
  }, [client, send]);

  // Shared open path for every way a file arrives. The read itself runs inside
  // the guarded section, so a file that stops being readable after it was
  // chosen (moved, deleted, locked) reports through the same error path as a
  // database the worker rejects, whichever entry point chose it. The handle is
  // remembered when a picker supplied one so a later save writes back in place.
  const runOpen = useCallback(
    async (
      read: () => Promise<{
        readonly name: string;
        readonly bytes: Uint8Array;
      }>,
      handle: WorkspaceFileHandle | null,
    ) => {
      // The model decides whether this open may start, and this honours its
      // answer: without it the worker would replace its database for a page
      // that had stopped listening, and the summary on screen would name a
      // workspace the worker no longer holds.
      //
      // The question asked is whether THIS event moved the state, not whether
      // the page happens to be opening, which a second open racing the first
      // would also see. An event the table refuses returns the state it was
      // given, by identity, so comparing the value before and after is exactly
      // that question. Nothing is said when it is refused: the command the page
      // is actually running has an outcome of its own, and that is what the
      // visitor is waiting to read.
      const before = stateRef.current;
      send({ type: "openStarted" });
      if (stateRef.current === before) {
        return;
      }
      try {
        const { name, bytes } = await read();
        const summary = await client().open(bytes);
        handleRef.current = handle;
        send({ type: "openSucceeded", summary, fileName: name });
      } catch (caught) {
        send({ type: "openFailed", message: describeFailure(caught) });
      }
    },
    [client, send],
  );

  // The work an Open does once the model has allowed it. The hidden file input
  // below is reachable only from here, so guarding the click that leads here
  // guards every way a visitor can replace the workspace with a file.
  const runOpenFile = useCallback(async () => {
    const picker = fileSystemWindow().showOpenFilePicker;
    if (picker === undefined) {
      // No File System Access API: fall back to the file input, whose change
      // handler continues the open. Nothing is in flight until it reports one,
      // because a file input that the visitor closes reports nothing at all.
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
      send({ type: "openPickerFailed", message: describeFailure(caught) });
      return;
    }
    if (handle === undefined) {
      return;
    }
    const chosen = handle;
    // A picker grants a writable handle, so a later Save writes back in place.
    await runOpen(async () => {
      const file = await chosen.getFile();
      return {
        name: file.name,
        bytes: new Uint8Array(await file.arrayBuffer()),
      };
    }, chosen);
  }, [runOpen, send]);

  const onFallbackFiles = useCallback(
    (files: readonly File[]) => {
      // A file input alone grants no write handle, so a later Save downloads a
      // copy, or offers Save as where the browser supports it.
      void runOpen(async () => {
        const [first] = await readUploads(files, WORKSPACE_FILES.accepts);
        if (first === undefined) {
          throw new Error(
            `That file is not ${WORKSPACE_FILES.description}, so nothing was opened`,
          );
        }
        return first;
      }, null);
    },
    [runOpen],
  );

  // Serialize once, then route the bytes to the right destination. Which of the
  // three it reaches cannot be known at the click, so the mode is what the
  // state carries and the destination is what comes back.
  const runSave = useCallback(
    async (how: SaveMode) => {
      try {
        const bytes = await client().serialize();
        const existing = handleRef.current;
        const saveAs = fileSystemWindow().showSaveFilePicker;
        const suggested =
          stateRef.current.held?.fileName ?? DEFAULT_WORKSPACE_NAME;

        // Save writes back in place when a handle is already held.
        if (how === "save" && existing !== null) {
          await writeToHandle(existing, bytes);
          // Reported here rather than beside the serialize above: the bytes
          // only reach the file once the write resolves, and a write that
          // throws has to leave the workspace unsaved.
          send({ type: "saveSucceeded", destination: "inPlace" });
          return;
        }

        // Otherwise ask for a destination where the API is available: this both
        // handles Save as and gives a first Save somewhere to write.
        if (saveAs !== undefined) {
          let handle: WorkspaceFileHandle;
          try {
            handle = await saveAs({
              suggestedName: suggested,
              types: PICKER_TYPES,
            });
          } catch (caught) {
            if (isAbort(caught)) {
              send({ type: "saveDismissed" });
              return;
            }
            throw caught;
          }
          await writeToHandle(handle, bytes);
          handleRef.current = handle;
          send({
            type: "saveSucceeded",
            destination: "saveAs",
            fileName: handle.name,
          });
          return;
        }

        // No File System Access API anywhere: download a copy.
        saveBinaryFile(bytes, suggested, WORKSPACE_MEDIA_TYPE);
        send({ type: "saveSucceeded", destination: "download" });
      } catch (caught) {
        send({ type: "saveFailed", message: describeFailure(caught) });
      }
    },
    [client, send],
  );

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
    (href: string | null, spares: number) => {
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

  /**
   * Push somewhere harmless for a Back press to land, and record that it is
   * there. What the page believes about history comes from what happened to
   * history, never from what the guard was doing at the time.
   */
  const armSpare = useCallback(() => {
    // Next's router keeps its own state on the entry, so the copy carries it
    // rather than a null that the router would not recognise on the way back.
    window.history.pushState(window.history.state, "", window.location.href);
    send({ type: "historySpareArmed" });
  }, [send]);

  /** Do what the model asked for, and nothing it did not. */
  const perform = useCallback(
    (effect: WorkspaceEffect | null) => {
      if (effect === null) {
        return;
      }
      switch (effect.kind) {
        case "create":
          void runCreate();
          return;
        case "openFile":
          void runOpenFile();
          return;
        case "save":
          void runSave(effect.how);
          return;
        case "leave":
          leaveFor(effect.href, effect.spares);
          return;
        case "armHistoryEntry":
          armSpare();
          return;
        case "holdNavigation":
          // Taking the click over is the caller's to do, inside the event.
          return;
        default: {
          // Not a runtime path: every effect above returns, and this is what
          // makes an effect added to the model without a home here a type
          // error rather than a click that quietly does nothing.
          const unhandled: never = effect;
          return unhandled;
        }
      }
    },
    [armSpare, leaveFor, runCreate, runOpenFile, runSave],
  );

  /** The common case: move the state and do what that asked for. */
  const dispatch = useCallback(
    (event: WorkspaceEvent) => {
      perform(send(event));
    },
    [perform, send],
  );

  // One stable reporter per thing a section can report.
  //
  // They go through `send` rather than `dispatch`, which matters twice. It is
  // what makes them depend on nothing (`send` is created once), and the grid
  // releases what it holds in an unmount cleanup keyed on `onDetached`: a
  // reporter whose identity moved would run that cleanup mid-life, releasing a
  // hold for work still on screen and rebuilding the grid under an open editor.
  // And it is honest about the model, which asks the page to do nothing for any
  // of these; a unit test holds that true, so an effect cannot be dropped here
  // by a later change without the test failing.
  const onReading = useCallback(
    () => send({ type: "importFileChosen" }),
    [send],
  );
  const onReadFinished = useCallback(
    () => send({ type: "importReadFinished" }),
    [send],
  );
  const onRunning = useCallback(() => send({ type: "importClicked" }), [send]);
  const onImported = useCallback(
    (summary: WorkspaceSummary, notice: string) =>
      send({ type: "importSucceeded", summary, notice }),
    [send],
  );
  const onImportFailed = useCallback(
    () => send({ type: "importFailed" }),
    [send],
  );
  const onEditorOpened = useCallback<ReportEditorOpened>(
    () => send({ type: "editorOpened" }),
    [send],
  );
  const onEditorClosed = useCallback<ReportEditorClosed>(
    (byTeardown) => send({ type: "editorClosed", byTeardown }),
    [send],
  );
  const onEditSent = useCallback<ReportEditSent>(
    () => send({ type: "editSent" }),
    [send],
  );
  const onEditSettled = useCallback<ReportEditSettled>(
    (accepted) => send({ type: "editSettled", accepted }),
    [send],
  );
  const onDetached = useCallback<ReportGridDetached>(
    () => send({ type: "gridDetached" }),
    [send],
  );

  const holding = unloadGuardInstalled(state);

  // Warn before the tab closes, reloads, or leaves for another site. This is
  // the browser's own dialog and the only guard available for those, but it
  // covers none of the ways of leaving that stay inside the app.
  //
  // Attached on the rendered answer rather than read at fire time, which is
  // deliberate. Unload cannot be reached in the same tick as a cell edit,
  // because closing or reloading is a gesture on the browser's own chrome
  // rather than a click on this page, and any later task sees the state already
  // rendered. Keeping the listener off a clean workspace is worth more than
  // covering a case that cannot happen: a page that always carries one gives up
  // the browser's back-forward cache.
  useEffect(() => {
    if (!holding) {
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
  }, [holding]);

  /**
   * Hold a link out of the page until the visitor has answered for the
   * workspace.
   *
   * A client-side transition unloads nothing, so `beforeunload` never fires; it
   * unmounts this component, and the cleanup above then terminates the worker
   * and the database with it. Catching the click in the capture phase is what
   * makes the guard run before any of that can start: the router's own handler
   * never sees the event, so there is nothing to undo afterward. Only a plain
   * left click on a same-origin link that actually leaves this page is offered
   * to the model; a modified click, a new tab, a download, and a jump to an
   * anchor on this page are all left alone, because none of them lose the
   * workspace.
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
      const effect = send({ type: "linkClicked", href: destination.href });
      if (effect === null) {
        // Nothing at stake and nothing to retire: the router's own handling is
        // exactly right, so it is left alone.
        return;
      }
      // The model took the click over, either to ask about it or to retire a
      // spare entry on the way out. Either way the browser must not follow it.
      event.preventDefault();
      event.stopPropagation();
      perform(effect);
    };
    document.addEventListener("click", hold, true);
    return () => document.removeEventListener("click", hold, true);
  }, [perform, send]);

  /**
   * Hold the Back button the same way.
   *
   * Back cannot be cancelled once it has happened, so the only way to catch it
   * is to have somewhere harmless for it to land: a spare history entry for
   * this same page, which the first Back press consumes without changing the
   * route. The trade is one Back press that appears to do nothing, in a session
   * that had something at stake, against losing that work outright. A press
   * with nothing at stake is not spent on nothing: the model retires what is
   * left and lets the navigation through.
   *
   * The listener is installed for the life of the page rather than while the
   * guard is armed, so a press that happens between arming and re-arming is
   * still counted: what the page believes about history comes from what
   * happened to history, never from what the guard was doing at the time.
   */
  useEffect(() => {
    const observed = (): void => {
      dispatch({ type: "backPressed" });
    };
    window.addEventListener("popstate", observed);
    return () => window.removeEventListener("popstate", observed);
  }, [dispatch]);

  /**
   * Notice the page coming back from the browser's back-forward cache.
   *
   * A confirmed leave puts the page into `leaving`, where every event is
   * ignored so that a worker reply arriving during the navigation cannot put
   * back the hold the visitor has just released. A tab restored by Forward
   * comes back with that state intact and would be a page that answers nothing,
   * so the restore is an event like any other and ends it.
   */
  useEffect(() => {
    const restored = (event: PageTransitionEvent): void => {
      if (event.persisted) {
        dispatch({ type: "pageRestored" });
      }
    };
    window.addEventListener("pageshow", restored);
    return () => window.removeEventListener("pageshow", restored);
  }, [dispatch]);

  // Arm whenever there is something to lose and no spare entry is held, which
  // covers both the first change and a change made after an earlier spare was
  // spent.
  const needsSpare = shouldArmHistorySpare(state);
  useEffect(() => {
    if (!needsSpare) {
      return;
    }
    // Nothing below this page means Back cannot leave it, so there is nothing
    // to guard against. Arming anyway would make a dead button look live, and
    // then offer to leave for somewhere that does not exist.
    if (window.history.length <= 1) {
      return;
    }
    armSpare();
  }, [armSpare, needsSpare]);

  const workspace = state.held;
  const question = questionView(state);
  const enabled = commandsEnabled(state);
  const busy = workspaceBusy(state);
  const spinning = spinningButton(state);
  const announcement = state.announcement;

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
            disabled={!enabled}
            onClick={() => dispatch({ type: "newClicked" })}
            type="button"
          >
            {spinning === "new" ? (
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
            disabled={!enabled}
            onClick={() => dispatch({ type: "openClicked" })}
            type="button"
          >
            {spinning === "open" ? (
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

      {question !== null ? (
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
              {question.heading}
            </h2>
          </div>
          <p className="mt-3 text-sm text-fd-muted-foreground">
            {/* Derived from the same state the hold decision was made from, so
                the sentence cannot describe a different one. */}
            {question.sentence} {question.consequence}, and that work is gone
          </p>
          <div className="mt-5 flex flex-wrap gap-3">
            <button
              className={secondaryButtonClass}
              data-testid="workspace-confirm-discard"
              disabled={!enabled}
              onClick={() => dispatch({ type: "confirmClicked" })}
              type="button"
            >
              {question.confirmLabel}
            </button>
            <button
              className={primaryButtonClass}
              data-testid="workspace-confirm-cancel"
              disabled={!enabled}
              onClick={() => dispatch({ type: "cancelClicked" })}
              type="button"
            >
              Keep this workspace
            </button>
          </div>
        </section>
      ) : null}

      {workspace !== null ? (
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
              disabled={!enabled}
              onClick={() => dispatch({ type: "saveClicked", how: "save" })}
              type="button"
            >
              {spinning === "save" ? (
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
              disabled={!enabled}
              onClick={() => dispatch({ type: "saveClicked", how: "saveAs" })}
              type="button"
            >
              {spinning === "saveAs" ? (
                <LoaderCircle
                  aria-hidden="true"
                  className="size-4 animate-spin"
                />
              ) : (
                <Download aria-hidden="true" className="size-4" />
              )}
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
      {workspace !== null ? (
        <WorkspaceImport
          busy={busy}
          client={client}
          existingTableNames={workspace.summary.tables.map(
            (table) => table.name,
          )}
          onFailed={onImportFailed}
          onImported={onImported}
          onReadFinished={onReadFinished}
          onReading={onReading}
          onRunning={onRunning}
        />
      ) : null}

      {/* The grid keeps no flag the shell already keeps: which tables exist and
          which database they belong to are both read from the summary, editing
          is locked by the one derived answer, and its edits and its teardown
          are reported as events. It does keep its own refusals, because those
          belong to the cells they name rather than to the page, and a slot
          shared with this component's own failures is one each would clear from
          under the other. */}
      {workspace !== null ? (
        <WorkspaceGrid
          getClient={client}
          locked={editingLocked(state)}
          onDetached={onDetached}
          onEditorClosed={onEditorClosed}
          onEditorOpened={onEditorOpened}
          onEditSent={onEditSent}
          onEditSettled={onEditSettled}
          summary={workspace.summary}
        />
      ) : null}

      {announcement?.kind === "notice" ? (
        <p
          aria-live="polite"
          className="rounded-lg border border-fd-primary/40 bg-fd-accent/30 px-4 py-3 text-sm text-fd-accent-foreground"
          data-testid="workspace-notice"
          role="status"
        >
          {announcement.text}
        </p>
      ) : null}

      {announcement?.kind === "error" ? (
        <pre
          aria-live="polite"
          className="overflow-x-auto whitespace-pre-wrap rounded-lg border border-fd-primary/40 bg-fd-accent/30 px-4 py-3 text-xs leading-6 text-fd-accent-foreground"
          data-testid="workspace-error"
        >
          {announcement.text}
        </pre>
      ) : null}
    </ToolShell>
  );
}
