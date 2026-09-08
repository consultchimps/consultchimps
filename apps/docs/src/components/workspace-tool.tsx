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
 * Data import lives in `workspace-import.tsx` and is mounted below the summary;
 * a grid is still to come. Until it does, the table listing in the summary is
 * how a person sees what a workspace holds: each table's name, row count,
 * Record ID prefix, and the column types import inferred.
 *
 * The workspace is one in-memory database, so replacing it or leaving the page
 * is the whole of losing it. The shell therefore owns the unsaved-changes flag
 * rather than each mutating feature: `markChanged` is the single place it is
 * set, a successful write is the only place it is cleared, and New, Open, and
 * closing the tab all ask first while it is set. A feature that changes the
 * workspace inherits the guard by calling that one marker.
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
import { WorkspaceImport } from "@/components/workspace-import";
import { WORKSPACE_FILES } from "@/lib/accepted-files";
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

/** The replacement a confirmation is standing in front of, or null for none. */
type PendingReplacement = "new" | "open" | null;

/**
 * The shell's record-a-change callback, handed to any section that mutates the
 * workspace. Import receives it folded into `onImported`; the record grid takes
 * it directly when its cell edits land. Passing a summary replaces the table
 * listing at the same time, and passing nothing marks the workspace changed
 * without one.
 */
export type MarkWorkspaceChanged = (summary?: WorkspaceSummary) => void;

type Busy = "creating" | "opening" | "saving" | null;

export function WorkspaceTool() {
  const clientRef = useRef<WorkspaceClient | null>(null);
  // The handle for an in-place save, held only when a picker granted one.
  const handleRef = useRef<WorkspaceFileHandle | null>(null);
  const fallbackInputRef = useRef<HTMLInputElement | null>(null);

  const [workspace, setWorkspace] = useState<OpenWorkspace | null>(null);
  const [busy, setBusy] = useState<Busy>(null);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [pending, setPending] = useState<PendingReplacement>(null);

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

  // Both entry points that replace the held workspace go through here, so a
  // mutating feature added later inherits the guard by doing nothing.
  const replaceWorkspace = useCallback(
    (replacement: Exclude<PendingReplacement, null>) => {
      if (hasUnsavedChanges) {
        setPending(replacement);
        return;
      }
      void (replacement === "new" ? startNew() : startOpen());
    },
    [hasUnsavedChanges, startNew, startOpen],
  );

  const onConfirmReplace = useCallback(() => {
    const replacement = pending;
    setPending(null);
    if (replacement !== null) {
      void (replacement === "new" ? startNew() : startOpen());
    }
  }, [pending, startNew, startOpen]);

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

  // Warn before the tab closes or navigates away, which is the one way to lose
  // the workspace that no button on this page controls.
  useEffect(() => {
    if (!hasUnsavedChanges) {
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
  }, [hasUnsavedChanges]);

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
  const confirming = pending !== null && hasUnsavedChanges;

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
              Unsaved changes
            </h2>
          </div>
          <p className="mt-3 text-sm text-fd-muted-foreground">
            This workspace has changes that have not been saved to a file.{" "}
            {pending === "new"
              ? "Starting a new workspace"
              : "Opening another workspace"}{" "}
            replaces it, and those changes are gone
          </p>
          <div className="mt-5 flex flex-wrap gap-3">
            <button
              className={secondaryButtonClass}
              data-testid="workspace-confirm-discard"
              disabled={isBusy}
              onClick={onConfirmReplace}
              type="button"
            >
              Discard the changes and continue
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
          client={client}
          disabled={isBusy}
          existingTableNames={workspace.summary.tables.map(
            (table) => table.name,
          )}
          onImported={onImported}
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
