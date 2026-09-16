"use client";

import {
  compactButtonClass,
  describeFailure,
  inputClass,
  primaryButtonClass,
  secondaryButtonClass,
  sectionClass,
  ToolShell,
} from "@/components/tool-kit";
import { WorkspaceImport } from "@/components/workspace-import";
import { WorkspaceSchemaReview } from "@/components/workspace-schema-review";
import { ConsultChimpsError, isConsultChimpsError } from "@consultchimps/core";
import type {
  WorkspaceDatabaseFormat,
  WorkspaceDatabaseListing,
  WorkspaceDeliveryPage,
  WorkspaceDeliverySummary,
  WorkspaceProgress,
  WorkspaceSchemaDocument,
  WorkspaceSchemaPlan,
  WorkspaceStoredDatabase,
  WorkspaceSummary,
} from "@/lib/workspace-protocol";
import { WorkspaceClient } from "@/lib/workspace-worker";
import {
  browserExportCleanupIntervalMilliseconds,
  cleanupExpiredBrowserExports,
  retainBrowserExportLease,
} from "@/lib/workspace-files";
import {
  claimDatabaseTool,
  type DatabaseToolOwnership,
  waitForDatabaseTool,
} from "@/lib/workspace-ownership";
import {
  CircleCheck,
  Database,
  Download,
  FilePlus,
  FolderOpen,
  LoaderCircle,
  RefreshCw,
  Trash2,
  TriangleAlert,
} from "lucide-react";
import { useCallback, useEffect, useRef, useState } from "react";

const DATABASE_ACCEPT = ".sqlite,.sqlite3,.db,.duckdb";
const CONFLICT_NOTICE_GRACE_MILLISECONDS = 300;
const RECENT_DATABASES_KEY = "consultchimps.workspace.databases.v1";
const DEFAULT_SCHEMA = `{
  "version": 1,
  "tables": [
    {
      "name": "datasets",
      "recordId": { "prefix": "DATASET", "padding": 6 },
      "columns": [
        { "name": "dataset_name", "type": "text", "nullable": false },
        { "name": "is_active", "type": "boolean" }
      ]
    }
  ]
}`;

/** The part of the page whose controls produced a notice or error. */
type StatusSection = "start" | "schema" | "import" | "history" | "export";

interface WorkspaceStatus {
  readonly kind: "error" | "notice" | "cancelled";
  readonly section: StatusSection;
  readonly message: string;
  /** A recovery the reader can take from the notice itself. */
  readonly action?: { readonly label: string; readonly run: () => void };
}

const STATUS_HEADINGS: Record<WorkspaceStatus["kind"], string> = {
  error: "Something went wrong",
  notice: "Done",
  cancelled: "Cancelled",
};

function StatusNotice({
  status,
  disabled,
}: {
  readonly status: WorkspaceStatus;
  readonly disabled: boolean;
}) {
  const error = status.kind === "error";
  return (
    <div
      className={
        error
          ? "mt-4 rounded-lg border-2 border-fd-primary bg-fd-primary/10 px-4 py-3"
          : "mt-4 rounded-lg border border-fd-primary/40 bg-fd-accent/30 px-4 py-3"
      }
      data-section={status.section}
      data-testid={error ? "workspace-error" : "workspace-notice"}
      role={error ? "alert" : "status"}
    >
      <div className="flex items-center gap-2 text-sm font-semibold">
        {error ? (
          <TriangleAlert
            aria-hidden="true"
            className="size-4 text-fd-primary"
          />
        ) : (
          <CircleCheck aria-hidden="true" className="size-4 text-fd-primary" />
        )}
        {STATUS_HEADINGS[status.kind]}
      </div>
      <pre className="mt-2 overflow-x-auto whitespace-pre-wrap text-xs leading-6">
        {status.message}
      </pre>
      {status.action === undefined ? null : (
        <button
          className={`${secondaryButtonClass} mt-3`}
          data-testid="workspace-status-action"
          disabled={disabled}
          onClick={status.action.run}
          type="button"
        >
          <RefreshCw aria-hidden="true" className="size-4" />
          {status.action.label}
        </button>
      )}
    </div>
  );
}

interface StorageUsage {
  readonly usage: number | null;
  readonly quota: number | null;
}

/**
 * The working copies themselves come from browser storage. This remembered
 * list of names only orders them, most recently used first, so a copy that
 * was removed behind the tool's back never shows up and a copy created
 * elsewhere does. Entries written by earlier versions carried a format as
 * well; only the name is read.
 */
function readRecentNames(): string[] {
  try {
    const stored: unknown = JSON.parse(
      window.localStorage.getItem(RECENT_DATABASES_KEY) ?? "[]",
    );
    if (!Array.isArray(stored)) return [];
    return stored.flatMap((entry): string[] =>
      isRecord(entry) && typeof entry["name"] === "string"
        ? [entry["name"]]
        : [],
    );
  } catch {
    return [];
  }
}

function writeRecentNames(names: readonly string[]): void {
  try {
    window.localStorage.setItem(
      RECENT_DATABASES_KEY,
      JSON.stringify(names.map((name) => ({ name }))),
    );
  } catch {
    // Ordering is a convenience; storage that refuses it changes nothing.
  }
}

function markRecent(name: string): void {
  writeRecentNames(
    [name, ...readRecentNames().filter((entry) => entry !== name)].slice(0, 8),
  );
}

function forgetRecent(name: string): void {
  writeRecentNames(readRecentNames().filter((entry) => entry !== name));
}

// Names outside the recent list keep the runtime's own order (code units),
// so the page and the runtime agree on "alphabetical".
function orderByRecentUse(
  stored: readonly WorkspaceStoredDatabase[],
): readonly WorkspaceStoredDatabase[] {
  const rank = new Map(readRecentNames().map((name, index) => [name, index]));
  return [...stored].sort((left, right) => {
    const leftRank = rank.get(left.name) ?? Number.POSITIVE_INFINITY;
    const rightRank = rank.get(right.name) ?? Number.POSITIVE_INFINITY;
    if (leftRank !== rightRank) return leftRank - rightRank;
    return left.name < right.name ? -1 : left.name > right.name ? 1 : 0;
  });
}

async function readStorageUsage(): Promise<StorageUsage | null> {
  const storage: unknown = navigator.storage;
  if (
    typeof storage !== "object" ||
    storage === null ||
    !("estimate" in storage) ||
    typeof storage.estimate !== "function"
  ) {
    return null;
  }
  try {
    const estimate: unknown = await (
      storage.estimate as () => Promise<unknown>
    ).call(storage);
    if (!isRecord(estimate)) return null;
    const usage = estimate["usage"];
    const quota = estimate["quota"];
    return {
      usage: typeof usage === "number" ? usage : null,
      quota: typeof quota === "number" ? quota : null,
    };
  } catch {
    return null;
  }
}

function formatBytes(bytes: number): string {
  const units = ["bytes", "KB", "MB", "GB", "TB"];
  let value = bytes;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit += 1;
  }
  return `${unit === 0 ? String(Math.round(value)) : value.toFixed(1)} ${units[unit]}`;
}

function storageUsageText(usage: StorageUsage): string | null {
  if (usage.usage === null) return null;
  return usage.quota === null
    ? `Browser storage used by this site: ${formatBytes(usage.usage)}`
    : `Browser storage used by this site: ${formatBytes(usage.usage)} of ${formatBytes(usage.quota)} available`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function parseSchema(text: string): WorkspaceSchemaDocument {
  const value: unknown = JSON.parse(text);
  if (!isRecord(value)) {
    throw new Error("The schema must be a JSON object");
  }
  return value;
}

async function downloadFile(file: File, name: string): Promise<() => void> {
  const releaseLease = await retainBrowserExportLease(file.name);
  let url: string | undefined;
  let anchor: HTMLAnchorElement | undefined;
  try {
    const downloadUrl = URL.createObjectURL(file);
    url = downloadUrl;
    anchor = document.createElement("a");
    anchor.href = downloadUrl;
    anchor.download = name;
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
    return () => {
      URL.revokeObjectURL(downloadUrl);
      releaseLease();
    };
  } catch (error) {
    anchor?.remove();
    if (url !== undefined) URL.revokeObjectURL(url);
    releaseLease();
    throw error;
  }
}

function ProgressNotice({
  progress,
}: {
  readonly progress: WorkspaceProgress;
}) {
  const percent =
    progress.total === null || progress.total === 0
      ? null
      : Math.min(100, Math.round((progress.completed / progress.total) * 100));
  return (
    <div
      className="rounded-lg border bg-fd-muted/45 px-4 py-3"
      data-testid="workspace-progress"
      role="status"
    >
      <div className="flex items-center justify-between gap-4 text-sm">
        <span>{progress.message}</span>
        {percent === null ? null : <span>{percent}%</span>}
      </div>
      {percent === null ? null : (
        <progress className="mt-2 w-full" max={100} value={percent} />
      )}
    </div>
  );
}

function Summary({ summary }: { readonly summary: WorkspaceSummary }) {
  return (
    <section className={sectionClass} data-testid="workspace-summary">
      <div className="flex flex-wrap items-start justify-between gap-4">
        <div>
          <p className="font-mono text-xs uppercase tracking-[0.16em] text-fd-primary">
            Persistent working database
          </p>
          <h2 className="mt-2 font-display text-2xl font-semibold">
            {summary.workingCopyName}
          </h2>
          <p className="mt-2 max-w-2xl text-sm text-fd-muted-foreground">
            Stored in this browser&apos;s private file system. Changes commit to
            this working copy. A file you opened and any file you exported are
            separate copies
          </p>
        </div>
        <span
          className="rounded-full border px-3 py-1 font-mono text-xs uppercase"
          data-testid="workspace-format"
        >
          {summary.format}
        </span>
      </div>
      <dl className="mt-5 grid gap-3 text-sm sm:grid-cols-4">
        <div>
          <dt className="text-fd-muted-foreground">Tables</dt>
          <dd data-testid="workspace-table-count">{summary.tables.length}</dd>
        </div>
        <div>
          <dt className="text-fd-muted-foreground">Applied batches</dt>
          <dd>{summary.importCount}</dd>
        </div>
        <div>
          <dt className="text-fd-muted-foreground">Recorded batches</dt>
          <dd data-testid="workspace-delivery-count">
            {summary.deliveryCount}
          </dd>
        </div>
        <div>
          <dt className="text-fd-muted-foreground">Format version</dt>
          <dd>{summary.formatVersion}</dd>
        </div>
      </dl>
      {summary.tables.length === 0 ? (
        <p className="mt-5 rounded-lg border border-dashed p-4 text-sm text-fd-muted-foreground">
          This database has no user tables yet
        </p>
      ) : (
        <div className="mt-5 grid gap-3 sm:grid-cols-2">
          {summary.tables.map((table) => (
            <article
              className="rounded-lg border p-4"
              key={table.id}
              data-testid="workspace-table"
            >
              <div className="flex items-baseline justify-between gap-3">
                <h3 className="font-semibold">{table.name}</h3>
                <span className="text-xs text-fd-muted-foreground">
                  {table.rowCount.toLocaleString()}{" "}
                  {table.rowCount === 1 ? "row" : "rows"}
                </span>
              </div>
              <p className="mt-2 text-xs text-fd-muted-foreground">
                {table.columns
                  .map((column) => `${column.name} (${column.type})`)
                  .join(", ")}
              </p>
            </article>
          ))}
        </div>
      )}
    </section>
  );
}

function deliveryScopeText(scope: WorkspaceDeliverySummary["scope"]): string {
  switch (scope.kind) {
    case "full":
      return "Full coverage";
    case "partial":
      return `Partial coverage: ${scope.description}`;
    case "changes":
      return `Changes since ${scope.baseline}`;
    case "unknown":
      return "Unknown coverage";
  }
}

export function WorkspaceTool() {
  const clientRef = useRef<WorkspaceClient | null>(null);
  const openInputRef = useRef<HTMLInputElement | null>(null);
  const controllerRef = useRef<AbortController | null>(null);
  const deliveriesRequestRef = useRef(0);
  const exportLeaseReleasesRef = useRef<Set<() => void>>(new Set());
  const schemaReviewRevisionRef = useRef(0);
  const [summary, setSummary] = useState<WorkspaceSummary | null>(null);
  const [workspaceGeneration, setWorkspaceGeneration] = useState(0);
  const [format, setFormat] = useState<WorkspaceDatabaseFormat>("sqlite");
  const [name, setName] = useState("consultchimps.sqlite");
  const [openName, setOpenName] = useState("");
  const [openOverwrite, setOpenOverwrite] = useState(false);
  const [busy, setBusy] = useState<string | null>(null);
  const [cancellable, setCancellable] = useState(true);
  const [progress, setProgress] = useState<WorkspaceProgress | null>(null);
  const [reviewActive, setReviewActive] = useState(false);
  const [status, setStatus] = useState<WorkspaceStatus | null>(null);
  const [schemaText, setSchemaText] = useState(DEFAULT_SCHEMA);
  const [schemaPlan, setSchemaPlan] = useState<WorkspaceSchemaPlan | null>(
    null,
  );
  const [deliveries, setDeliveries] = useState<WorkspaceDeliveryPage | null>(
    null,
  );
  const [deliveriesLoading, setDeliveriesLoading] = useState(false);
  const [deliveriesError, setDeliveriesError] = useState<string | null>(null);
  const [copies, setCopies] = useState<WorkspaceDatabaseListing | null>(null);
  const [copiesError, setCopiesError] = useState<string | null>(null);
  const [storageUsage, setStorageUsage] = useState<StorageUsage | null>(null);
  const [pendingDelete, setPendingDelete] = useState<string | null>(null);
  const copiesRequestRef = useRef(0);
  const [ownership, setOwnership] = useState<DatabaseToolOwnership | null>(
    null,
  );

  // Claim the tool for this page before any engine starts. A conflict keeps
  // the page waiting in the background and takes over when the other tab
  // releases the lock. The claim is held until the page unmounts.
  //
  // The conflict notice is shown after a short grace period: a remount (React
  // StrictMode in development, Fast Refresh) can see its own previous claim
  // still held for a few milliseconds, and the background wait then takes
  // over at once. A real conflict with another tab outlasts the grace period.
  useEffect(() => {
    const controller = new AbortController();
    let release: (() => void) | null = null;
    let conflictTimer: number | null = null;
    const accept = (result: DatabaseToolOwnership): boolean => {
      if (controller.signal.aborted) {
        if (result.state === "owned") result.release();
        return false;
      }
      if (conflictTimer !== null) {
        window.clearTimeout(conflictTimer);
        conflictTimer = null;
      }
      if (result.state === "owned") release = result.release;
      if (result.state === "conflict") {
        conflictTimer = window.setTimeout(() => {
          conflictTimer = null;
          setOwnership(result);
        }, CONFLICT_NOTICE_GRACE_MILLISECONDS);
      } else {
        setOwnership(result);
      }
      return true;
    };
    void claimDatabaseTool()
      .then(async (result) => {
        if (!accept(result) || result.state !== "conflict") return;
        accept(await waitForDatabaseTool(controller.signal));
      })
      .catch(() => accept({ state: "unavailable" }));
    return () => {
      controller.abort();
      if (conflictTimer !== null) window.clearTimeout(conflictTimer);
      release?.();
    };
  }, []);

  useEffect(() => {
    const exportLeaseReleases = exportLeaseReleasesRef.current;
    const cleanup = (): void => {
      void cleanupExpiredBrowserExports().catch(() => undefined);
    };
    cleanup();
    const interval = window.setInterval(
      cleanup,
      browserExportCleanupIntervalMilliseconds,
    );
    return () => {
      window.clearInterval(interval);
      for (const release of exportLeaseReleases) release();
      exportLeaseReleases.clear();
    };
  }, []);

  const client = useCallback(() => {
    clientRef.current ??= new WorkspaceClient();
    return clientRef.current;
  }, []);

  // Read the working copies from storage. Only the latest request may
  // publish its answer, so a slow listing cannot overwrite a newer one.
  const refreshCopies = useCallback(async () => {
    const request = copiesRequestRef.current + 1;
    copiesRequestRef.current = request;
    try {
      const listing = await client().listDatabases();
      const usage = await readStorageUsage();
      if (copiesRequestRef.current !== request) return;
      setCopies({
        databases: orderByRecentUse(listing.databases),
        ignored: listing.ignored,
      });
      setStorageUsage(usage);
      setCopiesError(null);
    } catch (error) {
      if (copiesRequestRef.current !== request) return;
      setCopiesError(describeFailure(error));
    }
  }, [client]);

  useEffect(() => {
    if (ownership === null || ownership.state === "conflict") return;
    void refreshCopies();
  }, [ownership, refreshCopies]);

  const remember = useCallback(
    (next: WorkspaceSummary) => {
      markRecent(next.workingCopyName);
      void refreshCopies();
    },
    [refreshCopies],
  );

  useEffect(
    () => () => {
      controllerRef.current?.abort();
      clientRef.current?.terminate();
    },
    [],
  );

  // A summary that could not be refreshed gets its recovery attached here, in
  // one place, whichever operation reported it. The action is reached through
  // a ref because it runs a long operation that itself reports errors.
  const refreshSummaryRef = useRef<(section: StatusSection) => void>(
    () => undefined,
  );
  const reportError = useCallback((error: unknown, section: StatusSection) => {
    const refreshRequired =
      isConsultChimpsError(error) &&
      error.code === "DB_BROWSER_SUMMARY_REFRESH_REQUIRED";
    setStatus({
      kind: "error",
      section,
      message: describeFailure(error),
      ...(refreshRequired
        ? {
            action: {
              label: "Refresh summary",
              run: () => refreshSummaryRef.current(section),
            },
          }
        : {}),
    });
  }, []);
  const reportImportError = useCallback(
    (error: unknown) => reportError(error, "import"),
    [reportError],
  );

  const clearDeliveries = useCallback(() => {
    deliveriesRequestRef.current += 1;
    setDeliveries(null);
    setDeliveriesLoading(false);
    setDeliveriesError(null);
  }, []);

  const invalidateSchemaReview = useCallback(() => {
    schemaReviewRevisionRef.current += 1;
    setSchemaPlan(null);
  }, []);

  const runLong = useCallback(
    async <T,>(
      label: string,
      run: (options: {
        readonly signal: AbortSignal;
        readonly onProgress: (next: WorkspaceProgress) => void;
      }) => Promise<T>,
      options: {
        readonly cancellable?: boolean;
        readonly section?: StatusSection;
      } = {},
    ): Promise<T | null> => {
      const controller = new AbortController();
      controllerRef.current = controller;
      const section = options.section ?? "start";
      setBusy(label);
      setCancellable(options.cancellable !== false);
      setProgress(null);
      setStatus(null);
      try {
        return await run({
          signal: controller.signal,
          onProgress: setProgress,
        });
      } catch (error) {
        if (
          (error instanceof DOMException && error.name === "AbortError") ||
          (isConsultChimpsError(error) && error.code === "OPERATION_ABORTED")
        ) {
          setStatus({
            kind: "cancelled",
            section,
            message:
              "Cancelled. The accepted database state was left unchanged",
          });
        } else {
          reportError(error, section);
        }
        return null;
      } finally {
        controllerRef.current = null;
        setBusy(null);
        setProgress(null);
      }
    },
    [reportError],
  );

  useEffect(() => {
    if (busy === null && !reviewActive) return;
    const hold = (event: BeforeUnloadEvent): void => {
      event.preventDefault();
      event.returnValue = "";
    };
    window.addEventListener("beforeunload", hold);
    return () => window.removeEventListener("beforeunload", hold);
  }, [busy, reviewActive]);

  // Batch history loads on its own whenever a database is opened and after
  // every applied or recorded batch, so the count in the summary and the
  // list below it never disagree until someone clicks Refresh. A load that
  // fails reports inside the history section rather than through the page
  // status: it often runs in the background, and must not replace the notice
  // or error of the action that started it.
  const loadDeliveries = useCallback(
    async (cursor: string | null) => {
      const request = deliveriesRequestRef.current + 1;
      deliveriesRequestRef.current = request;
      setDeliveriesLoading(true);
      setDeliveriesError(null);
      try {
        const page = await client().listDeliveries(cursor);
        if (deliveriesRequestRef.current === request) setDeliveries(page);
      } catch (error) {
        if (deliveriesRequestRef.current === request) {
          setDeliveriesError(describeFailure(error));
        }
      } finally {
        if (deliveriesRequestRef.current === request) {
          setDeliveriesLoading(false);
        }
      }
    },
    [client],
  );
  const reloadDeliveries = useCallback(() => {
    void loadDeliveries(null);
  }, [loadDeliveries]);
  const runImportOperation = useCallback(
    <T,>(
      label: string,
      run: (options: {
        readonly signal: AbortSignal;
        readonly onProgress: (next: WorkspaceProgress) => void;
      }) => Promise<T>,
    ) => runLong(label, run, { section: "import" }),
    [runLong],
  );

  const create = useCallback(async () => {
    const created = await runLong("Creating database", (options) =>
      client().create(format, name.trim(), options),
    );
    if (created === null) return;
    setSummary(created);
    setWorkspaceGeneration((current) => current + 1);
    remember(created);
    invalidateSchemaReview();
    clearDeliveries();
    void loadDeliveries(null);
    setStatus({
      kind: "notice",
      section: "start",
      message: "Created a persistent browser working database",
    });
  }, [
    clearDeliveries,
    client,
    format,
    invalidateSchemaReview,
    loadDeliveries,
    name,
    remember,
    runLong,
  ]);

  const open = useCallback(
    async (file: File) => {
      const workingCopyName = openName.trim();
      const opened = await runLong(
        "Copying database into browser storage",
        (options) =>
          client().open(
            file,
            {
              ...(workingCopyName === "" ? {} : { name: workingCopyName }),
              overwrite: openOverwrite,
            },
            options,
          ),
      );
      if (opened === null) return;
      setSummary(opened);
      setWorkspaceGeneration((current) => current + 1);
      remember(opened);
      invalidateSchemaReview();
      clearDeliveries();
      void loadDeliveries(null);
      setStatus({
        kind: "notice",
        section: "start",
        message: `Opened "${file.name}" as browser working copy "${opened.workingCopyName}". The selected file will not change`,
      });
    },
    [
      clearDeliveries,
      client,
      invalidateSchemaReview,
      loadDeliveries,
      openName,
      openOverwrite,
      remember,
      runLong,
    ],
  );

  const reopen = useCallback(
    async (workingCopyName: string) => {
      const opened = await runLong("Opening browser working copy", (options) =>
        client().reopen(workingCopyName, options),
      );
      if (opened === null) return;
      setSummary(opened);
      setWorkspaceGeneration((current) => current + 1);
      remember(opened);
      invalidateSchemaReview();
      clearDeliveries();
      void loadDeliveries(null);
      setStatus({
        kind: "notice",
        section: "start",
        message: `Reopened "${workingCopyName}" from browser storage`,
      });
    },
    [
      clearDeliveries,
      client,
      invalidateSchemaReview,
      loadDeliveries,
      remember,
      runLong,
    ],
  );

  // The page drops its view of the database whether or not the worker
  // finished every cleanup: a worker that could not close a resource keeps
  // it until the next open replaces it, and the page's error says so. The
  // reverse (a page that still shows a closed database) would route every
  // later action into "create or open a database first".
  const closeDatabase = useCallback(async () => {
    if (summary === null) return;
    const closedName = summary.workingCopyName;
    const closed = await runLong("Closing database", () => client().close(), {
      cancellable: false,
    });
    setSummary(null);
    setWorkspaceGeneration((current) => current + 1);
    invalidateSchemaReview();
    clearDeliveries();
    if (closed !== null) {
      setStatus({
        kind: "notice",
        section: "start",
        message: `Closed "${closedName}". It stays in browser storage until you delete it`,
      });
    }
    void refreshCopies();
  }, [
    clearDeliveries,
    client,
    invalidateSchemaReview,
    refreshCopies,
    runLong,
    summary,
  ]);

  const removeDatabase = useCallback(
    async (workingCopyName: string) => {
      setPendingDelete(null);
      const removed = await runLong(
        "Deleting working copy",
        () => client().removeDatabase(workingCopyName),
        { cancellable: false },
      );
      if (removed === null) return;
      forgetRecent(workingCopyName);
      setStatus({
        kind: "notice",
        section: "start",
        message: `Deleted "${workingCopyName}" from browser storage. Files you opened or exported are not affected`,
      });
      void refreshCopies();
    },
    [client, refreshCopies, runLong],
  );

  const planSchema = useCallback(async () => {
    try {
      const revision = schemaReviewRevisionRef.current;
      const schema = parseSchema(schemaText);
      setSchemaPlan(null);
      const plan = await runLong(
        "Reviewing schema",
        (options) => client().planSchema(schema, options),
        { section: "schema" },
      );
      if (plan === null || revision !== schemaReviewRevisionRef.current) return;
      setSchemaPlan(plan);
      setStatus({
        kind: "notice",
        section: "schema",
        message: plan.ready
          ? "Schema review is ready to apply"
          : "Schema conflicts need a decision before anything can change",
      });
    } catch (error) {
      reportError(error, "schema");
    }
  }, [client, reportError, runLong, schemaText]);

  const applySchema = useCallback(async () => {
    if (schemaPlan === null || !schemaPlan.ready) return;
    const result = await runLong(
      "Applying schema",
      (options) => client().applySchema(schemaPlan.id, options),
      { section: "schema" },
    );
    if (result === null) return;
    if (result.summary.state === "updated") {
      setSummary(result.summary.value);
    }
    invalidateSchemaReview();
    if (result.checkpoint.state === "failed") {
      reportError(
        new ConsultChimpsError(
          result.checkpoint.code,
          result.checkpoint.message,
        ),
        "schema",
      );
      return;
    }
    if (result.summary.state === "refresh-required") {
      reportError(
        new ConsultChimpsError(result.summary.code, result.summary.message),
        "schema",
      );
      return;
    }
    setStatus({
      kind: "notice",
      section: "schema",
      message: "Applied the reviewed schema changes",
    });
  }, [client, invalidateSchemaReview, reportError, runLong, schemaPlan]);

  // The refresh reports where the error that offered it was shown.
  const refreshSummary = useCallback(
    async (section: StatusSection) => {
      const refreshed = await runLong(
        "Refreshing summary",
        (options) => client().refreshSummary(options),
        { cancellable: false, section },
      );
      if (refreshed === null) return;
      setSummary(refreshed);
      setStatus({
        kind: "notice",
        section,
        message: "Refreshed the database summary",
      });
    },
    [client, runLong],
  );
  useEffect(() => {
    refreshSummaryRef.current = (section) => void refreshSummary(section);
  }, [refreshSummary]);

  const exportDatabase = useCallback(
    async (targetFormat: WorkspaceDatabaseFormat) => {
      const exported = await runLong(
        "Creating a consistent export",
        (options) => client().export(targetFormat, options),
        { section: "export" },
      );
      if (exported === null) return;
      try {
        exportLeaseReleasesRef.current.add(
          await downloadFile(exported.file, exported.name),
        );
        setStatus({
          kind: "notice",
          section: "export",
          message: `Exported an independent ${exported.format} copy. Later browser changes will not update it`,
        });
      } catch (error) {
        reportError(error, "export");
      }
    },
    [client, reportError, runLong],
  );

  const tabConflict = ownership?.state === "conflict";
  const disabled = busy !== null || ownership === null || tabConflict;
  const statusFor = (section: StatusSection) =>
    status?.section === section ? (
      <StatusNotice disabled={disabled} status={status} />
    ) : null;
  return (
    <ToolShell
      description="Create or open a persistent local database, review workbook batches, record batch context, and export a portable copy"
      guideHref="/docs/tools/data-workspace"
      guideLabel="Read the database guide"
      kicker="Local database"
      title="Database"
    >
      {tabConflict ? (
        <section
          className={`${sectionClass} border-fd-primary/40`}
          data-testid="workspace-tab-conflict"
          role="alert"
        >
          <h2 className="font-display text-xl font-semibold">
            The database tool is open in another tab
          </h2>
          <p className="mt-2 text-sm text-fd-muted-foreground">
            This browser can run the database tool in one tab or window at a
            time, because the working database keeps exclusive file handles.
            Finish your work in the other tab or close it. This page takes over
            on its own once the other tab lets go
          </p>
        </section>
      ) : null}
      <section className={sectionClass} data-testid="workspace-start">
        <div className="grid gap-6 lg:grid-cols-2">
          <div>
            <h2 className="font-display text-xl font-semibold">New database</h2>
            <div className="mt-4 grid gap-3 sm:grid-cols-[12rem_1fr]">
              <label className="text-sm">
                Format
                <select
                  className={`${inputClass} mt-1`}
                  data-testid="workspace-new-format"
                  disabled={disabled}
                  onChange={(event) => {
                    const next = event.target.value;
                    if (next !== "sqlite" && next !== "duckdb") return;
                    setFormat(next);
                    setName(
                      next === "sqlite"
                        ? "consultchimps.sqlite"
                        : "consultchimps.duckdb",
                    );
                  }}
                  value={format}
                >
                  <option value="sqlite">SQLite</option>
                  <option value="duckdb">DuckDB</option>
                </select>
              </label>
              <label className="text-sm">
                Working copy name
                <input
                  className={`${inputClass} mt-1`}
                  data-testid="workspace-new-name"
                  disabled={disabled}
                  onChange={(event) => setName(event.target.value)}
                  value={name}
                />
              </label>
            </div>
            <div className="mt-3">
              <button
                className={primaryButtonClass}
                data-testid="workspace-new"
                disabled={disabled || name.trim() === ""}
                onClick={() => void create()}
                type="button"
              >
                {busy === "Creating database" ? (
                  <LoaderCircle
                    aria-hidden="true"
                    className="size-4 animate-spin"
                  />
                ) : (
                  <FilePlus aria-hidden="true" className="size-4" />
                )}
                Create
              </button>
            </div>
          </div>
          <div>
            <h2 className="font-display text-xl font-semibold">
              Open database file
            </h2>
            <label className="mt-4 block text-sm">
              Imported working copy name
              <input
                className={`${inputClass} mt-1`}
                data-testid="workspace-open-name"
                disabled={disabled}
                onChange={(event) => {
                  const next = event.target.value;
                  setOpenName(next);
                  if (next.trim() === "") setOpenOverwrite(false);
                }}
                placeholder="Create a unique name from the file"
                value={openName}
              />
            </label>
            <p className="mt-2 text-xs text-fd-muted-foreground">
              Leave blank to create a unique name from the selected file
            </p>
            <label className="mt-3 flex items-start gap-2.5 text-sm font-medium leading-6">
              <input
                checked={openOverwrite}
                className="mt-1 size-4 shrink-0 rounded border-fd-border accent-fd-primary"
                data-testid="workspace-open-overwrite"
                disabled={disabled || openName.trim() === ""}
                onChange={(event) => setOpenOverwrite(event.target.checked)}
                type="checkbox"
              />
              Replace an existing browser working copy with this name
            </label>
            <div className="mt-3">
              <button
                className={secondaryButtonClass}
                data-testid="workspace-open"
                disabled={disabled}
                onClick={() => openInputRef.current?.click()}
                type="button"
              >
                <FolderOpen aria-hidden="true" className="size-4" />
                Open file
              </button>
            </div>
            <input
              accept={DATABASE_ACCEPT}
              className="sr-only"
              data-testid="workspace-open-input"
              onChange={(event) => {
                const file = event.target.files?.[0];
                event.target.value = "";
                if (file !== undefined) void open(file);
              }}
              ref={openInputRef}
              type="file"
            />
          </div>
        </div>
        <p className="mt-4 text-sm text-fd-muted-foreground">
          The working database stays in origin-private browser storage. Opening
          a file copies it there in bounded chunks and leaves the selected file
          unchanged. Verified in Chromium-based desktop browsers such as Chrome
          and Edge. Use the tool in one tab at a time
        </p>
        {copies === null && copiesError === null ? null : (
          <div className="mt-5" data-testid="workspace-recent">
            <div className="flex flex-wrap items-baseline justify-between gap-3">
              <h3 className="text-sm font-semibold">Browser working copies</h3>
              {storageUsage === null ||
              storageUsageText(storageUsage) === null ? null : (
                <span
                  className="text-xs text-fd-muted-foreground"
                  data-testid="workspace-storage-usage"
                >
                  {storageUsageText(storageUsage)}
                </span>
              )}
            </div>
            {copiesError === null ? null : (
              <div
                className="mt-2 rounded-lg border border-fd-primary/40 p-3 text-sm"
                data-testid="workspace-copies-error"
                role="alert"
              >
                <p>
                  The working copies stored in this browser could not be listed
                </p>
                <pre className="mt-2 overflow-x-auto whitespace-pre-wrap text-xs leading-6">
                  {copiesError}
                </pre>
                <button
                  className={`${secondaryButtonClass} mt-3`}
                  data-testid="workspace-copies-retry"
                  disabled={disabled}
                  onClick={() => void refreshCopies()}
                  type="button"
                >
                  <RefreshCw aria-hidden="true" className="size-4" />
                  Try listing again
                </button>
              </div>
            )}
            {copies === null ? null : copies.databases.length === 0 &&
              copies.ignored.length === 0 ? (
              <p className="mt-2 text-sm text-fd-muted-foreground">
                No working copies are stored in this browser yet
              </p>
            ) : (
              <ul className="mt-2 space-y-2">
                {copies.databases.map((entry) => {
                  const isOpen = summary?.workingCopyName === entry.name;
                  const confirming = pendingDelete === entry.name;
                  return (
                    <li
                      className="flex flex-wrap items-center gap-2"
                      data-testid="workspace-copy"
                      key={entry.name}
                    >
                      <button
                        className={secondaryButtonClass}
                        data-testid="workspace-reopen"
                        disabled={disabled}
                        onClick={() => void reopen(entry.name)}
                        type="button"
                      >
                        <FolderOpen aria-hidden="true" className="size-4" />
                        {entry.name} ({entry.format})
                      </button>
                      {isOpen ? (
                        <>
                          <span className="text-xs text-fd-muted-foreground">
                            Open now
                          </span>
                          <button
                            className={compactButtonClass}
                            data-testid="workspace-close"
                            disabled={disabled}
                            onClick={() => void closeDatabase()}
                            type="button"
                          >
                            Close
                          </button>
                        </>
                      ) : confirming ? (
                        <>
                          <span className="text-xs">
                            Delete this copy from browser storage?
                          </span>
                          <button
                            className={compactButtonClass}
                            data-testid="workspace-copy-delete-confirm"
                            disabled={disabled}
                            onClick={() => void removeDatabase(entry.name)}
                            type="button"
                          >
                            <Trash2 aria-hidden="true" className="size-3.5" />
                            Confirm delete
                          </button>
                          <button
                            className={compactButtonClass}
                            data-testid="workspace-copy-delete-cancel"
                            onClick={() => setPendingDelete(null)}
                            type="button"
                          >
                            Keep
                          </button>
                        </>
                      ) : (
                        <button
                          className={compactButtonClass}
                          data-testid="workspace-copy-delete"
                          disabled={disabled}
                          onClick={() => setPendingDelete(entry.name)}
                          type="button"
                        >
                          <Trash2 aria-hidden="true" className="size-3.5" />
                          Delete
                        </button>
                      )}
                    </li>
                  );
                })}
                {copies.ignored.map((entry) => (
                  <li
                    className="text-xs text-fd-muted-foreground"
                    data-testid="workspace-copy-ignored"
                    key={`ignored:${entry.name}`}
                  >
                    {entry.name} ({entry.code}): {entry.message}
                  </li>
                ))}
              </ul>
            )}
          </div>
        )}
        {statusFor("start")}
      </section>

      {progress === null ? null : <ProgressNotice progress={progress} />}
      {busy === null || !cancellable ? null : (
        <button
          className={secondaryButtonClass}
          data-testid="workspace-cancel"
          onClick={() => controllerRef.current?.abort()}
          type="button"
        >
          <TriangleAlert aria-hidden="true" className="size-4" />
          Cancel {busy.toLowerCase()}
        </button>
      )}

      {summary === null ? (
        <section className={sectionClass} data-testid="workspace-empty">
          <p className="text-sm text-fd-muted-foreground">
            No database is open
          </p>
        </section>
      ) : (
        <>
          <Summary summary={summary} />
          <section className={sectionClass} data-testid="workspace-schema">
            <h2 className="font-display text-xl font-semibold">Schema</h2>
            <p className="mt-2 text-sm text-fd-muted-foreground">
              Review additive table and column changes before applying them.
              Conflicting types stay blocked
            </p>
            <textarea
              className={`${inputClass} mt-4 min-h-56 font-mono text-xs`}
              data-testid="workspace-schema-input"
              disabled={disabled}
              onChange={(event) => {
                const hadReview = schemaPlan !== null;
                setSchemaText(event.target.value);
                invalidateSchemaReview();
                if (hadReview) {
                  setStatus({
                    kind: "notice",
                    section: "schema",
                    message:
                      "Schema document changed. Review it again before applying",
                  });
                }
              }}
              spellCheck={false}
              value={schemaText}
            />
            <div className="mt-3 flex gap-3">
              <button
                className={secondaryButtonClass}
                data-testid="workspace-schema-plan"
                disabled={disabled}
                onClick={() => void planSchema()}
                type="button"
              >
                Review schema
              </button>
              <button
                className={primaryButtonClass}
                data-testid="workspace-schema-apply"
                disabled={disabled || schemaPlan?.ready !== true}
                onClick={() => void applySchema()}
                type="button"
              >
                Apply reviewed changes
              </button>
            </div>
            {schemaPlan === null ? null : (
              <WorkspaceSchemaReview plan={schemaPlan} />
            )}
            {statusFor("schema")}
          </section>

          <WorkspaceImport
            busy={disabled}
            client={client}
            key={workspaceGeneration}
            notice={statusFor("import")}
            summary={summary}
            onBatchRecorded={reloadDeliveries}
            onSummary={(next) => {
              setSummary(next);
              invalidateSchemaReview();
            }}
            onReviewState={setReviewActive}
            reportError={reportImportError}
            runLong={runImportOperation}
          />

          <section className={sectionClass} data-testid="workspace-deliveries">
            <div className="flex items-center justify-between gap-4">
              <div>
                <h2 className="font-display text-xl font-semibold">
                  Batch history
                </h2>
                <p className="mt-1 text-sm text-fd-muted-foreground">
                  Batch records stay separate from captured file contents
                </p>
              </div>
              <button
                className={secondaryButtonClass}
                data-testid="workspace-deliveries-refresh"
                disabled={disabled || deliveriesLoading}
                onClick={() => void loadDeliveries(null)}
                type="button"
              >
                <RefreshCw aria-hidden="true" className="size-4" />
                Refresh
              </button>
            </div>
            {deliveriesError === null ? null : (
              <div
                className="mt-4 rounded-lg border-2 border-fd-primary bg-fd-primary/10 px-4 py-3"
                data-testid="workspace-deliveries-error"
                role="alert"
              >
                <p className="text-sm font-semibold">
                  The batch history could not be loaded
                </p>
                <pre className="mt-2 overflow-x-auto whitespace-pre-wrap text-xs leading-6">
                  {deliveriesError}
                </pre>
              </div>
            )}
            {deliveries === null ? (
              deliveriesLoading ? (
                <p
                  className="mt-4 text-sm text-fd-muted-foreground"
                  data-testid="workspace-deliveries-loading"
                  role="status"
                >
                  Loading batch history
                </p>
              ) : null
            ) : deliveries.deliveries.length === 0 ? (
              <p className="mt-4 text-sm text-fd-muted-foreground">
                No batches recorded
              </p>
            ) : (
              <ol className="mt-4 space-y-3">
                {deliveries.deliveries.map((delivery) => (
                  <li
                    className="rounded-lg border p-4"
                    data-testid="workspace-delivery"
                    key={delivery.id}
                  >
                    <div className="flex justify-between gap-3">
                      <span className="font-semibold">{delivery.label}</span>
                      <span className="font-mono text-xs text-fd-muted-foreground">
                        {delivery.id}
                      </span>
                    </div>
                    <p className="mt-1 text-sm">
                      {delivery.vendor || "Unspecified vendor"} ·{" "}
                      {delivery.entity || "Unspecified entity"} ·{" "}
                      {delivery.phase || "Unspecified phase"} ·{" "}
                      {deliveryScopeText(delivery.scope)}
                    </p>
                    {delivery.reusedCapture ? (
                      <p className="mt-1 text-xs text-fd-muted-foreground">
                        Includes previously captured data
                      </p>
                    ) : null}
                  </li>
                ))}
              </ol>
            )}
            {deliveries?.nextCursor === null || deliveries === null ? null : (
              <button
                className={`${secondaryButtonClass} mt-4`}
                data-testid="workspace-deliveries-next"
                disabled={disabled || deliveriesLoading}
                onClick={() => void loadDeliveries(deliveries.nextCursor)}
                type="button"
              >
                Next batches
              </button>
            )}
            {statusFor("history")}
          </section>

          <section className={sectionClass} data-testid="workspace-export">
            <h2 className="font-display text-xl font-semibold">
              Export a portable copy
            </h2>
            <p className="mt-2 text-sm text-fd-muted-foreground">
              The worker checkpoints the working database and streams an
              independent file. Exporting does not move the working copy
            </p>
            <div className="mt-4 flex flex-wrap gap-3">
              <button
                className={primaryButtonClass}
                data-testid="workspace-export-same"
                disabled={disabled}
                onClick={() => void exportDatabase(summary.format)}
                type="button"
              >
                <Download aria-hidden="true" className="size-4" />
                Export {summary.format}
              </button>
              <button
                className={secondaryButtonClass}
                data-testid="workspace-export-convert"
                disabled={disabled}
                onClick={() =>
                  void exportDatabase(
                    summary.format === "sqlite" ? "duckdb" : "sqlite",
                  )
                }
                type="button"
              >
                <Database aria-hidden="true" className="size-4" />
                Convert to {summary.format === "sqlite" ? "DuckDB" : "SQLite"}
              </button>
            </div>
            {statusFor("export")}
          </section>
        </>
      )}
    </ToolShell>
  );
}
