"use client";

import {
  describeFailure,
  inputClass,
  primaryButtonClass,
  secondaryButtonClass,
  sectionClass,
  ToolShell,
} from "@/components/tool-kit";
import { WorkspaceImport } from "@/components/workspace-import";
import { isConsultChimpsError } from "@consultchimps/core";
import type {
  WorkspaceDatabaseFormat,
  WorkspaceDeliveryPage,
  WorkspaceProgress,
  WorkspaceSchemaDocument,
  WorkspaceSchemaPlan,
  WorkspaceSummary,
} from "@/lib/workspace-protocol";
import { WorkspaceClient } from "@/lib/workspace-worker";
import { removeOpfsFile } from "@/lib/workspace-files";
import {
  Database,
  Download,
  FilePlus,
  FolderOpen,
  LoaderCircle,
  RefreshCw,
  TriangleAlert,
} from "lucide-react";
import { useCallback, useEffect, useRef, useState } from "react";

const DATABASE_ACCEPT = ".sqlite,.sqlite3,.db,.duckdb";
const RECENT_DATABASES_KEY = "consultchimps.workspace.databases.v1";
const DEFAULT_SCHEMA = `{
  "version": 1,
  "tables": [
    {
      "name": "datasets",
      "recordId": { "prefix": "DATASET", "padding": 6 },
      "columns": [
        { "name": "dataset_name", "type": "text", "nullable": false },
        { "name": "reported_cde", "type": "boolean" }
      ]
    }
  ]
}`;

interface WorkspaceStatus {
  readonly kind: "error" | "notice";
  readonly message: string;
}

interface RecentDatabase {
  readonly name: string;
  readonly format: WorkspaceDatabaseFormat;
}

function readRecentDatabases(): RecentDatabase[] {
  try {
    const stored: unknown = JSON.parse(
      window.localStorage.getItem(RECENT_DATABASES_KEY) ?? "[]",
    );
    if (!Array.isArray(stored)) return [];
    return stored.flatMap((entry): RecentDatabase[] => {
      if (!isRecord(entry) || typeof entry["name"] !== "string") return [];
      const format = entry["format"];
      return format === "sqlite" || format === "duckdb"
        ? [{ name: entry["name"], format }]
        : [];
    });
  } catch {
    return [];
  }
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

function downloadFile(file: File, name: string): void {
  const url = URL.createObjectURL(file);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => {
    URL.revokeObjectURL(url);
    void removeOpfsFile(file.name).catch(() => undefined);
  }, 60_000);
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
          <dt className="text-fd-muted-foreground">Applied imports</dt>
          <dd>{summary.importCount}</dd>
        </div>
        <div>
          <dt className="text-fd-muted-foreground">Deliveries</dt>
          <dd>{summary.deliveryCount}</dd>
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

export function WorkspaceTool() {
  const clientRef = useRef<WorkspaceClient | null>(null);
  const openInputRef = useRef<HTMLInputElement | null>(null);
  const controllerRef = useRef<AbortController | null>(null);
  const [summary, setSummary] = useState<WorkspaceSummary | null>(null);
  const [format, setFormat] = useState<WorkspaceDatabaseFormat>("sqlite");
  const [name, setName] = useState("consultchimps.sqlite");
  const [busy, setBusy] = useState<string | null>(null);
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
  const [recentDatabases, setRecentDatabases] = useState<
    readonly RecentDatabase[]
  >([]);

  useEffect(() => {
    const timeout = window.setTimeout(() => {
      setRecentDatabases(readRecentDatabases());
    }, 0);
    return () => window.clearTimeout(timeout);
  }, []);

  const remember = useCallback((next: WorkspaceSummary) => {
    setRecentDatabases((current) => {
      const updated = [
        { name: next.workingCopyName, format: next.format },
        ...current.filter((entry) => entry.name !== next.workingCopyName),
      ].slice(0, 8);
      window.localStorage.setItem(
        RECENT_DATABASES_KEY,
        JSON.stringify(updated),
      );
      return updated;
    });
  }, []);

  const client = useCallback(() => {
    clientRef.current ??= new WorkspaceClient();
    return clientRef.current;
  }, []);

  useEffect(
    () => () => {
      controllerRef.current?.abort();
      clientRef.current?.terminate();
    },
    [],
  );

  const reportError = useCallback((error: unknown) => {
    setStatus({ kind: "error", message: describeFailure(error) });
  }, []);

  const runLong = useCallback(
    async <T,>(
      label: string,
      run: (options: {
        readonly signal: AbortSignal;
        readonly onProgress: (next: WorkspaceProgress) => void;
      }) => Promise<T>,
    ): Promise<T | null> => {
      const controller = new AbortController();
      controllerRef.current = controller;
      setBusy(label);
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
            kind: "notice",
            message:
              "Cancelled. The accepted database state was left unchanged",
          });
        } else {
          reportError(error);
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

  const create = useCallback(async () => {
    const created = await runLong("Creating database", (options) =>
      client().create(format, name.trim(), options),
    );
    if (created === null) return;
    setSummary(created);
    remember(created);
    setSchemaPlan(null);
    setDeliveries(null);
    setStatus({
      kind: "notice",
      message: "Created a persistent browser working database",
    });
  }, [client, format, name, remember, runLong]);

  const open = useCallback(
    async (file: File) => {
      const opened = await runLong(
        "Copying database into browser storage",
        (options) => client().open(file, options),
      );
      if (opened === null) return;
      setSummary(opened);
      remember(opened);
      setSchemaPlan(null);
      setDeliveries(null);
      setStatus({
        kind: "notice",
        message: `Opened "${file.name}" as a persistent browser working copy. The selected file will not change`,
      });
    },
    [client, remember, runLong],
  );

  const reopen = useCallback(
    async (workingCopyName: string) => {
      const opened = await runLong("Opening browser working copy", (options) =>
        client().reopen(workingCopyName, options),
      );
      if (opened === null) return;
      setSummary(opened);
      remember(opened);
      setSchemaPlan(null);
      setDeliveries(null);
      setStatus({
        kind: "notice",
        message: `Reopened "${workingCopyName}" from browser storage`,
      });
    },
    [client, remember, runLong],
  );

  const planSchema = useCallback(async () => {
    try {
      const plan = await client().planSchema(parseSchema(schemaText));
      setSchemaPlan(plan);
      setStatus({
        kind: "notice",
        message: plan.ready
          ? "Schema review is ready to apply"
          : "Schema conflicts need a decision before anything can change",
      });
    } catch (error) {
      reportError(error);
    }
  }, [client, reportError, schemaText]);

  const applySchema = useCallback(async () => {
    if (schemaPlan === null || !schemaPlan.ready) return;
    try {
      const next = await client().applySchema(schemaPlan.id);
      setSummary(next);
      setSchemaPlan(null);
      setStatus({
        kind: "notice",
        message: "Applied the reviewed schema changes",
      });
    } catch (error) {
      reportError(error);
    }
  }, [client, reportError, schemaPlan]);

  const loadDeliveries = useCallback(async () => {
    try {
      setDeliveries(await client().listDeliveries(null));
    } catch (error) {
      reportError(error);
    }
  }, [client, reportError]);

  const exportDatabase = useCallback(
    async (targetFormat: WorkspaceDatabaseFormat) => {
      const exported = await runLong(
        "Creating a consistent export",
        (options) => client().export(targetFormat, options),
      );
      if (exported === null) return;
      downloadFile(exported.file, exported.name);
      setStatus({
        kind: "notice",
        message: `Exported an independent ${exported.format} copy. Later browser changes will not update it`,
      });
    },
    [client, runLong],
  );

  const disabled = busy !== null;
  return (
    <ToolShell
      description="Create or open a persistent local database, review workbook imports, record deliveries, and export a portable copy"
      guideHref="/docs/getting-started"
      guideLabel="Read the getting started guide"
      kicker="Local database"
      title="Data workspace"
    >
      <section className={sectionClass} data-testid="workspace-start">
        <div className="grid gap-5 lg:grid-cols-[1fr_auto]">
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
          </div>
          <div className="flex flex-wrap items-end gap-3">
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
          unchanged
        </p>
        {recentDatabases.length === 0 ? null : (
          <div className="mt-5" data-testid="workspace-recent">
            <h3 className="text-sm font-semibold">Browser working copies</h3>
            <div className="mt-2 flex flex-wrap gap-2">
              {recentDatabases.map((entry) => (
                <button
                  className={secondaryButtonClass}
                  data-testid="workspace-reopen"
                  disabled={disabled}
                  key={entry.name}
                  onClick={() => void reopen(entry.name)}
                  type="button"
                >
                  <FolderOpen aria-hidden="true" className="size-4" />
                  {entry.name} ({entry.format})
                </button>
              ))}
            </div>
          </div>
        )}
      </section>

      {progress === null ? null : <ProgressNotice progress={progress} />}
      {busy === null ? null : (
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
              onChange={(event) => setSchemaText(event.target.value)}
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
              <div
                className="mt-4 rounded-lg border p-4"
                data-testid="workspace-schema-review"
              >
                <p>{schemaPlan.changes.length} proposed changes</p>
                {schemaPlan.conflicts.map((conflict) => (
                  <p
                    className="mt-2 text-sm text-fd-primary"
                    key={`${conflict.table}:${conflict.column ?? "table"}`}
                  >
                    {conflict.message}
                  </p>
                ))}
              </div>
            )}
          </section>

          <WorkspaceImport
            busy={disabled}
            client={client}
            summary={summary}
            onSummary={setSummary}
            onReviewState={setReviewActive}
            reportError={reportError}
            runLong={runLong}
          />

          <section className={sectionClass} data-testid="workspace-deliveries">
            <div className="flex items-center justify-between gap-4">
              <div>
                <h2 className="font-display text-xl font-semibold">
                  Delivery history
                </h2>
                <p className="mt-1 text-sm text-fd-muted-foreground">
                  Delivery events stay separate from captured file contents
                </p>
              </div>
              <button
                className={secondaryButtonClass}
                data-testid="workspace-deliveries-refresh"
                disabled={disabled}
                onClick={() => void loadDeliveries()}
                type="button"
              >
                <RefreshCw aria-hidden="true" className="size-4" />
                Refresh
              </button>
            </div>
            {deliveries === null ? null : deliveries.deliveries.length === 0 ? (
              <p className="mt-4 text-sm text-fd-muted-foreground">
                No deliveries recorded
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
                      <span className="font-semibold">
                        {delivery.vendor || "Unspecified vendor"}
                      </span>
                      <span className="font-mono text-xs text-fd-muted-foreground">
                        {delivery.id}
                      </span>
                    </div>
                    <p className="mt-1 text-sm">
                      {delivery.entity || "Unspecified entity"} ·{" "}
                      {delivery.phase || "Unspecified phase"} ·{" "}
                      {delivery.coverage} coverage
                    </p>
                    {delivery.reusedCapture ? (
                      <p className="mt-1 text-xs text-fd-muted-foreground">
                        Reused captured data without adding observation rows
                      </p>
                    ) : null}
                  </li>
                ))}
              </ol>
            )}
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
          </section>
        </>
      )}

      {status === null ? null : (
        <pre
          aria-live="polite"
          className="overflow-x-auto whitespace-pre-wrap rounded-lg border border-fd-primary/40 bg-fd-accent/30 px-4 py-3 text-xs leading-6"
          data-testid={
            status.kind === "error" ? "workspace-error" : "workspace-notice"
          }
          role={status.kind === "error" ? "alert" : "status"}
        >
          {status.message}
        </pre>
      )}
    </ToolShell>
  );
}
