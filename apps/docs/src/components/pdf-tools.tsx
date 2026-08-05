"use client";

/**
 * Browser front ends for the byte-level PDF operations in
 * `@consultchimps/pdf/bytes`. Everything here runs in the visitor's tab: files
 * are read with the File API, processed in memory, and offered back as
 * downloads. There is no upload, no API route, and no server component work,
 * which keeps the pages compatible with the site's static export.
 *
 * The heavy modules (`@consultchimps/pdf/bytes` and `jszip`) are pulled in with
 * dynamic `import()` on first use so that visiting a tool page does not
 * download a PDF engine before anyone asks for one.
 */

import {
  isConsultChimpsError,
  type ByteArtifact,
  type ByteOperationOutcome,
  type OperationPlan,
  type OperationProgress,
  type ProgressReporter,
} from "@consultchimps/core";
import {
  formatHumanError,
  formatHumanResult,
  GENERIC_VOCABULARY,
  type MessageVocabulary,
} from "@consultchimps/messages";
// Type-only: the runtime module is loaded on demand with dynamic import().
import type { SplitPdfMetric } from "@consultchimps/pdf/bytes";
import {
  ArrowDown,
  ArrowRight,
  ArrowUp,
  Ban,
  Download,
  FileArchive,
  FileText,
  LoaderCircle,
  ShieldCheck,
  Trash2,
} from "lucide-react";
import Link from "next/link";
import {
  useCallback,
  useEffect,
  useId,
  useRef,
  useState,
  type ReactNode,
} from "react";

const PDF_MEDIA_TYPE = "application/pdf";
const PDF_ACCEPT = "application/pdf,.pdf";
// Re-planning parses the PDF again, so wait for a pause in typing first.
const PREVIEW_DEBOUNCE_MS = 250;

/**
 * Browser wording for the shared explanations. Only the sentences that would
 * otherwise describe a terminal are replaced; everything else stays generic.
 */
const WEB_VOCABULARY: MessageVocabulary = {
  ...GENERIC_VOCABULARY,
  artifactListReference: "shown under Results",
  examplesReference:
    "Open the guide linked at the top of this page if you need examples.",
  inputFormatReference:
    "Open the guide linked at the top of this page to review the expected input format.",
  pdfOptionsReference:
    "Open the guide linked at the top of this page to see the available PDF options and examples.",
  retryAfterChoosingDifferentOutput:
    "Choose a different output name on this page and start the task again.",
  retryWhenReady: "Choose Run again when you are ready.",
};

interface PdfFile {
  readonly id: string;
  readonly name: string;
  readonly bytes: Uint8Array;
}

let fileCounter = 0;

function isPdfFile(file: File): boolean {
  return file.type === PDF_MEDIA_TYPE || /\.pdf$/iu.test(file.name);
}

async function readPdfFiles(files: readonly File[]): Promise<PdfFile[]> {
  const accepted = files.filter(isPdfFile);
  return Promise.all(
    accepted.map(async (file) => {
      fileCounter += 1;
      return {
        id: `pdf-${fileCounter}`,
        name: file.name,
        bytes: new Uint8Array(await file.arrayBuffer()),
      };
    }),
  );
}

function formatBytes(length: number): string {
  if (length < 1024) {
    return `${length} B`;
  }
  if (length < 1024 * 1024) {
    return `${(length / 1024).toFixed(1)} KB`;
  }
  return `${(length / (1024 * 1024)).toFixed(1)} MB`;
}

function describeFailure(error: unknown): string {
  if (isConsultChimpsError(error)) {
    return formatHumanError(error.message, error.code, {
      vocabulary: WEB_VOCABULARY,
    });
  }
  const message =
    error instanceof Error ? error.message : "An unexpected problem occurred.";
  return formatHumanError(message, undefined, { vocabulary: WEB_VOCABULARY });
}

function saveBlob(blob: Blob, name: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  // Revoke on a later task so the browser has started reading the blob.
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

function saveArtifact(artifact: ByteArtifact): void {
  // Copy into a fresh buffer so the blob never aliases the operation's memory.
  const copy = new Uint8Array(artifact.bytes.byteLength);
  copy.set(artifact.bytes);
  saveBlob(
    new Blob([copy.buffer], { type: artifact.mediaType ?? PDF_MEDIA_TYPE }),
    artifact.name,
  );
}

async function saveArchive(
  artifacts: readonly ByteArtifact[],
  archiveName: string,
): Promise<void> {
  const { default: JSZip } = await import("jszip");
  const zip = new JSZip();
  for (const artifact of artifacts) {
    zip.file(artifact.name, artifact.bytes);
  }
  saveBlob(await zip.generateAsync({ type: "blob" }), archiveName);
}

type RunStatus = "idle" | "running" | "complete" | "failed";

interface RunState {
  readonly status: RunStatus;
  readonly progress: OperationProgress | null;
  readonly outputs: readonly ByteArtifact[];
  readonly message: string;
}

const IDLE_RUN: RunState = {
  status: "idle",
  progress: null,
  outputs: [],
  message: "",
};

interface RunControls {
  readonly signal: AbortSignal;
  readonly onProgress: ProgressReporter;
}

/**
 * Drives one byte-level operation: progress reporting, cancellation through an
 * AbortController, and the plain-language rendering of the outcome.
 */
function useOperationRun(): RunState & {
  cancel: () => void;
  reset: () => void;
  run: <TMetric extends string>(
    execute: (controls: RunControls) => Promise<ByteOperationOutcome<TMetric>>,
  ) => Promise<void>;
} {
  const [state, setState] = useState<RunState>(IDLE_RUN);
  const controllerRef = useRef<AbortController | null>(null);

  useEffect(
    () => () => {
      controllerRef.current?.abort();
    },
    [],
  );

  const cancel = useCallback(() => {
    controllerRef.current?.abort();
  }, []);

  const reset = useCallback(() => {
    controllerRef.current?.abort();
    setState(IDLE_RUN);
  }, []);

  const run = useCallback(
    async <TMetric extends string>(
      execute: (
        controls: RunControls,
      ) => Promise<ByteOperationOutcome<TMetric>>,
    ): Promise<void> => {
      const controller = new AbortController();
      controllerRef.current = controller;
      setState({ ...IDLE_RUN, status: "running" });
      try {
        const outcome = await execute({
          signal: controller.signal,
          onProgress: (progress) => {
            setState((previous) =>
              previous.status === "running"
                ? { ...previous, progress }
                : previous,
            );
          },
        });
        setState({
          status: "complete",
          progress: null,
          outputs: outcome.outputs,
          message: formatHumanResult(outcome.result, {
            vocabulary: WEB_VOCABULARY,
          }),
        });
      } catch (error) {
        setState({
          status: "failed",
          progress: null,
          outputs: [],
          message: describeFailure(error),
        });
      } finally {
        controllerRef.current = null;
      }
    },
    [],
  );

  return { ...state, cancel, reset, run };
}

const primaryButtonClass =
  "inline-flex items-center justify-center gap-2 rounded-lg bg-fd-primary px-5 py-3 text-sm font-semibold text-fd-primary-foreground shadow-[3px_3px_0_var(--color-fd-foreground)] transition-transform hover:-translate-y-0.5 disabled:pointer-events-none disabled:opacity-50 disabled:shadow-none";
const secondaryButtonClass =
  "inline-flex items-center justify-center gap-2 rounded-lg border bg-fd-card px-5 py-3 text-sm font-semibold transition-colors hover:bg-fd-accent disabled:pointer-events-none disabled:opacity-50";
const compactButtonClass =
  "inline-flex items-center justify-center gap-1.5 rounded-md border bg-fd-card px-2.5 py-1.5 text-xs font-semibold transition-colors hover:bg-fd-accent disabled:pointer-events-none disabled:opacity-40";
const inputClass =
  "w-full rounded-lg border bg-fd-card px-3 py-2 text-sm outline-none focus-visible:ring-2 focus-visible:ring-fd-ring";
const sectionClass =
  "rounded-xl border bg-fd-card/80 p-6 shadow-[0_12px_36px_hsl(15_10%_11%/6%)]";
const kickerClass =
  "font-mono text-xs font-semibold uppercase tracking-[0.18em] text-fd-primary";

function PrivacyNotice() {
  return (
    <p className="flex items-start gap-3 rounded-xl border border-fd-primary/35 bg-fd-accent/40 px-4 py-3 text-sm font-medium text-fd-accent-foreground">
      <ShieldCheck className="mt-0.5 size-4 shrink-0" aria-hidden="true" />
      Your files never leave this browser tab — there is no server.
    </p>
  );
}

interface ToolShellProps {
  readonly children: ReactNode;
  readonly description: string;
  readonly guideHref: string;
  readonly guideLabel: string;
  readonly kicker: string;
  readonly title: string;
}

function ToolShell({
  children,
  description,
  guideHref,
  guideLabel,
  kicker,
  title,
}: ToolShellProps) {
  return (
    <main className="flex-1">
      <div className="mx-auto w-full max-w-[900px] px-6 pb-24 pt-16 lg:px-10 lg:pt-24">
        <div className={kickerClass}>{kicker}</div>
        <h1 className="mt-3 text-4xl font-bold tracking-[-0.05em] md:text-5xl">
          {title}
        </h1>
        <p className="mt-5 max-w-2xl text-lg leading-8 text-fd-muted-foreground">
          {description}
        </p>
        <div className="mt-6 flex flex-wrap gap-x-6 gap-y-2 text-sm font-semibold">
          <Link
            className="inline-flex items-center gap-1.5 text-fd-primary hover:underline"
            href={guideHref}
          >
            {guideLabel}
            <ArrowRight className="size-3.5" aria-hidden="true" />
          </Link>
        </div>
        <div className="mt-8">
          <PrivacyNotice />
        </div>
        <div className="mt-8 flex flex-col gap-6">{children}</div>
      </div>
    </main>
  );
}

interface FilePickerProps {
  readonly accept?: string;
  readonly description: string;
  readonly disabled: boolean;
  readonly label: string;
  readonly multiple: boolean;
  readonly onFiles: (files: File[]) => void;
}

function FilePicker({
  accept = PDF_ACCEPT,
  description,
  disabled,
  label,
  multiple,
  onFiles,
}: FilePickerProps) {
  const inputId = useId();
  const [isDragging, setIsDragging] = useState(false);

  return (
    <div
      className={`rounded-xl border-2 border-dashed p-6 transition-colors ${
        isDragging ? "border-fd-primary bg-fd-accent/40" : "bg-fd-card/60"
      }`}
      onDragEnter={(event) => {
        event.preventDefault();
        if (!disabled) {
          setIsDragging(true);
        }
      }}
      onDragOver={(event) => {
        event.preventDefault();
        if (!disabled) {
          setIsDragging(true);
        }
      }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={(event) => {
        event.preventDefault();
        setIsDragging(false);
        if (!disabled) {
          onFiles([...event.dataTransfer.files]);
        }
      }}
    >
      <label className="block text-sm font-semibold" htmlFor={inputId}>
        {label}
      </label>
      <p className="mt-1 text-sm text-fd-muted-foreground">{description}</p>
      <input
        accept={accept}
        className="mt-4 block w-full cursor-pointer text-sm text-fd-muted-foreground file:mr-4 file:cursor-pointer file:rounded-lg file:border-0 file:bg-fd-foreground file:px-4 file:py-2 file:text-sm file:font-semibold file:text-fd-background disabled:cursor-not-allowed disabled:opacity-50"
        disabled={disabled}
        id={inputId}
        multiple={multiple}
        onChange={(event) => {
          const selected = [...(event.target.files ?? [])];
          // Clear the control so re-picking the same file still fires change.
          event.target.value = "";
          onFiles(selected);
        }}
        type="file"
      />
    </div>
  );
}

function ProgressReport({
  progress,
}: {
  readonly progress: OperationProgress;
}) {
  const percent =
    progress.total > 0
      ? Math.min(100, Math.round((progress.completed / progress.total) * 100))
      : 0;

  return (
    <div className="mt-4">
      <div className="flex items-center justify-between font-mono text-xs uppercase tracking-[0.12em] text-fd-muted-foreground">
        <span>
          {progress.stage} · {progress.completed} of {progress.total}
        </span>
        <span>{percent}%</span>
      </div>
      <div
        aria-label="Operation progress"
        aria-valuemax={progress.total}
        aria-valuemin={0}
        aria-valuenow={progress.completed}
        className="mt-2 h-2 w-full overflow-hidden rounded-full bg-fd-muted"
        role="progressbar"
      >
        <div
          className="h-full rounded-full bg-fd-primary transition-[width] duration-150"
          style={{ width: `${percent}%` }}
        />
      </div>
      {progress.detail ? (
        <p className="mt-2 truncate font-mono text-xs text-fd-muted-foreground">
          {progress.detail}
        </p>
      ) : null}
    </div>
  );
}

interface ResultsPanelProps {
  readonly archiveName: string;
  readonly state: RunState;
}

function ResultsPanel({ archiveName, state }: ResultsPanelProps) {
  const [archiveError, setArchiveError] = useState<string | null>(null);
  const [isArchiving, setIsArchiving] = useState(false);

  if (state.status === "idle" || state.status === "running") {
    return null;
  }

  const failed = state.status === "failed";

  return (
    <section aria-live="polite" className={sectionClass}>
      <h2 className="text-xl font-bold tracking-[-0.03em]">Results</h2>

      {state.outputs.length > 0 ? (
        <>
          <ul className="mt-4 flex flex-col gap-2">
            {state.outputs.map((output) => (
              <li
                className="flex flex-wrap items-center justify-between gap-3 rounded-lg border bg-fd-background/60 px-3 py-2"
                key={output.name}
              >
                <span className="flex min-w-0 items-center gap-2">
                  <FileText
                    aria-hidden="true"
                    className="size-4 shrink-0 text-fd-muted-foreground"
                  />
                  <span className="truncate font-mono text-sm">
                    {output.name}
                  </span>
                  <span className="shrink-0 text-xs text-fd-muted-foreground">
                    {formatBytes(output.bytes.byteLength)}
                  </span>
                </span>
                <button
                  className={compactButtonClass}
                  onClick={() => saveArtifact(output)}
                  type="button"
                >
                  <Download className="size-3.5" aria-hidden="true" />
                  Download
                </button>
              </li>
            ))}
          </ul>
          <div className="mt-4">
            <button
              className={secondaryButtonClass}
              disabled={isArchiving}
              onClick={() => {
                setArchiveError(null);
                setIsArchiving(true);
                void saveArchive(state.outputs, archiveName)
                  .catch((error: unknown) => {
                    setArchiveError(describeFailure(error));
                  })
                  .finally(() => setIsArchiving(false));
              }}
              type="button"
            >
              <FileArchive className="size-4" aria-hidden="true" />
              {isArchiving ? "Building archive…" : "Download all (.zip)"}
            </button>
          </div>
        </>
      ) : null}

      <pre
        className={`mt-5 overflow-x-auto whitespace-pre-wrap rounded-lg border px-4 py-3 text-xs leading-6 ${
          failed
            ? "border-fd-primary/40 bg-fd-accent/30 text-fd-accent-foreground"
            : "bg-fd-background/60"
        }`}
      >
        {state.message}
      </pre>

      {archiveError ? (
        <pre className="mt-4 overflow-x-auto whitespace-pre-wrap rounded-lg border border-fd-primary/40 bg-fd-accent/30 px-4 py-3 text-xs leading-6 text-fd-accent-foreground">
          {archiveError}
        </pre>
      ) : null}
    </section>
  );
}

export function PdfSplitTool() {
  const prefixId = useId();
  const [input, setInput] = useState<PdfFile | null>(null);
  const [prefix, setPrefix] = useState("");
  const [plan, setPlan] = useState<OperationPlan<SplitPdfMetric> | null>(null);
  const [planError, setPlanError] = useState<string | null>(null);
  const runState = useOperationRun();
  const isRunning = runState.status === "running";

  // Planning re-reads the PDF, so it stays outside render. Clearing a stale
  // preview happens where the source file changes, not here.
  useEffect(() => {
    if (!input) {
      return;
    }

    let active = true;
    const timer = window.setTimeout(() => {
      void (async () => {
        try {
          const { planSplitPdfBytes } =
            await import("@consultchimps/pdf/bytes");
          const nextPlan = await planSplitPdfBytes({
            input: { bytes: input.bytes, name: input.name },
            filenamePrefix: prefix.trim() || undefined,
          });
          if (active) {
            setPlan(nextPlan);
            setPlanError(null);
          }
        } catch (error) {
          if (active) {
            setPlan(null);
            setPlanError(describeFailure(error));
          }
        }
      })();
    }, PREVIEW_DEBOUNCE_MS);

    return () => {
      active = false;
      window.clearTimeout(timer);
    };
  }, [input, prefix]);

  const start = useCallback(() => {
    if (!input) {
      return;
    }
    void runState.run(async ({ onProgress, signal }) => {
      const { splitPdfBytes } = await import("@consultchimps/pdf/bytes");
      return splitPdfBytes({
        filenamePrefix: prefix.trim() || undefined,
        input: { bytes: input.bytes, name: input.name },
        onProgress,
        signal,
      });
    });
  }, [input, prefix, runState]);

  const archiveName = `${(prefix.trim() || input?.name.replace(/\.pdf$/iu, "") || "document").replace(/[<>:"/\\|?*]+/gu, "-")}-pages.zip`;

  return (
    <ToolShell
      description="Drop a PDF below to write every page to its own file. The split runs in this page using the same operation the ConsultChimps command line uses."
      guideHref="/docs/tools/pdf-split"
      guideLabel="Read the split guide"
      kicker="Online tool · PDF split"
      title="Split a PDF"
    >
      <section className={sectionClass}>
        <h2 className="text-xl font-bold tracking-[-0.03em]">
          1. Choose a PDF
        </h2>
        <div className="mt-4">
          <FilePicker
            description="Drag a PDF here, or pick one with the button below. Only the first PDF is used."
            disabled={isRunning}
            label="Source PDF"
            multiple={false}
            onFiles={(files) => {
              void readPdfFiles(files).then((read) => {
                const [first] = read;
                if (first) {
                  setInput(first);
                  setPlan(null);
                  setPlanError(null);
                  runState.reset();
                }
              });
            }}
          />
        </div>
        {input ? (
          <p className="mt-3 flex items-center gap-2 text-sm text-fd-muted-foreground">
            <FileText aria-hidden="true" className="size-4 shrink-0" />
            <span className="truncate font-mono">{input.name}</span>
            <span className="shrink-0">
              {formatBytes(input.bytes.byteLength)}
            </span>
          </p>
        ) : null}

        <div className="mt-6">
          <label className="block text-sm font-semibold" htmlFor={prefixId}>
            Filename prefix
          </label>
          <p className="mt-1 text-sm text-fd-muted-foreground">
            Optional. Defaults to the source filename, so `report.pdf` produces
            `report-page-001.pdf`.
          </p>
          <input
            className={`${inputClass} mt-3`}
            disabled={isRunning}
            id={prefixId}
            onChange={(event) => setPrefix(event.target.value)}
            placeholder="report"
            type="text"
            value={prefix}
          />
        </div>
      </section>

      <section className={sectionClass}>
        <h2 className="text-xl font-bold tracking-[-0.03em]">2. Preview</h2>
        {!input ? (
          <p className="mt-3 text-sm text-fd-muted-foreground">
            Choose a PDF to see the pages it contains and the files this task
            will create.
          </p>
        ) : null}
        {planError ? (
          <pre className="mt-4 overflow-x-auto whitespace-pre-wrap rounded-lg border border-fd-primary/40 bg-fd-accent/30 px-4 py-3 text-xs leading-6 text-fd-accent-foreground">
            {planError}
          </pre>
        ) : null}
        {plan && !planError ? (
          <>
            <dl className="mt-4 grid grid-cols-2 gap-4 sm:grid-cols-3">
              <div>
                <dt className="font-mono text-xs uppercase tracking-[0.12em] text-fd-muted-foreground">
                  Pages
                </dt>
                <dd className="text-2xl font-bold">{plan.metrics.pages}</dd>
              </div>
              <div>
                <dt className="font-mono text-xs uppercase tracking-[0.12em] text-fd-muted-foreground">
                  Files created
                </dt>
                <dd className="text-2xl font-bold">
                  {plan.metrics.outputFiles}
                </dd>
              </div>
              <div>
                <dt className="font-mono text-xs uppercase tracking-[0.12em] text-fd-muted-foreground">
                  Operation
                </dt>
                <dd className="font-mono text-sm leading-9">
                  {plan.operation}
                </dd>
              </div>
            </dl>
            <p className="mt-5 text-sm font-semibold">Planned output names</p>
            <ul className="mt-2 max-h-56 overflow-y-auto rounded-lg border bg-fd-background/60 px-4 py-3 font-mono text-xs leading-6">
              {plan.outputs.map((output) => (
                <li className="truncate" key={output.path}>
                  {output.path}
                </li>
              ))}
            </ul>
          </>
        ) : null}
      </section>

      <section className={sectionClass}>
        <h2 className="text-xl font-bold tracking-[-0.03em]">3. Run</h2>
        <div className="mt-4 flex flex-wrap gap-3">
          <button
            className={primaryButtonClass}
            disabled={!input || isRunning}
            onClick={start}
            type="button"
          >
            {isRunning ? (
              <LoaderCircle
                aria-hidden="true"
                className="size-4 animate-spin"
              />
            ) : null}
            {isRunning ? "Splitting…" : "Run split"}
          </button>
          <button
            className={secondaryButtonClass}
            disabled={!isRunning}
            onClick={runState.cancel}
            type="button"
          >
            <Ban className="size-4" aria-hidden="true" />
            Cancel
          </button>
        </div>
        {runState.progress ? (
          <ProgressReport progress={runState.progress} />
        ) : null}
        {isRunning && !runState.progress ? (
          <p className="mt-4 text-sm text-fd-muted-foreground">
            Reading the PDF…
          </p>
        ) : null}
      </section>

      <ResultsPanel archiveName={archiveName} state={runState} />
    </ToolShell>
  );
}

export function PdfMergeTool() {
  const outputNameId = useId();
  const [inputs, setInputs] = useState<readonly PdfFile[]>([]);
  const [outputName, setOutputName] = useState("");
  const runState = useOperationRun();
  const isRunning = runState.status === "running";

  const move = useCallback((index: number, offset: number) => {
    setInputs((previous) => {
      const target = index + offset;
      if (target < 0 || target >= previous.length) {
        return previous;
      }
      const next = [...previous];
      const moved = next[index];
      const displaced = next[target];
      if (!moved || !displaced) {
        return previous;
      }
      next[index] = displaced;
      next[target] = moved;
      return next;
    });
  }, []);

  const start = useCallback(() => {
    if (inputs.length === 0) {
      return;
    }
    void runState.run(async ({ onProgress, signal }) => {
      const { mergePdfsBytes } = await import("@consultchimps/pdf/bytes");
      return mergePdfsBytes({
        inputs: inputs.map((file) => ({ bytes: file.bytes, name: file.name })),
        onProgress,
        outputName: outputName.trim() || undefined,
        signal,
      });
    });
  }, [inputs, outputName, runState]);

  return (
    <ToolShell
      description="Add the PDFs you want to combine, arrange them in the order they should appear, and merge them without uploading anything."
      guideHref="/docs/tools/pdf-merge"
      guideLabel="Read the merge guide"
      kicker="Online tool · PDF merge"
      title="Merge PDFs"
    >
      <section className={sectionClass}>
        <h2 className="text-xl font-bold tracking-[-0.03em]">1. Add PDFs</h2>
        <div className="mt-4">
          <FilePicker
            description="Drag one or more PDFs here, or pick them with the button below. Added files keep the order shown."
            disabled={isRunning}
            label="Source PDFs"
            multiple
            onFiles={(files) => {
              void readPdfFiles(files).then((read) => {
                if (read.length > 0) {
                  setInputs((previous) => [...previous, ...read]);
                  runState.reset();
                }
              });
            }}
          />
        </div>

        {inputs.length === 0 ? (
          <p className="mt-4 text-sm text-fd-muted-foreground">
            No PDFs added yet.
          </p>
        ) : (
          <ol className="mt-4 flex flex-col gap-2">
            {inputs.map((file, index) => (
              <li
                className="flex flex-wrap items-center justify-between gap-3 rounded-lg border bg-fd-background/60 px-3 py-2"
                key={file.id}
              >
                <span className="flex min-w-0 items-center gap-2">
                  <span className="shrink-0 font-mono text-xs text-fd-muted-foreground">
                    {String(index + 1).padStart(2, "0")}
                  </span>
                  <span className="truncate font-mono text-sm">
                    {file.name}
                  </span>
                  <span className="shrink-0 text-xs text-fd-muted-foreground">
                    {formatBytes(file.bytes.byteLength)}
                  </span>
                </span>
                <span className="flex shrink-0 items-center gap-2">
                  <button
                    aria-label={`Move ${file.name} earlier`}
                    className={compactButtonClass}
                    disabled={index === 0 || isRunning}
                    onClick={() => move(index, -1)}
                    type="button"
                  >
                    <ArrowUp className="size-3.5" aria-hidden="true" />
                  </button>
                  <button
                    aria-label={`Move ${file.name} later`}
                    className={compactButtonClass}
                    disabled={index === inputs.length - 1 || isRunning}
                    onClick={() => move(index, 1)}
                    type="button"
                  >
                    <ArrowDown className="size-3.5" aria-hidden="true" />
                  </button>
                  <button
                    aria-label={`Remove ${file.name}`}
                    className={compactButtonClass}
                    disabled={isRunning}
                    onClick={() =>
                      setInputs((previous) =>
                        previous.filter((entry) => entry.id !== file.id),
                      )
                    }
                    type="button"
                  >
                    <Trash2 className="size-3.5" aria-hidden="true" />
                  </button>
                </span>
              </li>
            ))}
          </ol>
        )}

        <div className="mt-6">
          <label className="block text-sm font-semibold" htmlFor={outputNameId}>
            Output filename
          </label>
          <p className="mt-1 text-sm text-fd-muted-foreground">
            Optional. Defaults to `combined.pdf`. The `.pdf` extension is added
            for you.
          </p>
          <input
            className={`${inputClass} mt-3`}
            disabled={isRunning}
            id={outputNameId}
            onChange={(event) => setOutputName(event.target.value)}
            placeholder="client-pack"
            type="text"
            value={outputName}
          />
        </div>
      </section>

      <section className={sectionClass}>
        <h2 className="text-xl font-bold tracking-[-0.03em]">2. Run</h2>
        <p className="mt-3 text-sm text-fd-muted-foreground">
          {inputs.length === 0
            ? "Add at least one PDF to merge."
            : `Every page from ${inputs.length} ${
                inputs.length === 1 ? "PDF" : "PDFs"
              } will be copied in the order listed above.`}
        </p>
        <div className="mt-4 flex flex-wrap gap-3">
          <button
            className={primaryButtonClass}
            disabled={inputs.length === 0 || isRunning}
            onClick={start}
            type="button"
          >
            {isRunning ? (
              <LoaderCircle
                aria-hidden="true"
                className="size-4 animate-spin"
              />
            ) : null}
            {isRunning ? "Merging…" : "Run merge"}
          </button>
          <button
            className={secondaryButtonClass}
            disabled={!isRunning}
            onClick={runState.cancel}
            type="button"
          >
            <Ban className="size-4" aria-hidden="true" />
            Cancel
          </button>
        </div>
        {runState.progress ? (
          <ProgressReport progress={runState.progress} />
        ) : null}
        {isRunning && !runState.progress ? (
          <p className="mt-4 text-sm text-fd-muted-foreground">
            Reading the PDFs…
          </p>
        ) : null}
      </section>

      <ResultsPanel archiveName="merged-pdf.zip" state={runState} />
    </ToolShell>
  );
}
