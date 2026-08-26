"use client";

/**
 * Browser front ends for the byte-level PDF operations in
 * `@consultchimps/pdf/bytes`.
 *
 * The pages hold state and render; the actual page copying happens in the
 * shared operation worker, so a several-hundred-page split never freezes the
 * tab. See `tool-kit.tsx` for the shell these pages are assembled from.
 */

import {
  describeFailure,
  FilePicker,
  formatBytes,
  inputClass,
  noticeClass,
  readUploads,
  ResultsPanel,
  RunControls,
  sectionClass,
  ToolShell,
  useOperationRun,
  type UploadedFile,
  PREVIEW_DEBOUNCE_MS,
  compactButtonClass,
} from "@/components/tool-kit";
import { runOperation } from "@/lib/operation-worker";
import type { OperationPlan } from "@consultchimps/core";
// Type-only: the runtime module is loaded inside the worker.
import type { SplitPdfMetric } from "@consultchimps/pdf/bytes";
import { ArrowDown, ArrowUp, FileText, Trash2 } from "lucide-react";
import { useCallback, useEffect, useId, useState } from "react";

const PDF_MEDIA_TYPE = "application/pdf";
const PDF_ACCEPT = "application/pdf,.pdf";

function isPdfFile(file: File): boolean {
  return file.type === PDF_MEDIA_TYPE || /\.pdf$/iu.test(file.name);
}

export function PdfSplitTool() {
  const prefixId = useId();
  const previewHeadingId = useId();
  const [input, setInput] = useState<UploadedFile | null>(null);
  const [prefix, setPrefix] = useState("");
  const [plan, setPlan] = useState<OperationPlan<SplitPdfMetric> | null>(null);
  const [planError, setPlanError] = useState<string | null>(null);
  const runState = useOperationRun();
  const isRunning = runState.status === "running";

  // Planning re-reads the PDF, so it stays outside render and off the main
  // thread. Clearing a stale preview happens where the source file changes.
  useEffect(() => {
    if (!input) {
      return;
    }

    let active = true;
    const timer = window.setTimeout(() => {
      void (async () => {
        try {
          const nextPlan = await runOperation({
            kind: "pdf.plan-split",
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
    void runState.run({
      kind: "pdf.split",
      input: { bytes: input.bytes, name: input.name },
      filenamePrefix: prefix.trim() || undefined,
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
      <section className={sectionClass} data-testid="source-section">
        <h2 className="text-xl font-bold tracking-[-0.03em]">
          1. Choose a PDF
        </h2>
        <div className="mt-4">
          <FilePicker
            accept={PDF_ACCEPT}
            description="Drag a PDF here, or pick one with the button below. Only the first PDF is used."
            disabled={isRunning}
            label="Source PDF"
            multiple={false}
            onFiles={(files) => {
              void readUploads(files, isPdfFile).then((read) => {
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
          <p
            className="mt-3 flex items-center gap-2 text-sm text-fd-muted-foreground"
            data-testid="source-summary"
          >
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
            data-testid="prefix-input"
            disabled={isRunning}
            id={prefixId}
            onChange={(event) => setPrefix(event.target.value)}
            placeholder="report"
            type="text"
            value={prefix}
          />
        </div>
      </section>

      <section
        aria-labelledby={previewHeadingId}
        className={sectionClass}
        data-testid="preview-section"
      >
        <h2
          className="text-xl font-bold tracking-[-0.03em]"
          id={previewHeadingId}
        >
          2. Preview
        </h2>
        {!input ? (
          <p className="mt-3 text-sm text-fd-muted-foreground">
            Choose a PDF to see the pages it contains and the files this task
            will create.
          </p>
        ) : null}
        {planError ? (
          <pre className={noticeClass} data-testid="preview-error">
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
            <ul
              className="mt-2 max-h-56 overflow-y-auto rounded-lg border bg-fd-background/60 px-4 py-3 font-mono text-xs leading-6"
              data-testid="planned-outputs"
            >
              {plan.outputs.map((output) => (
                <li className="truncate" key={output.path}>
                  {output.path}
                </li>
              ))}
            </ul>
          </>
        ) : null}
      </section>

      <section className={sectionClass} data-testid="run-section">
        <h2 className="text-xl font-bold tracking-[-0.03em]">3. Run</h2>
        <RunControls
          busyLabel="Splitting…"
          disabled={!input}
          onCancel={runState.cancel}
          onRun={start}
          readingLabel="Reading the PDF…"
          runLabel="Run split"
          state={runState}
        />
      </section>

      <ResultsPanel
        archiveName={archiveName}
        fallbackMediaType={PDF_MEDIA_TYPE}
        state={runState}
      />
    </ToolShell>
  );
}

export function PdfMergeTool() {
  const outputNameId = useId();
  const [inputs, setInputs] = useState<readonly UploadedFile[]>([]);
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
    void runState.run({
      kind: "pdf.merge",
      inputs: inputs.map((file) => ({ bytes: file.bytes, name: file.name })),
      outputName: outputName.trim() || undefined,
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
      <section className={sectionClass} data-testid="source-section">
        <h2 className="text-xl font-bold tracking-[-0.03em]">1. Add PDFs</h2>
        <div className="mt-4">
          <FilePicker
            accept={PDF_ACCEPT}
            description="Drag one or more PDFs here, or pick them with the button below. Added files keep the order shown."
            disabled={isRunning}
            label="Source PDFs"
            multiple
            onFiles={(files) => {
              void readUploads(files, isPdfFile).then((read) => {
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
          <ol className="mt-4 flex flex-col gap-2" data-testid="source-list">
            {inputs.map((file, index) => (
              <li
                className="flex flex-wrap items-center justify-between gap-3 rounded-lg border bg-fd-background/60 px-3 py-2"
                data-testid="source-item"
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
            data-testid="output-name-input"
            disabled={isRunning}
            id={outputNameId}
            onChange={(event) => setOutputName(event.target.value)}
            placeholder="client-pack"
            type="text"
            value={outputName}
          />
        </div>
      </section>

      <section className={sectionClass} data-testid="run-section">
        <h2 className="text-xl font-bold tracking-[-0.03em]">2. Run</h2>
        <p className="mt-3 text-sm text-fd-muted-foreground">
          {inputs.length === 0
            ? "Add at least one PDF to merge."
            : `Every page from ${inputs.length} ${
                inputs.length === 1 ? "PDF" : "PDFs"
              } will be copied in the order listed above.`}
        </p>
        <RunControls
          busyLabel="Merging…"
          disabled={inputs.length === 0}
          onCancel={runState.cancel}
          onRun={start}
          readingLabel="Reading the PDFs…"
          runLabel="Run merge"
          state={runState}
        />
      </section>

      <ResultsPanel fallbackMediaType={PDF_MEDIA_TYPE} state={runState} />
    </ToolShell>
  );
}
