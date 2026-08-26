"use client";

/**
 * Browser front ends for the byte-level PowerPoint operations in
 * `@consultchimps/pptx/bytes`.
 *
 * As with the PDF and workbook pages, the components hold state and render;
 * reading the template package, parsing the records workbook, and rewriting
 * the presentation all happen in the shared operation worker. A deck with one
 * slide per record is the heaviest thing either tool does, so keeping it off
 * the main thread is what keeps the tab responsive on a long record set.
 *
 * Neither page uploads anything. The template and the records workbook are
 * read with the File API, populated in the worker, and handed back as a
 * download, which is the same operation the `pptx populate` command runs.
 */

import {
  describeFailure,
  FilePicker,
  formatBytes,
  inputClass,
  noticeClass,
  PREVIEW_DEBOUNCE_MS,
  readUploads,
  ResultsPanel,
  RunControls,
  sectionClass,
  ToolShell,
  useOperationRun,
  type UploadedFile,
} from "@/components/tool-kit";
import type { PresentationPopulateOptions } from "@/lib/operation-tasks";
import { runOperation } from "@/lib/operation-worker";
import type { OperationPlan } from "@consultchimps/core";
// Type-only: the runtime module is loaded inside the worker.
import type {
  PopulatePowerPointTemplatePlanMetric,
  PowerPointTemplateInspection,
} from "@consultchimps/pptx/bytes";
import { FileText } from "lucide-react";
import { useCallback, useEffect, useId, useMemo, useState } from "react";

const PRESENTATION_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const PRESENTATION_ACCEPT = `${PRESENTATION_MEDIA_TYPE},.pptx`;
const WORKBOOK_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const WORKBOOK_ACCEPT = `${WORKBOOK_MEDIA_TYPE},.xlsx`;

function isPresentationFile(file: File): boolean {
  return file.type === PRESENTATION_MEDIA_TYPE || /\.pptx$/iu.test(file.name);
}

function isWorkbookFile(file: File): boolean {
  return file.type === WORKBOOK_MEDIA_TYPE || /\.xlsx$/iu.test(file.name);
}

/**
 * Parse an optional one-based number field. A blank or unreadable field means
 * "not supplied", which lets the operation apply its own documented default
 * rather than having the page invent one.
 */
function optionalPositiveInteger(value: string): number | undefined {
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) && parsed >= 1 ? parsed : undefined;
}

const fieldLabelClass = "block text-sm font-semibold";
const fieldHintClass = "mt-1 text-sm text-fd-muted-foreground";
const metricLabelClass =
  "font-mono text-xs uppercase tracking-[0.12em] text-fd-muted-foreground";

interface TextFieldProps {
  readonly disabled: boolean;
  readonly hint: string;
  readonly label: string;
  readonly onChange: (value: string) => void;
  readonly placeholder: string;
  readonly testId: string;
  readonly value: string;
}

function TextField({
  disabled,
  hint,
  label,
  onChange,
  placeholder,
  testId,
  value,
}: TextFieldProps) {
  const fieldId = useId();
  return (
    <div>
      <label className={fieldLabelClass} htmlFor={fieldId}>
        {label}
      </label>
      <p className={fieldHintClass}>{hint}</p>
      <input
        className={`${inputClass} mt-2`}
        data-testid={testId}
        disabled={disabled}
        id={fieldId}
        onChange={(event) => onChange(event.target.value)}
        placeholder={placeholder}
        type="text"
        value={value}
      />
    </div>
  );
}

interface NumberFieldProps {
  readonly disabled: boolean;
  readonly hint: string;
  readonly label: string;
  readonly onChange: (value: string) => void;
  readonly placeholder: string;
  readonly testId: string;
  readonly value: string;
}

function NumberField({
  disabled,
  hint,
  label,
  onChange,
  placeholder,
  testId,
  value,
}: NumberFieldProps) {
  const fieldId = useId();
  return (
    <div>
      <label className={fieldLabelClass} htmlFor={fieldId}>
        {label}
      </label>
      <p className={fieldHintClass}>{hint}</p>
      <input
        className={`${inputClass} mt-2`}
        data-testid={testId}
        disabled={disabled}
        id={fieldId}
        min={1}
        onChange={(event) => onChange(event.target.value)}
        placeholder={placeholder}
        type="number"
        value={value}
      />
    </div>
  );
}

/** The chosen file's name and size, shown under whichever picker took it. */
function ChosenFile({
  file,
  testId,
}: {
  readonly file: UploadedFile;
  readonly testId: string;
}) {
  return (
    <p
      className="mt-3 flex items-center gap-2 text-sm text-fd-muted-foreground"
      data-testid={testId}
    >
      <FileText aria-hidden="true" className="size-4 shrink-0" />
      <span className="truncate font-mono">{file.name}</span>
      <span className="shrink-0">{formatBytes(file.bytes.byteLength)}</span>
    </p>
  );
}

export function PptxPopulateTool() {
  const previewHeadingId = useId();

  const [template, setTemplate] = useState<UploadedFile | null>(null);
  const [workbook, setWorkbook] = useState<UploadedFile | null>(null);
  const [worksheet, setWorksheet] = useState("");
  const [headerRow, setHeaderRow] = useState("");
  const [templateSlide, setTemplateSlide] = useState("");
  const [outputName, setOutputName] = useState("");
  const [plan, setPlan] =
    useState<OperationPlan<PopulatePowerPointTemplatePlanMetric> | null>(null);
  const [planError, setPlanError] = useState<string | null>(null);

  const runState = useOperationRun();
  const isRunning = runState.status === "running";

  const options = useMemo<PresentationPopulateOptions>(
    () => ({
      headerRow: optionalPositiveInteger(headerRow),
      outputName: outputName.trim() || undefined,
      templateSlide: optionalPositiveInteger(templateSlide),
      worksheet: worksheet.trim() || undefined,
    }),
    [headerRow, outputName, templateSlide, worksheet],
  );

  // Re-planning re-reads both packages, so wait for a pause in typing first.
  // Clearing a stale preview waits for the same pause, which keeps every
  // state change in this effect asynchronous.
  useEffect(() => {
    let active = true;
    const timer = window.setTimeout(() => {
      if (!template || !workbook) {
        setPlan(null);
        setPlanError(null);
        return;
      }
      void (async () => {
        try {
          const nextPlan = await runOperation({
            kind: "pptx.plan-populate",
            template: { bytes: template.bytes, name: template.name },
            workbook: { bytes: workbook.bytes, name: workbook.name },
            options,
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
  }, [options, template, workbook]);

  const start = useCallback(() => {
    if (!template || !workbook) {
      return;
    }
    void runState.run({
      kind: "pptx.populate",
      template: { bytes: template.bytes, name: template.name },
      workbook: { bytes: workbook.bytes, name: workbook.name },
      options,
    });
  }, [options, runState, template, workbook]);

  return (
    <ToolShell
      description="Choose a designed template slide and a workbook of records, and get one populated slide per record. Everything runs in this page using the same operation the ConsultChimps command line uses."
      guideHref="/docs/tools/powerpoint-populate"
      guideLabel="Read the PowerPoint guide"
      kicker="Online tool · PowerPoint populate"
      title="Populate a PowerPoint template"
    >
      <section className={sectionClass} data-testid="template-section">
        <h2 className="text-xl font-bold tracking-[-0.03em]">
          1. Choose a template
        </h2>
        <div className="mt-4">
          <FilePicker
            accept={PRESENTATION_ACCEPT}
            description="Drag a .pptx presentation here, or pick one with the button below. Its placeholders use the {{field_name}} syntax. Only the first presentation is used."
            disabled={isRunning}
            label="Template presentation"
            multiple={false}
            onFiles={(files) => {
              void readUploads(files, isPresentationFile).then((read) => {
                const [first] = read;
                if (first) {
                  setTemplate(first);
                  setPlan(null);
                  setPlanError(null);
                  runState.reset();
                }
              });
            }}
          />
        </div>
        {template ? (
          <ChosenFile file={template} testId="template-summary" />
        ) : null}
      </section>

      <section className={sectionClass} data-testid="records-section">
        <h2 className="text-xl font-bold tracking-[-0.03em]">
          2. Choose the records
        </h2>
        <p className="mt-3 text-sm text-fd-muted-foreground">
          Every nonempty row below the header row becomes one slide. Each
          template placeholder must exactly match a column header.
        </p>
        <div className="mt-4">
          <FilePicker
            accept={WORKBOOK_ACCEPT}
            description="Drag an .xlsx workbook here, or pick one with the button below. Only the first workbook is used."
            disabled={isRunning}
            label="Records workbook"
            multiple={false}
            onFiles={(files) => {
              void readUploads(files, isWorkbookFile).then((read) => {
                const [first] = read;
                if (first) {
                  setWorkbook(first);
                  setPlan(null);
                  setPlanError(null);
                  runState.reset();
                }
              });
            }}
          />
        </div>
        {workbook ? (
          <ChosenFile file={workbook} testId="records-summary" />
        ) : null}

        <div className="mt-6 flex flex-col gap-5">
          <TextField
            disabled={isRunning}
            hint="Optional. Defaults to `<template>-populated.pptx`. The `.pptx` extension is added for you."
            label="Output filename"
            onChange={setOutputName}
            placeholder="quarterly-review"
            testId="output-name-input"
            value={outputName}
          />
        </div>

        <details className="mt-6 rounded-lg border bg-fd-background/50 px-4 py-3">
          <summary className="cursor-pointer text-sm font-semibold">
            Advanced options
          </summary>
          <div className="mt-5 flex flex-col gap-5">
            <TextField
              disabled={isRunning}
              hint="Optional. Reads the records from one worksheet by name. Defaults to the first worksheet."
              label="Worksheet"
              onChange={setWorksheet}
              placeholder="Records"
              testId="worksheet-input"
              value={worksheet}
            />
            <NumberField
              disabled={isRunning}
              hint="Optional one-based row number. Defaults to the first nonempty row in the worksheet."
              label="Header row"
              onChange={setHeaderRow}
              placeholder="1"
              testId="header-row-input"
              value={headerRow}
            />
            <NumberField
              disabled={isRunning}
              hint="Optional one-based slide number. Defaults to the first slide in the template."
              label="Template slide"
              onChange={setTemplateSlide}
              placeholder="1"
              testId="template-slide-input"
              value={templateSlide}
            />
          </div>
        </details>
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
          3. Preview
        </h2>
        {!template || !workbook ? (
          <p className="mt-3 text-sm text-fd-muted-foreground">
            Choose a template and a records workbook to see the presentation
            this task will create.
          </p>
        ) : null}
        {planError ? (
          <pre className={noticeClass} data-testid="preview-error">
            {planError}
          </pre>
        ) : null}
        {plan && !planError ? (
          <>
            <dl className="mt-4 grid grid-cols-2 gap-4 sm:grid-cols-4">
              <div>
                <dt className={metricLabelClass}>Slides</dt>
                <dd className="text-2xl font-bold" data-testid="preview-slides">
                  {plan.metrics.generatedSlides}
                </dd>
              </div>
              <div>
                <dt className={metricLabelClass}>Rows read</dt>
                <dd className="text-2xl font-bold" data-testid="preview-rows">
                  {plan.metrics.inputRows}
                </dd>
              </div>
              <div>
                <dt className={metricLabelClass}>Placeholders</dt>
                <dd
                  className="text-2xl font-bold"
                  data-testid="preview-placeholders"
                >
                  {plan.metrics.placeholderFields}
                </dd>
              </div>
              <div>
                <dt className={metricLabelClass}>Rows skipped</dt>
                <dd
                  className="text-2xl font-bold"
                  data-testid="preview-skipped"
                >
                  {plan.metrics.skippedRows}
                </dd>
              </div>
            </dl>
            {plan.warnings.length > 0 ? (
              <ul className={noticeClass} data-testid="preview-warnings">
                {plan.warnings.map((warning) => (
                  <li key={warning}>{warning}</li>
                ))}
              </ul>
            ) : null}
            <p className="mt-5 text-sm font-semibold">Planned output name</p>
            <ul
              className="mt-2 rounded-lg border bg-fd-background/60 px-4 py-3 font-mono text-xs leading-6"
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
        <h2 className="text-xl font-bold tracking-[-0.03em]">4. Run</h2>
        <RunControls
          busyLabel="Populating…"
          disabled={!template || !workbook}
          onCancel={runState.cancel}
          onRun={start}
          readingLabel="Reading the template and the records…"
          runLabel="Run populate"
          state={runState}
        />
      </section>

      <ResultsPanel
        archiveName="populated-presentation.zip"
        fallbackMediaType={PRESENTATION_MEDIA_TYPE}
        state={runState}
      />
    </ToolShell>
  );
}

export function PptxInspectTool() {
  const reportHeadingId = useId();

  const [template, setTemplate] = useState<UploadedFile | null>(null);
  const [templateSlide, setTemplateSlide] = useState("");
  const [inspection, setInspection] =
    useState<PowerPointTemplateInspection | null>(null);
  const [inspectionError, setInspectionError] = useState<string | null>(null);

  const slideNumber = optionalPositiveInteger(templateSlide);

  // Inspection only reads; it writes nothing and produces no file, so there is
  // no run step to confirm. The report simply follows the chosen template and
  // slide, debounced so typing a slide number does not re-read the package on
  // every keystroke.
  useEffect(() => {
    let active = true;
    const timer = window.setTimeout(() => {
      if (!template) {
        setInspection(null);
        setInspectionError(null);
        return;
      }
      void (async () => {
        try {
          const report = await runOperation({
            kind: "pptx.inspect",
            template: { bytes: template.bytes, name: template.name },
            templateSlide: slideNumber,
          });
          if (active) {
            setInspection(report);
            setInspectionError(null);
          }
        } catch (error) {
          if (active) {
            setInspection(null);
            setInspectionError(describeFailure(error));
          }
        }
      })();
    }, PREVIEW_DEBOUNCE_MS);

    return () => {
      active = false;
      window.clearTimeout(timer);
    };
  }, [slideNumber, template]);

  return (
    <ToolShell
      description="Choose a template presentation and see every placeholder its slide expects, with occurrence counts, before you populate it. The template is read in this page and never uploaded."
      guideHref="/docs/tools/powerpoint-populate#inspect-the-template"
      guideLabel="Read the PowerPoint guide"
      kicker="Online tool · PowerPoint inspect"
      title="Inspect a PowerPoint template"
    >
      <section className={sectionClass} data-testid="source-section">
        <h2 className="text-xl font-bold tracking-[-0.03em]">
          1. Choose a template
        </h2>
        <div className="mt-4">
          <FilePicker
            accept={PRESENTATION_ACCEPT}
            description="Drag a .pptx presentation here, or pick one with the button below. Only the first presentation is used."
            disabled={false}
            label="Template presentation"
            multiple={false}
            onFiles={(files) => {
              void readUploads(files, isPresentationFile).then((read) => {
                const [first] = read;
                if (first) {
                  setTemplate(first);
                  setInspection(null);
                  setInspectionError(null);
                }
              });
            }}
          />
        </div>
        {template ? (
          <ChosenFile file={template} testId="source-summary" />
        ) : null}

        <div className="mt-6">
          <NumberField
            disabled={false}
            hint="Optional one-based slide number. Defaults to the first slide in the template."
            label="Template slide"
            onChange={setTemplateSlide}
            placeholder="1"
            testId="template-slide-input"
            value={templateSlide}
          />
        </div>
      </section>

      <section
        aria-labelledby={reportHeadingId}
        aria-live="polite"
        className={sectionClass}
        data-testid="inspection-section"
      >
        <h2
          className="text-xl font-bold tracking-[-0.03em]"
          id={reportHeadingId}
        >
          2. Placeholders
        </h2>
        {!template ? (
          <p className="mt-3 text-sm text-fd-muted-foreground">
            Choose a template to see the placeholders its slide expects.
          </p>
        ) : null}
        {inspectionError ? (
          <pre className={noticeClass} data-testid="inspection-error">
            {inspectionError}
          </pre>
        ) : null}
        {inspection && !inspectionError ? (
          <>
            <dl className="mt-4 grid grid-cols-2 gap-4 sm:grid-cols-4">
              <div>
                <dt className={metricLabelClass}>Slide</dt>
                <dd
                  className="text-2xl font-bold"
                  data-testid="inspection-slide"
                >
                  {inspection.slideNumber}
                </dd>
              </div>
              <div>
                <dt className={metricLabelClass}>Placeholders</dt>
                <dd
                  className="text-2xl font-bold"
                  data-testid="inspection-fields"
                >
                  {inspection.placeholders.length}
                </dd>
              </div>
              <div>
                <dt className={metricLabelClass}>Occurrences</dt>
                <dd
                  className="text-2xl font-bold"
                  data-testid="inspection-occurrences"
                >
                  {inspection.placeholderOccurrences}
                </dd>
              </div>
              <div>
                <dt className={metricLabelClass}>Malformed</dt>
                <dd
                  className="text-2xl font-bold"
                  data-testid="inspection-malformed"
                >
                  {inspection.malformedPlaceholderCount}
                </dd>
              </div>
            </dl>

            {inspection.placeholders.length > 0 ? (
              <ul
                className="mt-5 flex flex-col gap-2"
                data-testid="placeholder-list"
              >
                {inspection.placeholders.map((placeholder) => (
                  <li
                    className="flex flex-wrap items-center justify-between gap-3 rounded-lg border bg-fd-background/60 px-3 py-2"
                    data-testid="placeholder-item"
                    key={placeholder.name}
                  >
                    <span
                      className="truncate font-mono text-sm"
                      data-testid="placeholder-name"
                    >
                      {`{{${placeholder.name}}}`}
                    </span>
                    <span
                      className="shrink-0 text-xs text-fd-muted-foreground"
                      data-testid="placeholder-occurrences"
                    >
                      {placeholder.occurrences}{" "}
                      {placeholder.occurrences === 1
                        ? "occurrence"
                        : "occurrences"}
                    </span>
                  </li>
                ))}
              </ul>
            ) : (
              <p className="mt-5 text-sm text-fd-muted-foreground">
                This slide contains no {"{{field_name}}"} placeholders. A
                populate would refuse it.
              </p>
            )}

            {inspection.malformedPlaceholderCount > 0 ? (
              <p className={noticeClass} data-testid="malformed-warning">
                {inspection.malformedPlaceholderCount} shape
                {inspection.malformedPlaceholderCount === 1 ? "" : "s"} on this
                slide contain unbalanced braces. Use the exact{" "}
                {"{{field_name}}"} syntax, or a populate will refuse the
                template.
              </p>
            ) : null}

            {inspection.unsupportedPlacementPlaceholders.length > 0 ? (
              <p className={noticeClass} data-testid="unsupported-placement">
                Placeholders outside a supported text shape are not populated:{" "}
                {inspection.unsupportedPlacementPlaceholders.join(", ")}.
              </p>
            ) : null}

            {inspection.unsupportedSplitRunPlaceholders.length > 0 ? (
              <p className={noticeClass} data-testid="unsupported-split-run">
                Placeholders split across text runs are not populated:{" "}
                {inspection.unsupportedSplitRunPlaceholders.join(", ")}.
              </p>
            ) : null}
          </>
        ) : null}
      </section>
    </ToolShell>
  );
}
