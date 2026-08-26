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
  PresentationInspectionOutcome,
} from "@consultchimps/pptx/bytes";
import { FileText } from "lucide-react";
import {
  useCallback,
  useEffect,
  useId,
  useMemo,
  useRef,
  useState,
} from "react";

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
 * A one-based number field's three states. A field is either not supplied at
 * all — the operation then applies its own documented default — or supplied
 * with a usable value, or supplied with something the operation cannot honour.
 *
 * The third state has to be distinct from the first. Reading `0`, `1.5`, or
 * `1e2` as "not supplied" would silently populate or inspect a different row
 * or slide than the one that was typed, and reading them with `parseInt` would
 * silently truncate `1.5` and `1e2` to slide 1. Both are the kind of quiet
 * substitution the toolkit refuses to make on a user's documents, so an
 * unusable value stops the preview and the run and says so on the field.
 */
type NumberFieldState =
  | { readonly kind: "empty" }
  | { readonly kind: "invalid"; readonly message: string }
  | { readonly kind: "value"; readonly value: number };

/**
 * Parse an optional one-based number field, accepting only whole numbers
 * counted from 1 — the same values the `--template-slide` and `--header-row`
 * command-line options accept.
 */
function positiveIntegerField(raw: string, label: string): NumberFieldState {
  const text = raw.trim();
  if (text === "") {
    return { kind: "empty" };
  }
  // Digits only: this rejects "1.5" and "1e2", which Number.parseInt would
  // quietly read as 1, along with signs, spaces, and hexadecimal.
  if (!/^\d+$/u.test(text)) {
    return {
      kind: "invalid",
      message: `${label} must be a whole number counted from 1, so "${text}" cannot be used.`,
    };
  }
  const value = Number(text);
  if (value < 1 || !Number.isSafeInteger(value)) {
    return {
      kind: "invalid",
      message: `${label} is counted from 1, so ${text} cannot be used.`,
    };
  }
  return { kind: "value", value };
}

/** The value to send, or undefined when the field was left blank. */
function suppliedNumber(state: NumberFieldState): number | undefined {
  return state.kind === "value" ? state.value : undefined;
}

function fieldMessage(state: NumberFieldState): string | undefined {
  return state.kind === "invalid" ? state.message : undefined;
}

interface FileSelection {
  /** The chosen file, or null while nothing usable is selected. */
  readonly file: UploadedFile | null;
  /** Set when the last pick was rejected; cleared by the next one. */
  readonly rejected: string | null;
  /** True between choosing a file and finishing its read. */
  readonly reading: boolean;
  readonly choose: (files: readonly File[], onChange: () => void) => void;
}

/**
 * One picker's selection, held so that at no instant does the page offer to
 * run against a document the visitor is not looking at.
 *
 * Reading an upload is asynchronous, and both hazards that follow from that
 * are handled here rather than at each call site:
 *
 * - The moment a pick starts, the previous selection is cleared. A large or
 *   cloud-backed replacement can take a noticeable time to read, and leaving
 *   the old file live for that window would let Run populate the very
 *   document that was just replaced.
 * - A read applies only while it is still the newest. Picking twice in quick
 *   succession leaves two reads in flight, and a big deck picked first can
 *   finish after a small one picked second; without the token the older read
 *   would win.
 */
function useFileSelection(
  accepts: (file: File) => boolean,
  expected: string,
): FileSelection {
  const [file, setFile] = useState<UploadedFile | null>(null);
  const [rejected, setRejected] = useState<string | null>(null);
  const [reading, setReading] = useState(false);
  const latest = useRef(0);

  const choose = useCallback(
    (files: readonly File[], onChange: () => void) => {
      const token = (latest.current += 1);
      setFile(null);
      setRejected(null);
      setReading(true);
      onChange();

      void readUploads(files, accepts)
        .then((read) => {
          if (token !== latest.current) {
            return;
          }
          const [first] = read;
          setFile(first ?? null);
          setRejected(first ? null : rejectedUploadMessage(files, expected));
          setReading(false);
        })
        .catch(() => {
          // A cloud-backed or removable file can go unreadable mid-read.
          // Without this the picker would sit in its reading state forever,
          // with nothing selected and Run disabled, and say nothing about why.
          if (token !== latest.current) {
            return;
          }
          setFile(null);
          setRejected(
            "That file could not be read. It may have moved, gone offline, or been removed. Choose it again, or pick another file.",
          );
          setReading(false);
        });
    },
    [accepts, expected],
  );

  return { choose, file, reading, rejected };
}

/** Shown in place of the file summary while a pick is still being read. */
function ReadingFile({ testId }: { readonly testId: string }) {
  return (
    <p
      className="mt-3 text-sm text-fd-muted-foreground"
      data-testid={testId}
      role="status"
    >
      Reading the file…
    </p>
  );
}

/**
 * An answer together with the inputs it was computed from.
 *
 * Both pages recompute in the background after a debounce, and Run applies a
 * changed option the moment it is typed. Rendering a stored answer only while
 * its `key` still matches the page is what stops a preview from describing
 * one presentation while the button would produce another, and stops the
 * inspector from labelling slide 2 while still listing slide 1.
 */
interface PlannedPopulate {
  readonly error: string | null;
  readonly key: string;
  readonly plan: OperationPlan<PopulatePowerPointTemplatePlanMetric> | null;
}

interface InspectedTemplate {
  readonly error: string | null;
  readonly key: string;
  readonly outcome: PresentationInspectionOutcome | null;
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
  /** Set when the typed value cannot be used; shown under the field. */
  readonly error: string | undefined;
  readonly hint: string;
  readonly label: string;
  readonly onChange: (value: string) => void;
  readonly placeholder: string;
  readonly testId: string;
  readonly value: string;
}

function NumberField({
  disabled,
  error,
  hint,
  label,
  onChange,
  placeholder,
  testId,
  value,
}: NumberFieldProps) {
  const fieldId = useId();
  const errorId = useId();
  return (
    <div>
      <label className={fieldLabelClass} htmlFor={fieldId}>
        {label}
      </label>
      <p className={fieldHintClass}>{hint}</p>
      <input
        aria-describedby={error ? errorId : undefined}
        aria-invalid={error ? true : undefined}
        className={`${inputClass} mt-2`}
        data-testid={testId}
        disabled={disabled}
        id={fieldId}
        min={1}
        onChange={(event) => onChange(event.target.value)}
        placeholder={placeholder}
        step={1}
        type="number"
        value={value}
      />
      {error ? (
        <p
          className={noticeClass}
          data-testid={`${testId}-error`}
          id={errorId}
          role="alert"
        >
          {error}
        </p>
      ) : null}
    </div>
  );
}

/**
 * The message shown when a picker rejects everything it was given.
 *
 * `readUploads` drops files the tool cannot process and says nothing, which is
 * right when nothing was chosen yet. It is not right when a document is
 * already selected: silently keeping it would let someone who meant to replace
 * a template populate the old one instead. Both pages therefore clear the
 * selection and say why, so a run can only ever use a file that was chosen on
 * purpose.
 */
function rejectedUploadMessage(
  files: readonly File[],
  expected: string,
): string {
  const [first] = files;
  const named = first ? `"${first.name}" is not` : "Those files are not";
  return `${named} ${expected}. Nothing is selected now, so choose ${expected} to continue.`;
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

  const templateSelection = useFileSelection(
    isPresentationFile,
    "a PowerPoint .pptx presentation",
  );
  const recordsSelection = useFileSelection(
    isWorkbookFile,
    "an Excel .xlsx workbook",
  );
  const template = templateSelection.file;
  const workbook = recordsSelection.file;
  const [worksheet, setWorksheet] = useState("");
  const [headerRow, setHeaderRow] = useState("");
  const [templateSlide, setTemplateSlide] = useState("");
  const [outputName, setOutputName] = useState("");
  const [planned, setPlanned] = useState<PlannedPopulate | null>(null);

  const runState = useOperationRun();
  const isRunning = runState.status === "running";

  const headerRowField = positiveIntegerField(headerRow, "Header row");
  const templateSlideField = positiveIntegerField(
    templateSlide,
    "Template slide",
  );
  // An unusable number must stop the task rather than fall back to a default,
  // so neither the preview nor the run reads a row or slide nobody asked for.
  const hasUnusableNumber =
    headerRowField.kind === "invalid" || templateSlideField.kind === "invalid";

  // Parsed again from the raw text rather than reusing the states above so
  // every dependency here is a plain string straight from useState.
  const options = useMemo<PresentationPopulateOptions>(
    () => ({
      headerRow: suppliedNumber(positiveIntegerField(headerRow, "Header row")),
      outputName: outputName.trim() || undefined,
      templateSlide: suppliedNumber(
        positiveIntegerField(templateSlide, "Template slide"),
      ),
      worksheet: worksheet.trim() || undefined,
    }),
    [headerRow, outputName, templateSlide, worksheet],
  );

  // Everything a plan depends on, in one comparable value. A stored plan is
  // shown only while its key still matches what is on the page: Run applies a
  // changed option immediately, so a preview computed from the previous
  // options would describe a different presentation than the button produces.
  // Keyed on the raw field text, not on the parsed options: "" and "0" both
  // parse to "no template slide supplied", yet one is a plan worth showing and
  // the other is a value the page refuses, so they must not share a key.
  const previewKey = JSON.stringify([
    template?.id ?? null,
    workbook?.id ?? null,
    worksheet.trim(),
    headerRow.trim(),
    templateSlide.trim(),
    outputName.trim(),
  ]);
  const current = planned?.key === previewKey ? planned : null;
  const readyToPlan =
    template !== null && workbook !== null && !hasUnusableNumber;

  // Re-planning re-reads both packages, so wait for a pause in typing first.
  useEffect(() => {
    if (!template || !workbook || hasUnusableNumber) {
      return;
    }
    let active = true;
    // Planning reads both packages in full, so a superseded attempt is
    // cancelled rather than merely ignored: several abandoned parses queued
    // behind each other would make the run the visitor actually wants wait.
    const controller = new AbortController();
    const timer = window.setTimeout(() => {
      void (async () => {
        try {
          const plan = await runOperation(
            {
              kind: "pptx.plan-populate",
              template: { bytes: template.bytes, name: template.name },
              workbook: { bytes: workbook.bytes, name: workbook.name },
              options,
            },
            { signal: controller.signal },
          );
          if (active) {
            setPlanned({ error: null, key: previewKey, plan });
          }
        } catch (error) {
          if (active) {
            setPlanned({
              error: describeFailure(error),
              key: previewKey,
              plan: null,
            });
          }
        }
      })();
    }, PREVIEW_DEBOUNCE_MS);

    return () => {
      active = false;
      window.clearTimeout(timer);
      controller.abort();
    };
  }, [hasUnusableNumber, options, previewKey, template, workbook]);

  const start = useCallback(() => {
    if (!template || !workbook || hasUnusableNumber) {
      return;
    }
    void runState.run({
      kind: "pptx.populate",
      template: { bytes: template.bytes, name: template.name },
      workbook: { bytes: workbook.bytes, name: workbook.name },
      options,
    });
  }, [hasUnusableNumber, options, runState, template, workbook]);

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
              templateSelection.choose(files, () => {
                setPlanned(null);
                runState.reset();
              });
            }}
          />
        </div>
        {templateSelection.reading ? (
          <ReadingFile testId="template-reading" />
        ) : null}
        {template ? (
          <ChosenFile file={template} testId="template-summary" />
        ) : null}
        {templateSelection.rejected ? (
          <p
            className={noticeClass}
            data-testid="template-rejected"
            role="alert"
          >
            {templateSelection.rejected}
          </p>
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
              recordsSelection.choose(files, () => {
                setPlanned(null);
                runState.reset();
              });
            }}
          />
        </div>
        {recordsSelection.reading ? (
          <ReadingFile testId="records-reading" />
        ) : null}
        {workbook ? (
          <ChosenFile file={workbook} testId="records-summary" />
        ) : null}
        {recordsSelection.rejected ? (
          <p
            className={noticeClass}
            data-testid="records-rejected"
            role="alert"
          >
            {recordsSelection.rejected}
          </p>
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
              error={fieldMessage(headerRowField)}
              hint="Optional one-based row number. Defaults to the first nonempty row in the worksheet."
              label="Header row"
              onChange={setHeaderRow}
              placeholder="1"
              testId="header-row-input"
              value={headerRow}
            />
            <NumberField
              disabled={isRunning}
              error={fieldMessage(templateSlideField)}
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
        {/*
          Repeated here because the advanced options can be collapsed: a
          disabled Run button with its explanation hidden behind a disclosure
          would leave no way to tell what is wrong.
        */}
        {hasUnusableNumber ? (
          <ul className={noticeClass} data-testid="preview-invalid-options">
            {[headerRowField, templateSlideField]
              .map(fieldMessage)
              .filter((message) => message !== undefined)
              .map((message) => (
                <li key={message}>{message}</li>
              ))}
          </ul>
        ) : null}
        {readyToPlan && current === null ? (
          <p
            className="mt-3 text-sm text-fd-muted-foreground"
            data-testid="preview-pending"
          >
            Working out what this task will create…
          </p>
        ) : null}
        {current?.error ? (
          <pre className={noticeClass} data-testid="preview-error">
            {current.error}
          </pre>
        ) : null}
        {current?.plan ? (
          <>
            <dl className="mt-4 grid grid-cols-2 gap-4 sm:grid-cols-4">
              <div>
                <dt className={metricLabelClass}>Slides</dt>
                <dd className="text-2xl font-bold" data-testid="preview-slides">
                  {current.plan.metrics.generatedSlides}
                </dd>
              </div>
              <div>
                <dt className={metricLabelClass}>Rows read</dt>
                <dd className="text-2xl font-bold" data-testid="preview-rows">
                  {current.plan.metrics.inputRows}
                </dd>
              </div>
              <div>
                <dt className={metricLabelClass}>Placeholders</dt>
                <dd
                  className="text-2xl font-bold"
                  data-testid="preview-placeholders"
                >
                  {current.plan.metrics.placeholderFields}
                </dd>
              </div>
              <div>
                <dt className={metricLabelClass}>Rows skipped</dt>
                <dd
                  className="text-2xl font-bold"
                  data-testid="preview-skipped"
                >
                  {current.plan.metrics.skippedRows}
                </dd>
              </div>
            </dl>
            {current.plan.warnings.length > 0 ? (
              <ul className={noticeClass} data-testid="preview-warnings">
                {current.plan.warnings.map((warning) => (
                  <li key={warning}>{warning}</li>
                ))}
              </ul>
            ) : null}
            <p className="mt-5 text-sm font-semibold">Planned output name</p>
            <ul
              className="mt-2 rounded-lg border bg-fd-background/60 px-4 py-3 font-mono text-xs leading-6"
              data-testid="planned-outputs"
            >
              {current.plan.outputs.map((output) => (
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
          disabled={!template || !workbook || hasUnusableNumber}
          onCancel={runState.cancel}
          onRun={start}
          readingLabel="Reading the template and the records…"
          runLabel="Run populate"
          state={runState}
        />
      </section>

      {/*
        No archive name: a populate produces exactly one deck, and a zip
        holding one file is friction rather than a convenience.
      */}
      <ResultsPanel
        fallbackMediaType={PRESENTATION_MEDIA_TYPE}
        state={runState}
      />
    </ToolShell>
  );
}

export function PptxInspectTool() {
  const reportHeadingId = useId();

  const templateSelection = useFileSelection(
    isPresentationFile,
    "a PowerPoint .pptx presentation",
  );
  const template = templateSelection.file;
  const [templateSlide, setTemplateSlide] = useState("");
  const [inspected, setInspected] = useState<InspectedTemplate | null>(null);

  const templateSlideField = positiveIntegerField(
    templateSlide,
    "Template slide",
  );
  const slideNumber = suppliedNumber(templateSlideField);
  const slideNumberError = fieldMessage(templateSlideField);

  // Inspection only reads; it writes nothing and produces no file, so there is
  // no run step to confirm. The report simply follows the chosen template and
  // slide, debounced so typing a slide number does not re-read the package on
  // every keystroke. A slide number that cannot be used stops the inspection
  // instead of quietly reporting on slide 1.
  // Raw text again, for the reason the populate page keys on it: "" and "0"
  // both parse to "no slide supplied" and must not share a key.
  const inspectionKey = `${template?.id ?? ""}:${templateSlide.trim()}`;
  const current = inspected?.key === inspectionKey ? inspected : null;
  const readyToInspect = template !== null && slideNumberError === undefined;

  useEffect(() => {
    if (!template || slideNumberError !== undefined) {
      return;
    }
    let active = true;
    const timer = window.setTimeout(() => {
      void (async () => {
        try {
          const outcome = await runOperation({
            kind: "pptx.inspect",
            template: { bytes: template.bytes, name: template.name },
            templateSlide: slideNumber,
          });
          if (active) {
            setInspected({ error: null, key: inspectionKey, outcome });
          }
        } catch (error) {
          if (active) {
            setInspected({
              error: describeFailure(error),
              key: inspectionKey,
              outcome: null,
            });
          }
        }
      })();
    }, PREVIEW_DEBOUNCE_MS);

    return () => {
      active = false;
      window.clearTimeout(timer);
    };
  }, [inspectionKey, slideNumber, slideNumberError, template]);

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
              templateSelection.choose(files, () => {
                setInspected(null);
              });
            }}
          />
        </div>
        {templateSelection.reading ? (
          <ReadingFile testId="source-reading" />
        ) : null}
        {template ? (
          <ChosenFile file={template} testId="source-summary" />
        ) : null}
        {templateSelection.rejected ? (
          <p
            className={noticeClass}
            data-testid="template-rejected"
            role="alert"
          >
            {templateSelection.rejected}
          </p>
        ) : null}

        <div className="mt-6">
          <NumberField
            disabled={false}
            error={slideNumberError}
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
        {readyToInspect && current === null ? (
          <p
            className="mt-3 text-sm text-fd-muted-foreground"
            data-testid="inspection-pending"
          >
            Reading the template…
          </p>
        ) : null}
        {current?.error ? (
          <pre className={noticeClass} data-testid="inspection-error">
            {current.error}
          </pre>
        ) : null}
        {slideNumberError ? (
          <p className={noticeClass} data-testid="inspection-invalid-slide">
            {slideNumberError}
          </p>
        ) : null}
        {current?.outcome ? (
          <>
            <dl className="mt-4 grid grid-cols-2 gap-4 sm:grid-cols-4">
              <div>
                <dt className={metricLabelClass}>Slide</dt>
                <dd
                  className="text-2xl font-bold"
                  data-testid="inspection-slide"
                >
                  {current.outcome.inspection.slideNumber}
                </dd>
              </div>
              <div>
                <dt className={metricLabelClass}>Placeholders</dt>
                <dd
                  className="text-2xl font-bold"
                  data-testid="inspection-fields"
                >
                  {current.outcome.result.metrics.placeholderFields}
                </dd>
              </div>
              <div>
                <dt className={metricLabelClass}>Occurrences</dt>
                <dd
                  className="text-2xl font-bold"
                  data-testid="inspection-occurrences"
                >
                  {current.outcome.result.metrics.placeholderOccurrences}
                </dd>
              </div>
              <div>
                <dt className={metricLabelClass}>Malformed</dt>
                <dd
                  className="text-2xl font-bold"
                  data-testid="inspection-malformed"
                >
                  {current.outcome.result.metrics.malformedPlaceholderLocations}
                </dd>
              </div>
            </dl>

            {current.outcome.inspection.placeholders.length > 0 ? (
              <ul
                className="mt-5 flex flex-col gap-2"
                data-testid="placeholder-list"
              >
                {current.outcome.inspection.placeholders.map((placeholder) => (
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
            ) : null}

            {/*
              The warnings come from the operation's own structured result, so
              the page never invents a sentence about a template: every
              condition that would make a populate refuse this slide is
              described once, in the library, and rendered here verbatim.
            */}
            {current.outcome.result.warnings.length > 0 ? (
              <ul className={noticeClass} data-testid="inspection-warnings">
                {current.outcome.result.warnings.map((warning) => (
                  <li data-testid="inspection-warning" key={warning}>
                    {warning}
                  </li>
                ))}
              </ul>
            ) : null}
          </>
        ) : null}
      </section>
    </ToolShell>
  );
}
