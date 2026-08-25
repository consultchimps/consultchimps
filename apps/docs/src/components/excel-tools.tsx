"use client";

/**
 * Browser front ends for the byte-level workbook operations in
 * `@consultchimps/xlsx/bytes`.
 *
 * As with the PDF pages, the components hold state and render; the workbook
 * parsing, grouping, and package rewriting all happen in the shared operation
 * worker. Rewriting an OOXML package is the heaviest thing either tool does,
 * so keeping it off the main thread is what makes a large workbook usable.
 *
 * The split here is the byte API's single-source split: it reads one
 * worksheet, Excel Table, or named range and writes compact workbooks. The
 * command line's default is the newer all-worksheet split, which is a
 * deliberate divergence documented in the split guide.
 */

import {
  compactButtonClass,
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
import type {
  WorkbookSplitOptions,
  WorksheetColumns,
} from "@/lib/operation-tasks";
import { runOperation } from "@/lib/operation-worker";
import type { OperationPlan } from "@consultchimps/core";
// Type-only: the runtime module is loaded inside the worker.
import type { SplitWorkbookByColumnPlanMetric } from "@consultchimps/xlsx/bytes";
import { ArrowDown, ArrowUp, FileText, Trash2 } from "lucide-react";
import Link from "next/link";
import { useCallback, useEffect, useId, useMemo, useState } from "react";

const WORKBOOK_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const WORKBOOK_ACCEPT = `${WORKBOOK_MEDIA_TYPE},.xlsx`;
/** The dropdown entry that hands column naming back to the text field. */
const MANUAL_COLUMN = "\u0000manual";

function isWorkbookFile(file: File): boolean {
  return file.type === WORKBOOK_MEDIA_TYPE || /\.xlsx$/iu.test(file.name);
}

function withoutWorkbookExtension(name: string): string {
  return name.replace(/\.xlsx$/iu, "");
}

const checkboxLabelClass =
  "flex items-start gap-2.5 text-sm font-medium leading-6";
const checkboxClass =
  "mt-1 size-4 shrink-0 rounded border-fd-border accent-fd-primary";
const fieldLabelClass = "block text-sm font-semibold";
const fieldHintClass = "mt-1 text-sm text-fd-muted-foreground";

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

interface CheckboxFieldProps {
  readonly checked: boolean;
  readonly disabled: boolean;
  readonly hint: string;
  readonly label: string;
  readonly onChange: (checked: boolean) => void;
  readonly testId: string;
}

function CheckboxField({
  checked,
  disabled,
  hint,
  label,
  onChange,
  testId,
}: CheckboxFieldProps) {
  return (
    <label className={checkboxLabelClass}>
      <input
        checked={checked}
        className={checkboxClass}
        data-testid={testId}
        disabled={disabled}
        onChange={(event) => onChange(event.target.checked)}
        type="checkbox"
      />
      <span>
        {label}
        <span className="block text-sm font-normal text-fd-muted-foreground">
          {hint}
        </span>
      </span>
    </label>
  );
}

export function ExcelSplitTool() {
  const columnSelectId = useId();
  const headerRowId = useId();
  const previewHeadingId = useId();

  const [input, setInput] = useState<UploadedFile | null>(null);
  const [detected, setDetected] = useState<WorksheetColumns | null>(null);
  const [column, setColumn] = useState("");
  const [isManualColumn, setIsManualColumn] = useState(false);
  const [prefix, setPrefix] = useState("");
  const [sheet, setSheet] = useState("");
  const [table, setTable] = useState("");
  const [range, setRange] = useState("");
  const [headerRow, setHeaderRow] = useState("");
  const [includeBlank, setIncludeBlank] = useState(false);
  const [includeHiddenSheets, setIncludeHiddenSheets] = useState(false);
  // The byte API preserves the whole workbook by default whenever a table is
  // named, so the control starts on and only applies while a table is named.
  const [preserveWorkbook, setPreserveWorkbook] = useState(true);
  const [values, setValues] = useState(false);
  const [plan, setPlan] =
    useState<OperationPlan<SplitWorkbookByColumnPlanMetric> | null>(null);
  const [planError, setPlanError] = useState<string | null>(null);

  const runState = useOperationRun();
  const isRunning = runState.status === "running";
  const tableName = table.trim();
  const rangeName = range.trim();
  // An Excel Table and a named range both carry their own headers, so the API
  // refuses a header row alongside either one.
  const headerRowAllowed = !tableName && !rangeName;

  const options = useMemo<WorkbookSplitOptions>(() => {
    const parsedHeaderRow = Number.parseInt(headerRow, 10);
    return {
      column: column.trim(),
      filenamePrefix: prefix.trim() || undefined,
      headerRow:
        headerRowAllowed && Number.isFinite(parsedHeaderRow)
          ? parsedHeaderRow
          : undefined,
      includeBlank,
      includeHiddenSheets,
      preserveWorkbook: tableName ? preserveWorkbook : undefined,
      range: rangeName || undefined,
      sheet: sheet.trim() || undefined,
      table: tableName || undefined,
      values,
    };
  }, [
    column,
    headerRow,
    headerRowAllowed,
    includeBlank,
    includeHiddenSheets,
    prefix,
    preserveWorkbook,
    rangeName,
    sheet,
    tableName,
    values,
  ]);

  const resetSource = useCallback(() => {
    setDetected(null);
    setPlan(null);
    setPlanError(null);
  }, []);

  // Read the chosen worksheet's headers so the column can be picked from a
  // list. A workbook the reader cannot make sense of simply leaves the manual
  // field in charge; the split itself reports the real problem.
  useEffect(() => {
    if (!input) {
      return;
    }

    let active = true;
    void (async () => {
      try {
        const columns = await runOperation({
          kind: "xlsx.columns",
          input: { bytes: input.bytes, name: input.name },
          headerRow: options.headerRow,
          worksheet: options.sheet,
        });
        if (active) {
          setDetected(columns);
        }
      } catch {
        if (active) {
          setDetected(null);
        }
      }
    })();

    return () => {
      active = false;
    };
  }, [input, options.headerRow, options.sheet]);

  // Re-planning re-parses the workbook, so wait for a pause in typing first.
  // Clearing a stale preview waits for the same pause, which keeps every
  // state change in this effect asynchronous.
  useEffect(() => {
    let active = true;
    const timer = window.setTimeout(() => {
      if (!input || !options.column) {
        setPlan(null);
        setPlanError(null);
        return;
      }
      void (async () => {
        try {
          const nextPlan = await runOperation({
            kind: "xlsx.plan-split",
            input: { bytes: input.bytes, name: input.name },
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
  }, [input, options]);

  const start = useCallback(() => {
    if (!input || !options.column) {
      return;
    }
    void runState.run({
      kind: "xlsx.split",
      input: { bytes: input.bytes, name: input.name },
      options,
    });
  }, [input, options, runState]);

  const detectedColumns = detected?.columns ?? [];
  const showManualField = isManualColumn || detectedColumns.length === 0;
  const archiveName = `${(prefix.trim() || (input ? withoutWorkbookExtension(input.name) : "workbook")).replace(/[<>:"/\\|?*]+/gu, "-")}-split.zip`;

  return (
    <ToolShell
      description="Choose a workbook and a column, and get one workbook per distinct value in that column. Everything runs in this page using the same operation the ConsultChimps library uses."
      guideHref="/docs/tools/spreadsheet-split"
      guideLabel="Read the split guide"
      kicker="Online tool · Excel split"
      title="Split an Excel workbook"
    >
      <section className={sectionClass} data-testid="source-section">
        <h2 className="text-xl font-bold tracking-[-0.03em]">
          1. Choose a workbook
        </h2>
        <div className="mt-4">
          <FilePicker
            accept={WORKBOOK_ACCEPT}
            description="Drag an .xlsx workbook here, or pick one with the button below. Only the first workbook is used."
            disabled={isRunning}
            label="Source workbook"
            multiple={false}
            onFiles={(files) => {
              void readUploads(files, isWorkbookFile).then((read) => {
                const [first] = read;
                if (first) {
                  setInput(first);
                  resetSource();
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
      </section>

      <section className={sectionClass} data-testid="column-section">
        <h2 className="text-xl font-bold tracking-[-0.03em]">
          2. Choose the split column
        </h2>
        <p className="mt-3 text-sm text-fd-muted-foreground">
          {detectedColumns.length > 0
            ? `Headers found on worksheet “${detected?.worksheet ?? ""}”. Pick one, or type a column name instead.`
            : "Type the column name exactly as it appears in the header row."}
        </p>

        {detectedColumns.length > 0 ? (
          <div className="mt-4">
            <label className={fieldLabelClass} htmlFor={columnSelectId}>
              Detected column
            </label>
            <select
              className={`${inputClass} mt-2`}
              data-testid="column-select"
              disabled={isRunning}
              id={columnSelectId}
              onChange={(event) => {
                const chosen = event.target.value;
                if (chosen === MANUAL_COLUMN) {
                  setIsManualColumn(true);
                  return;
                }
                setIsManualColumn(false);
                setColumn(chosen);
              }}
              value={isManualColumn ? MANUAL_COLUMN : column}
            >
              <option value="">Choose a column…</option>
              {detectedColumns.map((name) => (
                <option key={name} value={name}>
                  {name}
                </option>
              ))}
              <option value={MANUAL_COLUMN}>Other (type a name)…</option>
            </select>
          </div>
        ) : null}

        {showManualField ? (
          <div className="mt-4">
            <TextField
              disabled={isRunning}
              hint="Matching ignores surrounding whitespace and letter case."
              label="Column name"
              onChange={setColumn}
              placeholder="Region"
              testId="column-input"
              value={column}
            />
          </div>
        ) : null}

        <details className="mt-6 rounded-lg border bg-fd-background/50 px-4 py-3">
          <summary className="cursor-pointer text-sm font-semibold">
            Advanced options
          </summary>
          <div className="mt-5 flex flex-col gap-5">
            <TextField
              disabled={isRunning}
              hint="Optional. Defaults to the source filename, so `clients.xlsx` produces `clients-North.xlsx`."
              label="Filename prefix"
              onChange={setPrefix}
              placeholder="clients"
              testId="prefix-input"
              value={prefix}
            />
            <TextField
              disabled={isRunning}
              hint="Optional. Limits the search to one worksheet by name."
              label="Worksheet"
              onChange={setSheet}
              placeholder="Clients"
              testId="sheet-input"
              value={sheet}
            />
            <TextField
              disabled={isRunning}
              hint="Optional. Reads a named Excel Table, which gives the safest data boundaries."
              label="Excel table"
              onChange={setTable}
              placeholder="ClientData"
              testId="table-input"
              value={table}
            />
            <TextField
              disabled={isRunning}
              hint="Optional. Reads a workbook-level named range instead. Cannot be combined with an Excel table."
              label="Named range"
              onChange={setRange}
              placeholder="ClientRange"
              testId="range-input"
              value={range}
            />
            <div>
              <label className={fieldLabelClass} htmlFor={headerRowId}>
                Header row
              </label>
              <p className={fieldHintClass}>
                Optional one-based row number. Not available when an Excel table
                or a named range provides the headers.
              </p>
              <input
                className={`${inputClass} mt-2`}
                data-testid="header-row-input"
                disabled={isRunning || !headerRowAllowed}
                id={headerRowId}
                min={1}
                onChange={(event) => setHeaderRow(event.target.value)}
                placeholder="1"
                type="number"
                value={headerRowAllowed ? headerRow : ""}
              />
            </div>
            <CheckboxField
              checked={includeBlank}
              disabled={isRunning}
              hint="Write a workbook for rows whose split value is blank."
              label="Include blank values"
              onChange={setIncludeBlank}
              testId="include-blank-checkbox"
            />
            <CheckboxField
              checked={includeHiddenSheets}
              disabled={isRunning}
              hint="Search hidden and very hidden worksheets as well."
              label="Include hidden worksheets"
              onChange={setIncludeHiddenSheets}
              testId="include-hidden-checkbox"
            />
            <CheckboxField
              checked={Boolean(tableName) && preserveWorkbook}
              disabled={isRunning || !tableName}
              hint="Keep the complete source workbook and replace only the table's rows. Requires an Excel table, and is on by default when one is named."
              label="Preserve the whole workbook"
              onChange={setPreserveWorkbook}
              testId="preserve-workbook-checkbox"
            />
            <CheckboxField
              checked={values}
              disabled={isRunning}
              hint="Replace every formula with its most recently saved result."
              label="Values only"
              onChange={setValues}
              testId="values-checkbox"
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
        {!input || !options.column ? (
          <p className="mt-3 text-sm text-fd-muted-foreground">
            Choose a workbook and a column to see the workbooks this task will
            create.
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
                <dt className="font-mono text-xs uppercase tracking-[0.12em] text-fd-muted-foreground">
                  Groups
                </dt>
                <dd className="text-2xl font-bold">{plan.metrics.groups}</dd>
              </div>
              <div>
                <dt className="font-mono text-xs uppercase tracking-[0.12em] text-fd-muted-foreground">
                  Rows read
                </dt>
                <dd className="text-2xl font-bold">{plan.metrics.inputRows}</dd>
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
                  Rows skipped
                </dt>
                <dd className="text-2xl font-bold">
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
        <h2 className="text-xl font-bold tracking-[-0.03em]">4. Run</h2>
        <RunControls
          busyLabel="Splitting…"
          disabled={!input || !options.column}
          onCancel={runState.cancel}
          onRun={start}
          readingLabel="Reading the workbook…"
          runLabel="Run split"
          state={runState}
        />
      </section>

      <ResultsPanel
        archiveName={archiveName}
        fallbackMediaType={WORKBOOK_MEDIA_TYPE}
        state={runState}
      />
    </ToolShell>
  );
}

export function ExcelMergeTool() {
  const [inputs, setInputs] = useState<readonly UploadedFile[]>([]);
  const [outputName, setOutputName] = useState("");
  const [values, setValues] = useState(false);
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
      kind: "xlsx.merge",
      inputs: inputs.map((file) => ({ bytes: file.bytes, name: file.name })),
      outputName: outputName.trim() || undefined,
      values,
    });
  }, [inputs, outputName, runState, values]);

  return (
    <ToolShell
      description="Add the workbooks you want to combine, arrange them in the order their tabs should appear, and merge them without uploading anything. Every source worksheet stays its own separate tab — no rows are stacked into one sheet."
      guideHref="/docs/tools/spreadsheets"
      guideLabel="Read the spreadsheet guide"
      kicker="Online tool · Excel merge"
      title="Merge Excel workbooks"
    >
      <p className="text-sm text-fd-muted-foreground">
        Looking for one combined table instead of separate tabs? That is the
        `sheets consolidate` command, which runs from the command line —{" "}
        <Link
          className="font-semibold text-fd-primary hover:underline"
          href="/docs/tools/spreadsheets#which-one-do-i-want"
        >
          which one do I want?
        </Link>
      </p>
      <section className={sectionClass} data-testid="source-section">
        <h2 className="text-xl font-bold tracking-[-0.03em]">
          1. Add workbooks
        </h2>
        <div className="mt-4">
          <FilePicker
            accept={WORKBOOK_ACCEPT}
            description="Drag one or more .xlsx workbooks here, or pick them with the button below. Added files keep the order shown."
            disabled={isRunning}
            label="Source workbooks"
            multiple
            onFiles={(files) => {
              void readUploads(files, isWorkbookFile).then((read) => {
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
            No workbooks added yet.
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

        <div className="mt-6 flex flex-col gap-5">
          <TextField
            disabled={isRunning}
            hint="Optional. Defaults to `merged.xlsx`. The `.xlsx` extension is added for you."
            label="Output filename"
            onChange={setOutputName}
            placeholder="all-sheets"
            testId="output-name-input"
            value={outputName}
          />
          <CheckboxField
            checked={values}
            disabled={isRunning}
            hint="Replace every formula with its most recently saved result."
            label="Values only"
            onChange={setValues}
            testId="values-checkbox"
          />
        </div>
      </section>

      <section className={sectionClass} data-testid="run-section">
        <h2 className="text-xl font-bold tracking-[-0.03em]">2. Run</h2>
        <p className="mt-3 text-sm text-fd-muted-foreground">
          {inputs.length === 0
            ? "Add at least one workbook to merge."
            : `Every worksheet from ${inputs.length} ${
                inputs.length === 1 ? "workbook" : "workbooks"
              } will become its own tab, in the order listed above, alongside a Sheet Index tab.`}
        </p>
        <RunControls
          busyLabel="Merging…"
          disabled={inputs.length === 0}
          onCancel={runState.cancel}
          onRun={start}
          readingLabel="Reading the workbooks…"
          runLabel="Run merge"
          state={runState}
        />
      </section>

      <ResultsPanel
        archiveName="merged-workbook.zip"
        fallbackMediaType={WORKBOOK_MEDIA_TYPE}
        state={runState}
      />
    </ToolShell>
  );
}
