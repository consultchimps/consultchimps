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
 * The split runs the byte API's default all-worksheet mode, matching the
 * command line: with no source named it filters every worksheet that carries
 * the chosen column and carries the rest of the workbook into each output
 * untouched. Naming a worksheet, Excel Table, or named range — or clearing
 * whole-workbook preservation — selects the narrower single-source modes that
 * write compact, data-only workbooks, so the controls below have to say which
 * mode the current selection is really in.
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
import { BROWSER_TOOLS } from "@/lib/tools";
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

/**
 * The route of another online tool, taken from the registry so a cross-link
 * can never point at a browser surface that does not exist. An operation whose
 * browser surface is off has no route, and the pointer to it is left out.
 */
function browserToolHref(slug: string): string | undefined {
  return BROWSER_TOOLS.find((tool) => tool.slug === slug)?.surfaces.browser
    .href;
}

const inlineLinkClass = "font-semibold text-fd-primary hover:underline";
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

interface OrderedUploads {
  readonly add: (added: readonly UploadedFile[]) => void;
  readonly files: readonly UploadedFile[];
  readonly move: (index: number, offset: number) => void;
  readonly remove: (id: string) => void;
}

/**
 * The hand-arranged list of workbooks both multi-input tools collect. Order is
 * part of the request in each of them — it decides tab order for the merge and
 * row order for the consolidate — so entries stay where they were put and are
 * only ever moved deliberately.
 */
function useOrderedUploads(): OrderedUploads {
  const [files, setFiles] = useState<readonly UploadedFile[]>([]);

  const add = useCallback((added: readonly UploadedFile[]) => {
    setFiles((previous) => [...previous, ...added]);
  }, []);

  const move = useCallback((index: number, offset: number) => {
    setFiles((previous) => {
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

  const remove = useCallback((id: string) => {
    setFiles((previous) => previous.filter((entry) => entry.id !== id));
  }, []);

  return { add, files, move, remove };
}

interface SourceWorkbookListProps {
  readonly disabled: boolean;
  readonly uploads: OrderedUploads;
}

/** The numbered source list, with the controls that reorder and prune it. */
function SourceWorkbookList({ disabled, uploads }: SourceWorkbookListProps) {
  const { files, move, remove } = uploads;

  if (files.length === 0) {
    return (
      <p className="mt-4 text-sm text-fd-muted-foreground">
        No workbooks added yet.
      </p>
    );
  }

  return (
    <ol className="mt-4 flex flex-col gap-2" data-testid="source-list">
      {files.map((file, index) => (
        <li
          className="flex flex-wrap items-center justify-between gap-3 rounded-lg border bg-fd-background/60 px-3 py-2"
          data-testid="source-item"
          key={file.id}
        >
          <span className="flex min-w-0 items-center gap-2">
            <span className="shrink-0 font-mono text-xs text-fd-muted-foreground">
              {String(index + 1).padStart(2, "0")}
            </span>
            <span className="truncate font-mono text-sm">{file.name}</span>
            <span className="shrink-0 text-xs text-fd-muted-foreground">
              {formatBytes(file.bytes.byteLength)}
            </span>
          </span>
          <span className="flex shrink-0 items-center gap-2">
            <button
              aria-label={`Move ${file.name} earlier`}
              className={compactButtonClass}
              disabled={index === 0 || disabled}
              onClick={() => move(index, -1)}
              type="button"
            >
              <ArrowUp className="size-3.5" aria-hidden="true" />
            </button>
            <button
              aria-label={`Move ${file.name} later`}
              className={compactButtonClass}
              disabled={index === files.length - 1 || disabled}
              onClick={() => move(index, 1)}
              type="button"
            >
              <ArrowDown className="size-3.5" aria-hidden="true" />
            </button>
            <button
              aria-label={`Remove ${file.name}`}
              className={compactButtonClass}
              disabled={disabled}
              onClick={() => remove(file.id)}
              type="button"
            >
              <Trash2 className="size-3.5" aria-hidden="true" />
            </button>
          </span>
        </li>
      ))}
    </ol>
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
  // Whole-workbook preservation is the API default wherever it is offered, so
  // the control starts on and only ever sends the opt-out.
  const [preserveWorkbook, setPreserveWorkbook] = useState(true);
  const [strict, setStrict] = useState(false);
  const [values, setValues] = useState(false);
  const [plan, setPlan] =
    useState<OperationPlan<SplitWorkbookByColumnPlanMetric> | null>(null);
  const [planError, setPlanError] = useState<string | null>(null);

  const runState = useOperationRun();
  const isRunning = runState.status === "running";
  const tableName = table.trim();
  const rangeName = range.trim();
  const sheetName = sheet.trim();
  // An Excel Table and a named range both carry their own headers, so the API
  // refuses a header row alongside either one.
  const headerRowAllowed = !tableName && !rangeName;
  // Whole-workbook preservation exists for the all-worksheet default and for
  // an Excel Table split. A named range or a named worksheet is always rebuilt
  // compactly, so the option would be a promise the API cannot keep.
  const preserveWorkbookAllowed =
    Boolean(tableName) || (!rangeName && !sheetName);
  // The API's dispatch rule, mirrored so the controls can tell the truth: no
  // source named and preservation left on means every worksheet is filtered
  // in place, which is also the only mode that ignores the blank and hidden
  // worksheet options.
  const allWorksheetMode =
    !tableName && !rangeName && !sheetName && preserveWorkbook;

  const options = useMemo<WorkbookSplitOptions>(() => {
    const parsedHeaderRow = Number.parseInt(headerRow, 10);
    return {
      column: column.trim(),
      filenamePrefix: prefix.trim() || undefined,
      headerRow:
        headerRowAllowed && Number.isFinite(parsedHeaderRow)
          ? parsedHeaderRow
          : undefined,
      includeBlank: allWorksheetMode ? undefined : includeBlank,
      includeHiddenSheets: allWorksheetMode ? undefined : includeHiddenSheets,
      // Only the opt-out travels; leaving it unset lets each mode apply its own
      // default, and no mode is asked for a preservation it does not offer.
      preserveWorkbook:
        preserveWorkbookAllowed && !preserveWorkbook ? false : undefined,
      range: rangeName || undefined,
      sheet: sheetName || undefined,
      strict,
      table: tableName || undefined,
      values,
    };
  }, [
    allWorksheetMode,
    column,
    headerRow,
    headerRowAllowed,
    includeBlank,
    includeHiddenSheets,
    prefix,
    preserveWorkbook,
    preserveWorkbookAllowed,
    rangeName,
    sheetName,
    strict,
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
      description="Choose a workbook and a column, and get one workbook per distinct value in that column. By default each new workbook keeps the source workbook's sheets, formatting, and supported workbook structure, removing only the rows that belong to other values. Pivot tables and their caches are removed and reported, so review complex workbooks — pivots, external links, charts, ActiveX controls — in Excel before you deliver them. Everything runs in this page using the same operation the ConsultChimps library uses."
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
              hint="Matching ignores surrounding whitespace and letter case, unless strict matching is turned on below."
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
          <p className={`${fieldHintClass} mt-3`}>
            Leave the worksheet, table, and range fields empty to split every
            worksheet that contains your column and keep the rest of the
            workbook in each new file.
          </p>
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
              hint="Optional. Splits one worksheet by name, and the new files then hold only that worksheet's matching rows."
              label="Worksheet"
              onChange={setSheet}
              placeholder="Clients"
              testId="sheet-input"
              value={sheet}
            />
            <TextField
              disabled={isRunning}
              hint="Optional. Splits a named Excel table, which gives the safest data boundaries. The rest of the workbook is kept unless you turn that off below."
              label="Excel table"
              onChange={setTable}
              placeholder="ClientData"
              testId="table-input"
              value={table}
            />
            <TextField
              disabled={isRunning}
              hint="Optional. Splits a workbook-level named range instead, and the new files then hold only that range's matching rows. Cannot be combined with an Excel table."
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
              checked={preserveWorkbookAllowed && preserveWorkbook}
              disabled={isRunning || !preserveWorkbookAllowed}
              hint={
                preserveWorkbookAllowed
                  ? "On by default. The source workbook's sheets, formatting, and supported workbook structure are kept, and only the rows that do not belong are removed. Pivot tables and their caches are removed, with a warning on the result. Turn it off to get small, plain workbooks holding just the matching rows of the source being split."
                  : "Not offered for a named worksheet or a named range: those always produce small, plain workbooks holding just the matching rows."
              }
              label="Keep the whole workbook"
              onChange={setPreserveWorkbook}
              testId="preserve-workbook-checkbox"
            />
            <CheckboxField
              checked={!allWorksheetMode && includeBlank}
              disabled={isRunning || allWorksheetMode}
              hint={
                allWorksheetMode
                  ? "Only applies when you name a worksheet, table, or range, or turn off keeping the whole workbook. Otherwise rows with a blank value never get a workbook of their own."
                  : "Write a workbook for rows whose split value is blank."
              }
              label="Include blank values"
              onChange={setIncludeBlank}
              testId="include-blank-checkbox"
            />
            <CheckboxField
              checked={!allWorksheetMode && includeHiddenSheets}
              disabled={isRunning || allWorksheetMode}
              hint={
                allWorksheetMode
                  ? "Only applies when you name a worksheet, table, or range, or turn off keeping the whole workbook. Otherwise hidden worksheets are always split too, and stay hidden in every new file."
                  : "Search hidden and very hidden worksheets as well."
              }
              label="Include hidden worksheets"
              onChange={setIncludeHiddenSheets}
              testId="include-hidden-checkbox"
            />
            <CheckboxField
              checked={strict}
              disabled={isRunning}
              hint="Treat differences in letter case, surrounding whitespace, and value type as different values."
              label="Strict matching"
              onChange={setStrict}
              testId="strict-checkbox"
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
            <dl className="mt-4 grid grid-cols-2 gap-4 sm:grid-cols-3">
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
              {/* The two counts that make the split's reach visible: how many
                  tabs get filtered, and how many ride along unchanged. */}
              <div>
                <dt className="font-mono text-xs uppercase tracking-[0.12em] text-fd-muted-foreground">
                  Tabs filtered
                </dt>
                <dd className="text-2xl font-bold">
                  {plan.metrics.sheetsFiltered}
                </dd>
              </div>
              <div>
                <dt className="font-mono text-xs uppercase tracking-[0.12em] text-fd-muted-foreground">
                  Tabs kept as is
                </dt>
                <dd className="text-2xl font-bold">
                  {plan.metrics.sheetsCopiedUnchanged}
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
        <p className="mt-3 text-sm text-fd-muted-foreground">
          {preserveWorkbookAllowed && preserveWorkbook
            ? "Each new workbook keeps your file's sheets, formatting, and supported workbook structure, holding only the rows for one value. Pivot tables and their caches are removed and reported; review complex workbooks in Excel before delivery."
            : "Each new workbook is a small, plain file holding only the matching rows from the source you chose."}
        </p>
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
  const uploads = useOrderedUploads();
  const { add, files } = uploads;
  const [outputName, setOutputName] = useState("");
  const [values, setValues] = useState(false);
  const runState = useOperationRun();
  const isRunning = runState.status === "running";
  const consolidateHref = browserToolHref("spreadsheet-consolidate");

  const start = useCallback(() => {
    if (files.length === 0) {
      return;
    }
    void runState.run({
      kind: "xlsx.merge",
      inputs: files.map((file) => ({ bytes: file.bytes, name: file.name })),
      outputName: outputName.trim() || undefined,
      values,
    });
  }, [files, outputName, runState, values]);

  return (
    <ToolShell
      description="Add the workbooks you want to combine, arrange them in the order their tabs should appear, and merge them without uploading anything. Every source worksheet stays its own separate tab — no rows are stacked into one sheet."
      guideHref="/docs/tools/workbook-merge"
      guideLabel="Read the merge guide"
      kicker="Online tool · Excel merge"
      title="Merge Excel workbooks"
    >
      {consolidateHref ? (
        <p className="text-sm text-fd-muted-foreground">
          Looking for one combined table instead of separate tabs?{" "}
          <Link className={inlineLinkClass} href={consolidateHref}>
            Consolidate the workbooks
          </Link>{" "}
          to stack every worksheet&rsquo;s rows into a single sheet —{" "}
          <Link
            className={inlineLinkClass}
            href="/docs/tools/workbook-merge#which-one-do-i-want"
          >
            which one do I want?
          </Link>
        </p>
      ) : null}
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
            onFiles={(chosen) => {
              void readUploads(chosen, isWorkbookFile).then((read) => {
                if (read.length > 0) {
                  add(read);
                  runState.reset();
                }
              });
            }}
          />
        </div>

        <SourceWorkbookList disabled={isRunning} uploads={uploads} />

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
          {files.length === 0
            ? "Add at least one workbook to merge."
            : `Every worksheet from ${files.length} ${
                files.length === 1 ? "workbook" : "workbooks"
              } will become its own tab, in the order listed above, alongside a Sheet Index tab.`}
        </p>
        <RunControls
          busyLabel="Merging…"
          disabled={files.length === 0}
          onCancel={runState.cancel}
          onRun={start}
          readingLabel="Reading the workbooks…"
          runLabel="Run merge"
          state={runState}
        />
      </section>

      <ResultsPanel fallbackMediaType={WORKBOOK_MEDIA_TYPE} state={runState} />
    </ToolShell>
  );
}

export function ExcelConsolidateTool() {
  const uploads = useOrderedUploads();
  const { add, files } = uploads;
  const [outputName, setOutputName] = useState("");
  const [normalizeHeaders, setNormalizeHeaders] = useState(false);
  // The core adds the source columns unless told otherwise, so the control
  // starts on and the page never has to restate that default elsewhere.
  const [addSourceColumns, setAddSourceColumns] = useState(true);
  const [includeHiddenSheets, setIncludeHiddenSheets] = useState(false);
  const runState = useOperationRun();
  const isRunning = runState.status === "running";
  const mergeHref = browserToolHref("workbook-merge");

  const start = useCallback(() => {
    if (files.length === 0) {
      return;
    }
    void runState.run({
      kind: "xlsx.consolidate",
      inputs: files.map((file) => ({ bytes: file.bytes, name: file.name })),
      addSourceColumns,
      includeHiddenSheets,
      normalizeHeaders,
      outputName: outputName.trim() || undefined,
    });
  }, [
    addSourceColumns,
    files,
    includeHiddenSheets,
    normalizeHeaders,
    outputName,
    runState,
  ]);

  return (
    <ToolShell
      description="Add the workbooks you want to stack, arrange them in the order the rows should follow, and get one table holding every row from every visible worksheet that holds data. Hidden worksheets are skipped unless you ask for them. Nothing is uploaded — the whole task runs in this browser tab."
      guideHref="/docs/tools/spreadsheet-consolidate"
      guideLabel="Read the consolidate guide"
      kicker="Online tool · Excel consolidate"
      title="Consolidate Excel workbooks"
    >
      {mergeHref ? (
        <p className="text-sm text-fd-muted-foreground">
          Want each worksheet kept as its own tab instead of one stacked table?{" "}
          <Link className={inlineLinkClass} href={mergeHref}>
            Merge the workbooks
          </Link>{" "}
          —{" "}
          <Link
            className={inlineLinkClass}
            href="/docs/tools/workbook-merge#which-one-do-i-want"
          >
            which one do I want?
          </Link>
        </p>
      ) : null}
      <section className={sectionClass} data-testid="source-section">
        <h2 className="text-xl font-bold tracking-[-0.03em]">
          1. Add workbooks
        </h2>
        <div className="mt-4">
          <FilePicker
            accept={WORKBOOK_ACCEPT}
            description="Drag one or more .xlsx workbooks here, or pick them with the button below. Rows are stacked in the order shown."
            disabled={isRunning}
            label="Source workbooks"
            multiple
            onFiles={(chosen) => {
              void readUploads(chosen, isWorkbookFile).then((read) => {
                if (read.length > 0) {
                  add(read);
                  runState.reset();
                }
              });
            }}
          />
        </div>

        <SourceWorkbookList disabled={isRunning} uploads={uploads} />

        <div className="mt-6 flex flex-col gap-5">
          <TextField
            disabled={isRunning}
            hint="Optional. Defaults to `consolidated.xlsx`. The `.xlsx` extension is added for you."
            label="Output filename"
            onChange={setOutputName}
            placeholder="all-rows"
            testId="output-name-input"
            value={outputName}
          />
          <CheckboxField
            checked={normalizeHeaders}
            disabled={isRunning}
            hint="Treat columns whose headers differ only in case, spacing, or punctuation — “Failed Checks” and “Failed_Checks”, say — as one column. The first spelling seen names the output column."
            label="Normalize headers"
            onChange={setNormalizeHeaders}
            testId="normalize-headers-checkbox"
          />
          <CheckboxField
            checked={addSourceColumns}
            disabled={isRunning}
            hint="Record where each row came from in added `_source_file`, `_source_sheet`, and `_source_row` columns."
            label="Add source columns"
            onChange={setAddSourceColumns}
            testId="source-columns-checkbox"
          />
          <CheckboxField
            checked={includeHiddenSheets}
            disabled={isRunning}
            hint="Read hidden and very hidden worksheets as well."
            label="Include hidden worksheets"
            onChange={setIncludeHiddenSheets}
            testId="include-hidden-checkbox"
          />
        </div>
      </section>

      <section className={sectionClass} data-testid="run-section">
        <h2 className="text-xl font-bold tracking-[-0.03em]">2. Run</h2>
        <p className="mt-3 text-sm text-fd-muted-foreground">
          {files.length === 0
            ? "Add at least one workbook to consolidate."
            : `Rows from every ${
                includeHiddenSheets ? "" : "visible "
              }worksheet that holds data in ${files.length} ${
                files.length === 1 ? "workbook" : "workbooks"
              } will be stacked into one table, in the order listed above.`}
        </p>
        <RunControls
          busyLabel="Consolidating…"
          disabled={files.length === 0}
          onCancel={runState.cancel}
          onRun={start}
          readingLabel="Reading the workbooks…"
          runLabel="Run consolidate"
          state={runState}
        />
      </section>

      <ResultsPanel fallbackMediaType={WORKBOOK_MEDIA_TYPE} state={runState} />
    </ToolShell>
  );
}
