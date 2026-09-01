"use client";

/**
 * The shared workbook inspection report.
 *
 * Hand it a workbook and it renders what `sheets.inspect` says is inside:
 * worksheets with their visibility, the header row an operation would key on,
 * the Excel Tables and named ranges the package declares, and the bounded
 * sample of each column's values the library collects. The inspection creates
 * nothing, so there is no Run button and no download: choosing a workbook is
 * the whole interaction, and the report follows it.
 *
 * The component owns no page chrome on purpose. It renders one section and
 * takes its workbook as a prop, so the split and consolidate pages can put it
 * beside their own pickers later without inheriting a second heading, a second
 * privacy notice, or a picker they already have. `useWorkbookDescription` is
 * exported beside it for a page that wants the description itself, such as a
 * column-mapping review, rather than this rendering of it.
 *
 * As everywhere else here, the workbook is read in the shared operation
 * worker: parsing a package and scanning every worksheet is the heaviest thing
 * this page does, and the library yields between worksheets so an inspection
 * superseded by another pick is actually cancelled rather than merely ignored.
 */

import {
  compactButtonClass,
  describeFailure,
  noticeClass,
  PREVIEW_DEBOUNCE_MS,
  sectionClass,
  type UploadedFile,
} from "@/components/tool-kit";
import type { WorkbookInspectOptions } from "@/lib/operation-tasks";
import { runOperation } from "@/lib/operation-worker";
import {
  COLUMN_PREVIEW_LIMIT,
  columnPreviewNote,
  hiddenWorksheetCallout,
  sampleValueText,
  visibilityBadge,
  worksheetSummary,
  type SampleValue,
} from "@/lib/workbook-inspection";
// Type-only: the runtime module is loaded inside the worker.
import type {
  WorkbookColumnDescription,
  WorkbookDescriptionOutcome,
} from "@consultchimps/xlsx/bytes";
import { useEffect, useId, useState } from "react";

/** What the hook below reports while, and after, an inspection runs. */
export interface WorkbookInspection {
  /** The plain-language failure, when the workbook could not be described. */
  readonly error: string | null;
  /** The description and the operation's structured result, once they arrive. */
  readonly outcome: WorkbookDescriptionOutcome | null;
  /** True while a chosen workbook has no answer on screen yet. */
  readonly pending: boolean;
}

/** An answer together with the workbook and options it was computed from. */
interface InspectedWorkbook {
  readonly error: string | null;
  readonly key: string;
  readonly outcome: WorkbookDescriptionOutcome | null;
}

/**
 * Describe a workbook in the operation worker, keeping the answer only while
 * it still matches what is on the page.
 *
 * Hidden worksheets are described by default, which is this surface's one
 * deliberate departure from the library default. A page whose entire job is to
 * answer "what is in this file" would otherwise leave out the worksheets a
 * reader is most likely hunting for, and say so only in a warning. Every
 * worksheet the report lists carries its visibility, so nothing is passed off
 * as ordinary, and a caller that wants the library's own default can pass
 * `includeHiddenSheets: false`.
 */
export function useWorkbookDescription(
  file: UploadedFile | null,
  options: WorkbookInspectOptions = {},
): WorkbookInspection {
  // Destructured to primitives so the effect below depends on values rather
  // than on the identity of an options object rebuilt on every render.
  const { headerRow, includeHiddenSheets = true } = options;
  const [inspected, setInspected] = useState<InspectedWorkbook | null>(null);

  // Everything the answer depends on, in one comparable value. A stored
  // description is rendered only while its key still matches the page, so the
  // report can never describe the workbook that was just replaced.
  const key = `${file?.id ?? ""}:${headerRow ?? ""}:${includeHiddenSheets}`;
  const current = inspected?.key === key ? inspected : null;

  useEffect(() => {
    if (!file) {
      return;
    }
    let active = true;
    // Inspecting reopens the package and scans every worksheet, so a
    // superseded attempt is cancelled rather than merely ignored: several
    // abandoned scans queued behind each other would make the one the visitor
    // is actually waiting for wait.
    const controller = new AbortController();
    const timer = window.setTimeout(() => {
      void (async () => {
        try {
          const outcome = await runOperation(
            {
              kind: "xlsx.inspect",
              input: { bytes: file.bytes, name: file.name },
              options: { headerRow, includeHiddenSheets },
            },
            { signal: controller.signal },
          );
          if (active) {
            setInspected({ error: null, key, outcome });
          }
        } catch (error) {
          if (active) {
            setInspected({
              error: describeFailure(error),
              key,
              outcome: null,
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
  }, [file, headerRow, includeHiddenSheets, key]);

  return {
    error: current?.error ?? null,
    outcome: current?.outcome ?? null,
    pending: file !== null && current === null,
  };
}

export interface WorkbookInspectorProps {
  /** Shown in place of the report while no workbook is chosen. */
  readonly emptyMessage?: string;
  /** The workbook to describe, or null while nothing is chosen. */
  readonly file: UploadedFile | null;
  /** Optional one-based header row, as every workbook reader here takes it. */
  readonly headerRow?: number | undefined;
  /** The report's own heading. Left out when the host page supplies one. */
  readonly heading?: string | undefined;
  /** Describe hidden and very hidden worksheets too. Defaults to true. */
  readonly includeHiddenSheets?: boolean | undefined;
}

const metricLabelClass =
  "font-mono text-xs uppercase tracking-[0.12em] text-fd-muted-foreground";
const subheadingClass = "mt-8 text-sm font-semibold";
const emptyListClass = "mt-2 text-sm text-fd-muted-foreground";
const rowClass = "rounded-lg border bg-fd-background/60 px-3 py-2";
const badgeClass =
  "shrink-0 rounded-md border border-fd-primary/40 bg-fd-accent/40 px-2 py-0.5 text-xs font-semibold text-fd-accent-foreground";

/** A sample value's identity, which is its text and its type: 1 is not "1". */
function sampleKey(value: SampleValue): string {
  return `${typeof value}:${String(value)}`;
}

/**
 * One worksheet's header preview, shortened until asked.
 *
 * The library bounds the samples inside a column but not the number of
 * columns, and a worksheet may carry thousands of them. Laying out a row and
 * its sample chips for every one is main-thread work the worker cannot take
 * away, so the list stops at `COLUMN_PREVIEW_LIMIT` and the rest are built
 * only when someone asks for them. An ordinary workbook is well under the
 * limit and never sees the control.
 */
function ColumnList({
  columns,
}: {
  readonly columns: readonly WorkbookColumnDescription[];
}) {
  const [showEveryColumn, setShowEveryColumn] = useState(false);
  const shown = showEveryColumn
    ? columns
    : columns.slice(0, COLUMN_PREVIEW_LIMIT);
  const note = columnPreviewNote(columns.length, shown.length);

  return (
    <>
      <ul className="mt-3 flex flex-col gap-1.5" data-testid="column-list">
        {shown.map((column) => (
          <li
            className="flex flex-wrap items-baseline gap-x-3 gap-y-1 text-sm"
            data-testid="column-item"
            key={column.index}
          >
            <span className="font-mono font-medium" data-testid="column-header">
              {column.header}
            </span>
            {column.sampleValues.length === 0 ? (
              <span className="text-xs text-fd-muted-foreground">
                No values below the header
              </span>
            ) : (
              <span className="flex flex-wrap gap-1.5">
                {column.sampleValues.map((value) => (
                  <span
                    className="rounded border bg-fd-card px-1.5 py-0.5 font-mono text-xs"
                    data-testid="sample-value"
                    key={sampleKey(value)}
                  >
                    {sampleValueText(value)}
                  </span>
                ))}
              </span>
            )}
          </li>
        ))}
      </ul>
      {columns.length > COLUMN_PREVIEW_LIMIT ? (
        <div className="mt-3 flex flex-wrap items-center gap-3">
          {note ? (
            <span
              className="text-xs text-fd-muted-foreground"
              data-testid="column-preview-note"
            >
              {note}
            </span>
          ) : null}
          <button
            className={compactButtonClass}
            data-testid="column-toggle"
            onClick={() => setShowEveryColumn((shownAll) => !shownAll)}
            type="button"
          >
            {showEveryColumn
              ? `Show the first ${COLUMN_PREVIEW_LIMIT} columns`
              : `Show all ${columns.length} columns`}
          </button>
        </div>
      ) : null}
    </>
  );
}

export function WorkbookInspector({
  emptyMessage = "Choose a workbook to see the worksheets, header rows, and structures it holds",
  file,
  headerRow,
  heading,
  includeHiddenSheets,
}: WorkbookInspectorProps) {
  const headingId = useId();
  const { error, outcome, pending } = useWorkbookDescription(file, {
    headerRow,
    includeHiddenSheets,
  });
  const description = outcome?.description;
  const hiddenCallout = description
    ? hiddenWorksheetCallout(description.sheets)
    : undefined;

  return (
    <section
      aria-label={heading === undefined ? "Workbook inspection" : undefined}
      aria-labelledby={heading === undefined ? undefined : headingId}
      aria-live="polite"
      className={sectionClass}
      data-testid="inspection-section"
    >
      {heading === undefined ? null : (
        <h2 className="text-xl font-bold tracking-[-0.03em]" id={headingId}>
          {heading}
        </h2>
      )}

      {file === null ? (
        <p className="mt-3 text-sm text-fd-muted-foreground">{emptyMessage}</p>
      ) : null}

      {pending ? (
        <p
          className="mt-3 text-sm text-fd-muted-foreground"
          data-testid="inspection-pending"
        >
          Reading the workbook…
        </p>
      ) : null}

      {error ? (
        <pre className={noticeClass} data-testid="inspection-error">
          {error}
        </pre>
      ) : null}

      {outcome && description ? (
        <>
          <dl className="mt-4 grid grid-cols-2 gap-4 sm:grid-cols-4">
            <div>
              <dt className={metricLabelClass}>Worksheets</dt>
              <dd
                className="text-2xl font-bold"
                data-testid="inspection-worksheets"
              >
                {outcome.result.metrics.worksheets}
              </dd>
            </div>
            <div>
              <dt className={metricLabelClass}>Data rows</dt>
              <dd
                className="text-2xl font-bold"
                data-testid="inspection-data-rows"
              >
                {outcome.result.metrics.dataRows}
              </dd>
            </div>
            <div>
              <dt className={metricLabelClass}>Excel Tables</dt>
              <dd
                className="text-2xl font-bold"
                data-testid="inspection-excel-tables"
              >
                {outcome.result.metrics.excelTables}
              </dd>
            </div>
            <div>
              <dt className={metricLabelClass}>Named ranges</dt>
              <dd
                className="text-2xl font-bold"
                data-testid="inspection-named-ranges"
              >
                {outcome.result.metrics.namedRanges}
              </dd>
            </div>
          </dl>

          {hiddenCallout ? (
            <p className={noticeClass} data-testid="hidden-worksheets-callout">
              {hiddenCallout}
            </p>
          ) : null}

          <p className={subheadingClass}>Worksheets</p>
          {description.sheets.length === 0 ? (
            <p className={emptyListClass}>No worksheets matched</p>
          ) : (
            <ul
              className="mt-2 flex flex-col gap-3"
              data-testid="worksheet-list"
            >
              {description.sheets.map((sheet) => {
                const badge = visibilityBadge(sheet.visibility);
                return (
                  <li
                    className={rowClass}
                    data-testid="worksheet-item"
                    key={sheet.name}
                  >
                    <span className="flex flex-wrap items-center gap-2">
                      <span
                        className="truncate font-mono text-sm font-semibold"
                        data-testid="worksheet-name"
                      >
                        {sheet.name}
                      </span>
                      {badge ? (
                        <span
                          className={badgeClass}
                          data-testid="worksheet-visibility"
                        >
                          {badge}
                        </span>
                      ) : null}
                    </span>
                    <p
                      className="mt-1 text-xs text-fd-muted-foreground"
                      data-testid="worksheet-summary"
                    >
                      {worksheetSummary(sheet)}
                    </p>
                    {sheet.columns.length > 0 ? (
                      <ColumnList columns={sheet.columns} />
                    ) : null}
                  </li>
                );
              })}
            </ul>
          )}

          <p className={subheadingClass}>Excel Tables</p>
          {description.excelTables.length === 0 ? (
            <p className={emptyListClass} data-testid="no-excel-tables">
              None in this workbook
            </p>
          ) : (
            <ul
              className="mt-2 flex flex-col gap-2"
              data-testid="excel-table-list"
            >
              {description.excelTables.map((table, position) => (
                <li
                  className={rowClass}
                  data-testid="excel-table-item"
                  key={`${position}:${table.sheet}:${table.name}`}
                >
                  <span className="flex flex-wrap items-baseline gap-x-3 gap-y-1">
                    <span
                      className="font-mono text-sm font-semibold"
                      data-testid="excel-table-name"
                    >
                      {table.name}
                    </span>
                    <span className="text-xs text-fd-muted-foreground">
                      {table.sheet} · {table.range}
                    </span>
                  </span>
                  {table.headers.length > 0 ? (
                    <p className="mt-1 truncate font-mono text-xs text-fd-muted-foreground">
                      {table.headers.join(", ")}
                    </p>
                  ) : null}
                </li>
              ))}
            </ul>
          )}

          <p className={subheadingClass}>Named ranges</p>
          {description.namedRanges.length === 0 ? (
            <p className={emptyListClass} data-testid="no-named-ranges">
              None in this workbook
            </p>
          ) : (
            <ul
              className="mt-2 flex flex-col gap-2"
              data-testid="named-range-list"
            >
              {/*
                Keyed by position first. A workbook-scoped name and a
                sheet-scoped one may share a spelling on the same sheet, which
                is valid in Excel and which the description reports as two
                entries, so name and sheet together are not an identity. The
                list is a pure function of one description and is never
                reordered or edited in place, so position is the stable part
                and the rest of the key is there to keep it readable.
              */}
              {description.namedRanges.map((range, position) => (
                <li
                  className={rowClass}
                  data-testid="named-range-item"
                  key={`${position}:${range.sheet}:${range.name}:${range.ref}`}
                >
                  <span className="flex flex-wrap items-baseline gap-x-3 gap-y-1">
                    <span
                      className="font-mono text-sm font-semibold"
                      data-testid="named-range-name"
                    >
                      {range.name}
                    </span>
                    <span className="text-xs text-fd-muted-foreground">
                      {range.sheet} · {range.ref}
                    </span>
                  </span>
                </li>
              ))}
            </ul>
          )}

          {/*
            The warnings come from the operation's own structured result, so
            the report never invents a sentence about a workbook: hidden
            worksheets left out of the description, or a worksheet with no
            header row for an operation to match columns by, are each described
            once, in the library, and rendered here verbatim.
          */}
          {outcome.result.warnings.length > 0 ? (
            <ul className={noticeClass} data-testid="inspection-warnings">
              {outcome.result.warnings.map((warning) => (
                <li data-testid="inspection-warning" key={warning}>
                  {warning}
                </li>
              ))}
            </ul>
          ) : null}
        </>
      ) : null}
    </section>
  );
}
