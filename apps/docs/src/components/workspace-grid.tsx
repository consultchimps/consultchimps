"use client";

/**
 * The record grid: one workspace table shown as an editable spreadsheet.
 *
 * The grid is Tabulator (`tabulator-tables`, MIT, per docs/adr/0003 Decision 4)
 * used directly, with its stylesheet imported from the installed package so the
 * static export bundles it and no request leaves the machine.
 *
 * Three rules shape everything here.
 *
 * The grid sends the worker two kinds of command, and only one of them is the
 * shell's business. A cell edit changes the workspace, so it is counted while
 * it is in flight and the shell holds New, Open, and leaving for it: an edit
 * that has been sent and not answered is work no file has, even though nothing
 * is marked unsaved until the worker accepts it. A table read changes nothing,
 * so it is not held for; one that races a create or an open is refused as
 * belonging to a workspace that has moved on, and the summary that replaced it
 * brings a fresh read with it. Holding for a read would put a question in front
 * of a visitor with nothing at stake.
 *
 * - The grid owns no data. Rows are read through the workspace worker, which
 *   holds the one `@consultchimps/db` database, and every edit is written back
 *   through it before the grid believes it. What a cell shows after an edit is
 *   the value the database read back, not the text that was typed, so the grid
 *   and the file that will be saved cannot disagree.
 * - The grid enforces no types. A value that does not fit a column is refused
 *   by the library, at the one conversion point every write already passes
 *   through, and the refusal comes back as a message. A second check here could
 *   only ever disagree with the first, so there is none: the editors below
 *   choose an input that suits the column, they do not decide what is valid.
 * - A record is addressed by its Record ID, never by its position. Rows are
 *   indexed by it, edits are sent with it, and reverts are applied through it,
 *   so a sorted or reloaded grid still writes to the record the visitor edited.
 *
 * The Record ID column is deliberately editor-free. It is assigned once
 * (ADR 0003 Decision 8), and a column with no editor is skipped by click, tab,
 * and keyboard navigation alike, so there is no path through the grid that
 * offers to change it. The library refuses the write too, and the database has
 * a trigger under that, which is what makes it true rather than merely
 * discouraged.
 */

import "tabulator-tables/dist/css/tabulator.min.css";

import { describeFailure, sectionClass } from "@/components/tool-kit";
import type {
  MarkWorkspaceChanged,
  ReportPendingEdits,
} from "@/components/workspace-tool";
import { referenceLabel } from "@/lib/workspace-labels";
import {
  WORKSPACE_REFERENCE_LIMIT,
  WORKSPACE_STALE_READ,
  type WorkspaceColumn,
  type WorkspaceReference,
  type WorkspaceSummary,
  type WorkspaceTable,
} from "@/lib/workspace-protocol";
import type { WorkspaceClient } from "@/lib/workspace-worker";
import { isConsultChimpsError } from "@consultchimps/core";
import { RECORD_ID_COLUMN } from "@consultchimps/db";
import type { CellValue, TableRow } from "@consultchimps/tabular";
import { Table2 } from "lucide-react";
import { useEffect, useId, useRef, useState } from "react";
import type {
  CellComponent,
  ColumnDefinition,
  Formatter,
  ListEditorParams,
  RowComponent,
  Tabulator,
} from "tabulator-tables";

/** Render text as a node rather than markup, so no stored value can be HTML. */
function textElement(text: string): HTMLElement {
  const element = document.createElement("span");
  element.textContent = text;
  return element;
}

/**
 * A column title, rendered as text.
 *
 * Tabulator writes a plain `title` into the header with `innerHTML`, and a
 * column name is workspace data: the schema accepts any identifier without a
 * control character or a double quote, angle brackets included. A title
 * formatter is handed the title as its cell value, so returning a node shows
 * the name as the name it is.
 */
const titleElement: Formatter = (cell) =>
  textElement(String(cell.getValue() ?? ""));

/**
 * A foreign-key option label, rendered as text, for the same reason: Tabulator
 * writes a plain label into the list with `innerHTML`, and a label is built
 * from the referenced record's own text. The published types say the formatter
 * returns a string; Tabulator appends an `HTMLElement` when it is given one,
 * which is the only way to keep a label literal.
 */
const labelElement = ((label: string) =>
  textElement(label)) as unknown as ListEditorParams["itemFormatter"];

/** How a value reads in a cell when nothing prettier applies. */
function asText(value: CellValue | undefined): string {
  return value === null || value === undefined ? "" : String(value);
}

/**
 * The rows the grid is showing, as the database last confirmed them.
 *
 * This is what a table referring to itself reads its labels from. Not the
 * snapshot the table was read with, which stops being true the moment a record
 * is renamed on screen, and not what Tabulator holds, which can carry a value
 * the worker has not accepted yet.
 */
export interface LiveRecords {
  /** Every Record ID, in the order the rows are shown. */
  ids: () => readonly string[];
  /** One confirmed value, or undefined when the record is not on screen. */
  value: (recordId: string, column: string) => CellValue | undefined;
}

/**
 * What a foreign-key column offers and what it shows, from one place.
 *
 * The picker's options and the cell's label are the same question asked twice,
 * so they are answered by one object. Wiring them to two sources is how a cell
 * comes to show a name the picker no longer offers.
 */
interface ReferenceSource {
  /** The records the picker offers, at most `WORKSPACE_REFERENCE_LIMIT`. */
  options: () => readonly WorkspaceReference[];
  /** How a stored Record ID reads in a cell. */
  label: (value: string) => string;
  /** Whether more records exist than the picker can offer. */
  truncated: () => boolean;
}

function referenceSource(
  references: NonNullable<WorkspaceColumn["references"]>,
  live: LiveRecords,
): ReferenceSource {
  if (references.kind === "onScreen") {
    // Read each time it is asked, never captured: these rows change under the
    // visitor's hands, and an edit to the naming column has to show in every
    // cell pointing at that record at once.
    const { labelColumn } = references;
    const label = (value: string): string =>
      referenceLabel(
        value,
        labelColumn === null ? null : live.value(value, labelColumn),
      );
    return {
      options: () =>
        live
          .ids()
          .slice(0, WORKSPACE_REFERENCE_LIMIT)
          .map((value) => ({ value, label: label(value) })),
      label,
      truncated: () => live.ids().length > WORKSPACE_REFERENCE_LIMIT,
    };
  }
  // A table that is not on screen cannot change while this one is shown, so its
  // records are read once and hold for this generation. Both answers come from
  // that one array.
  const { records, truncated } = references;
  const labels = new Map(
    records.map((reference) => [reference.value, reference.label]),
  );
  return {
    options: () => records,
    // A Record ID past the cap has no label to show, so it shows as itself.
    label: (value) => labels.get(value) ?? value,
    truncated: () => truncated,
  };
}

/**
 * The editor a column gets. Each is Tabulator's own: the point is to hand the
 * visitor an input that suits the column, not to police what they put in it.
 *
 * `editable` is asked on every attempt to open an editor, so the page can lock
 * the grid while a workspace command is in flight without the grid being
 * rebuilt. That is a lifecycle question, not a validity one: what a cell may
 * hold is still the library's answer alone.
 */
function gridColumn(
  column: WorkspaceColumn,
  editable: () => boolean,
  live: LiveRecords,
): ColumnDefinition {
  const base: ColumnDefinition = {
    title: column.name,
    titleFormatter: titleElement,
    field: column.name,
    minWidth: 120,
    editable,
  };

  // A foreign key stores the referenced Record ID and shows a readable label.
  // The list editor searches the labels and writes the id behind them.
  if (column.references !== null) {
    const source = referenceSource(column.references, live);
    return {
      ...base,
      editor: "list",
      // A function, so the options are built when the editor opens rather than
      // when the grid was. Tabulator calls it for each edit; the published
      // types describe only the plain object it also accepts.
      editorParams: (() => ({
        values: source.options().map((reference) => ({
          label: reference.label,
          value: reference.value,
        })),
        autocomplete: true,
        listOnEmpty: true,
        clearable: column.nullable,
        // Past the option cap the list cannot name every record, so a Record ID
        // can still be typed. Whether it exists is the database's answer, not
        // this editor's.
        freetext: source.truncated(),
        placeholderEmpty: "No matching record",
        itemFormatter: labelElement,
      })) as unknown as ColumnDefinition["editorParams"],
      formatter: (cell: CellComponent) => {
        const value = cell.getValue() as CellValue | undefined;
        const text = asText(value);
        return textElement(text === "" ? "" : source.label(text));
      },
    };
  }

  switch (column.type) {
    case "boolean":
      return {
        ...base,
        editor: "tickCross",
        // A nullable boolean has three states, so the editor cycles through
        // them rather than pretending an empty cell is false.
        editorParams: { tristate: column.nullable },
        formatter: "tickCross",
        formatterParams: { allowEmpty: true },
        hozAlign: "center",
      };
    case "integer":
    case "real":
      return { ...base, editor: "number", hozAlign: "right" };
    case "date":
    case "text":
      return { ...base, editor: "input" };
  }
}

function gridColumns(
  table: WorkspaceTable,
  editable: () => boolean,
  live: LiveRecords,
): ColumnDefinition[] {
  return [
    {
      title: "Record ID",
      titleFormatter: titleElement,
      field: RECORD_ID_COLUMN,
      headerTooltip:
        "Assigned once when the record is created, and never changes",
      width: 150,
    },
    ...table.columns.map((column) => gridColumn(column, editable, live)),
  ];
}

export interface WorkspaceGridProps {
  /**
   * Hands back the client that owns the workspace database. It is a function
   * rather than the client itself because the page holds it in a ref, and a ref
   * is only safe to read where this component reads it: inside an effect.
   */
  readonly getClient: () => WorkspaceClient;
  /**
   * Whether the shell is holding the workspace still: any lifecycle or import
   * command is in flight, or a confirmation is waiting for an answer. While it
   * is, the grid opens no editor, cancels one that is open, and refuses an edit
   * that commits in the same instant, so a visitor is never invited to make an
   * edit the worker would refuse and none can slip in between a save's snapshot
   * and its notice. The grid keeps no notion of busy of its own.
   */
  readonly locked: boolean;
  /**
   * The shell's single unsaved-changes setter, called once per edit the worker
   * accepted and never for one it refused. The grid tracks no dirtiness itself.
   */
  readonly markChanged: MarkWorkspaceChanged;
  /**
   * How many edits have been sent to the worker and not answered. An edit is
   * only unsaved once it is accepted, so between sending and the reply the
   * workspace still reads as clean; this is what lets the shell hold New, Open,
   * and leaving for an edit that is still on its way. It is the shell's guard
   * that reads it, not a second idea of unsaved kept here.
   */
  readonly onEditsPending: ReportPendingEdits;
  /** Where a refusal is shown. Called with null once an edit succeeds. */
  readonly onError: (message: string | null) => void;
  /**
   * The workspace as the shell knows it. This is the only place the grid learns
   * which tables it may show and which database they belong to, so a table an
   * import has just created appears here the moment the shell hears about it,
   * and the generation the worker checks can never disagree with the listing.
   */
  readonly summary: WorkspaceSummary;
}

/** What the last completed read produced, and what it was a read of. */
interface LoadedTable {
  readonly table: string;
  /** The workspace it came from, so a read of a replaced one is recognised. */
  readonly generation: number;
  /** The table, or null when the read was refused. */
  readonly data: WorkspaceTable | null;
}

export function WorkspaceGrid({
  getClient,
  locked,
  markChanged,
  onEditsPending,
  onError,
  summary,
}: WorkspaceGridProps) {
  const containerRef = useRef<HTMLDivElement | null>(null);
  // Read by Tabulator's editable callback whenever an editor is about to open,
  // so the lock takes effect without rebuilding the grid.
  const lockedRef = useRef(locked);
  // The cell whose editor is open, so the lock can close it.
  const editingRef = useRef<CellComponent | null>(null);
  // Edits sent and not answered. It lives here rather than inside the effect
  // that builds the grid, so switching tables while one is in flight does not
  // lose count of it.
  const inFlightRef = useRef(0);
  const selectId = useId();

  // Which table the visitor picked. A preference, not a fact: which tables exist
  // is the summary's answer, and this is honoured only while it names one.
  const [chosen, setChosen] = useState<string | null>(null);
  const [loaded, setLoaded] = useState<LoadedTable | null>(null);

  const generation = summary.generation;
  // The one table listing. An import adds to it and a new workspace replaces it,
  // without the grid asking anyone a second time.
  const tables = summary.tables.map((table) => table.name);
  // Derived, not stored: a pick the workspace no longer has (a different file
  // was opened) falls back to the first, with no reset path to remember.
  const selected =
    chosen !== null && tables.includes(chosen) ? chosen : (tables[0] ?? null);

  // The selected table's columns and rows, re-read whenever the workspace moves
  // on: a create, an open, and an import all move the generation, and an import
  // can have added rows to the very table on screen.
  useEffect(() => {
    if (selected === null) {
      return;
    }
    let cancelled = false;
    void getClient()
      .readTable(selected, generation)
      .then((table) => {
        if (!cancelled) {
          setLoaded({ table: selected, generation, data: table });
        }
      })
      .catch((caught: unknown) => {
        if (cancelled) {
          return;
        }
        setLoaded({ table: selected, generation, data: null });
        // A read refused because the workspace moved on is not worth reporting:
        // a newer summary is already on its way here with the read that replaces
        // this one. Anything else is a real failure.
        if (
          isConsultChimpsError(caught) &&
          caught.code === WORKSPACE_STALE_READ
        ) {
          return;
        }
        onError(describeFailure(caught));
      });
    return () => {
      cancelled = true;
    };
  }, [generation, getClient, onError, selected]);

  // A grid that goes away with edits still in flight must not leave the shell
  // holding for answers nothing is waiting on any more.
  useEffect(
    () => () => {
      inFlightRef.current = 0;
      onEditsPending(0);
    },
    [onEditsPending],
  );

  // The lock is a ref so the grid does not have to be rebuilt to honour it, and
  // an effect so an editor already open is closed rather than left to commit
  // after the page has taken its snapshot.
  useEffect(() => {
    lockedRef.current = locked;
    if (locked) {
      editingRef.current?.cancelEdit();
      editingRef.current = null;
    }
  }, [locked]);

  // Derived rather than stored, so switching tables needs no state reset: the
  // grid shows a table only while it is the selected one, and is reading until
  // an answer for that table has come back.
  const showing =
    loaded !== null &&
    loaded.table === selected &&
    loaded.generation === generation;
  const data = showing ? loaded.data : null;
  const loading = selected !== null && !showing;

  // Build the grid for the table now loaded, and tear it down when the table
  // changes or the component unmounts.
  useEffect(() => {
    const element = containerRef.current;
    if (data === null || element === null) {
      return;
    }

    const client = getClient();
    let instance: Tabulator | null = null;
    let destroyed = false;

    // The last value the database confirmed for each record.
    //
    // The rows this is built from are a snapshot of one generation, and this is
    // what carries them forward: an accepted edit updates it, so it stays the
    // truth for as long as the grid is showing that generation. Everything in
    // here that has to be current reads it rather than the rows.
    //
    // It is what a refused edit reverts to as well. Reverting to the cell's
    // previous value would be wrong the moment two edits to one cell overlap:
    // the second one's "previous" is the first one's unstored text.
    const committed = new Map<string, TableRow>(
      data.rows.map((row) => [asText(row[RECORD_ID_COLUMN]), { ...row }]),
    );
    // What a table referring to itself reads its labels from. `committed` is
    // insertion ordered, which is row order, and holds what the database
    // confirmed, so a label changes exactly when the edit behind it was
    // accepted.
    const live: LiveRecords = {
      ids: () => [...committed.keys()],
      value: (recordId, columnName) => committed.get(recordId)?.[columnName],
    };
    // The columns whose value names a record of this same table. Editing one
    // changes what every foreign-key cell pointing at that record shows, and
    // nothing else on screen would say so.
    const namingColumns = new Set(
      data.columns.flatMap((column) =>
        column.references?.kind === "onScreen" &&
        column.references.labelColumn !== null
          ? [column.references.labelColumn]
          : [],
      ),
    );

    // Edits are numbered per cell so a reply that has been overtaken by a newer
    // edit to the same cell reports its outcome without touching the grid.
    const latest = new Map<string, number>();
    let sequence = 0;
    // Putting a value back is itself a cell change, so the edit handler ignores
    // anything written while this is raised.
    let applying = 0;

    // Set when a naming column was edited while an editor was open, so the
    // refresh below can wait for it rather than pull the editor's element out
    // from under it and take the edit in progress with it.
    let labelsStale = false;

    /**
     * Show the labels again after the record they name was renamed.
     *
     * Only the cells have to be told: the picker builds its options when it
     * opens, so it is never stale. A rendered cell keeps whatever the formatter
     * last returned, and nothing about a row pointing at the renamed record has
     * itself changed, so nothing would redraw it.
     */
    const refreshLabels = (): void => {
      if (destroyed || instance === null) {
        return;
      }
      if (editingRef.current !== null) {
        labelsStale = true;
        return;
      }
      labelsStale = false;
      for (const row of instance.getRows()) {
        row.reformat();
      }
    };

    // Tell the shell how many edits are on their way, before the command is
    // sent and after it is answered. Never clamped below zero, so an edit that
    // outlives the grid it was made in cannot drive the count negative and
    // release a hold something else is waiting on.
    const countInFlight = (delta: number): void => {
      inFlightRef.current = Math.max(0, inFlightRef.current + delta);
      onEditsPending(inFlightRef.current);
    };

    // Put a cell back to the last value the database confirmed. A record with
    // no baseline is left alone rather than blanked, so a missing entry could
    // never turn a refusal into data loss.
    const restore = (cell: CellComponent, recordId: string): void => {
      const row = committed.get(recordId);
      if (row === undefined) {
        return;
      }
      applying += 1;
      try {
        cell.setValue(row[cell.getField()] ?? null);
      } finally {
        applying -= 1;
      }
    };

    const persist = (cell: CellComponent): void => {
      if (applying > 0) {
        return;
      }
      const recordId = String(cell.getRow().getIndex());
      // The lock closes an open editor, and the editable callback stops a new
      // one opening, so reaching here while locked takes an edit committed in
      // the same instant the page became busy. Refuse it: written now it would
      // land after a save has taken its snapshot, and the file would not hold
      // what the grid shows.
      if (lockedRef.current) {
        restore(cell, recordId);
        onError(
          "That edit was not made because the workspace was busy. Make it again now the workspace is ready",
        );
        return;
      }
      const column = cell.getField();
      // A separator no identifier can hold, so two cells cannot share a key.
      const key = `${recordId}\u0000${column}`;
      sequence += 1;
      const mine = sequence;
      latest.set(key, mine);

      // An emptied cell means no value, not the empty string. Tabulator's input
      // editors hand back "" when a visitor clears one, and "" is a real text
      // value that a foreign key or a date column would then have to carry.
      const raw = cell.getValue() as CellValue | undefined;
      const value: CellValue = raw === "" || raw === undefined ? null : raw;

      // Counted before the command is posted, and while still inside the event
      // that committed the edit. The editor commits on blur, so a click on New
      // or Open is what commits the edit it is about to replace, and the shell
      // has to be holding by the time that click is handled.
      countInFlight(1);
      void client
        .updateCell({
          // The workspace these rows came from. The worker refuses the edit if
          // that is no longer the workspace it holds, rather than applying it to
          // whichever record in the new one happens to carry this Record ID.
          generation: data.generation,
          table: data.name,
          recordId,
          column,
          value,
        })
        .then((stored) => {
          const row = committed.get(recordId);
          if (row !== undefined) {
            row[column] = stored;
          }
          // A record's name is what other rows show for it, so renaming one
          // changes cells this edit never touched.
          if (namingColumns.has(column)) {
            refreshLabels();
          }
          // The worker took it, so the workspace differs from its file. The
          // shell owns that flag; this is the one place the grid touches it, and
          // only here, after the reply, never on the way out. A refused edit
          // falls to the catch below and marks nothing. No summary is passed:
          // an edit changes what a table holds, not which tables exist.
          markChanged();
          onError(null);
        })
        .catch((caught: unknown) => {
          onError(describeFailure(caught));
        })
        .finally(() => {
          // Released whatever became of it, and after the branches above: an
          // accepted edit has already marked the workspace unsaved, so the
          // shell's hold never lapses between the two reasons for it.
          countInFlight(-1);
          if (destroyed || latest.get(key) !== mine) {
            return;
          }
          latest.delete(key);
          // One path for both outcomes: the cell is set to what the database
          // holds, which a success has just updated and a refusal has not.
          restore(cell, recordId);
        });
    };

    void (async () => {
      const { TabulatorFull } = await import("tabulator-tables");
      if (destroyed) {
        return;
      }
      instance = new TabulatorFull(element, {
        // Copies, because Tabulator writes into the row objects it is given and
        // `committed` above is built from the same rows.
        data: data.rows.map((row) => ({ ...row })),
        columns: gridColumns(data, () => !lockedRef.current, live),
        // A field is a column name and nothing else. Tabulator otherwise reads
        // a dot in a field as a path into nested data, so a legal column such
        // as "billing.address" would render blank and its edits would go to a
        // property nobody asked for. The schema accepts the dot, so the grid
        // has to address it literally.
        nestedFieldSeparator: false,
        // Rows are keyed by Record ID, so getRow and every edit address a
        // record rather than a position in an array.
        index: RECORD_ID_COLUMN,
        layout: "fitColumns",
        maxHeight: "60vh",
        // A single click opens the editor. Range selection in later work may
        // want this on double click instead, which is Tabulator's advice when
        // dragging a selection and starting an edit share the mouse.
        editTriggerEvent: "click",
        placeholder: "This table has no records yet",
        rowFormatter: (row: RowComponent) => {
          // The Record ID on the row element, so a test or an assistive tool
          // can find a record without counting rows.
          row.getElement().dataset["recordId"] = asText(
            row.getData()[RECORD_ID_COLUMN] as CellValue | undefined,
          );
        },
      });
      // Registered after construction because a cell edit is a table event in
      // Tabulator 6, not a table option.
      instance.on("cellEdited", persist);
      // Which cell is being edited, so the lock can close an editor that is
      // already open rather than let it commit after the page has moved on.
      instance.on("cellEditing", (cell: CellComponent) => {
        editingRef.current = cell;
      });
      // Registered after `persist`, so by the time these run the edit has been
      // handed on and the editor is closed: a refresh that was waiting for it
      // can go ahead.
      instance.on("cellEdited", () => {
        editingRef.current = null;
        if (labelsStale) {
          refreshLabels();
        }
      });
      instance.on("cellEditCancelled", () => {
        editingRef.current = null;
        if (labelsStale) {
          refreshLabels();
        }
      });
    })();

    return () => {
      destroyed = true;
      editingRef.current = null;
      instance?.destroy();
    };
  }, [data, getClient, markChanged, onEditsPending, onError]);

  // Only a snapshot can be short of records: a table referring to itself offers
  // the rows on screen, and the note below says so only when there are more of
  // those than the picker takes.
  const truncated =
    data?.columns.filter((column) =>
      column.references?.kind === "snapshot"
        ? column.references.truncated
        : column.references?.kind === "onScreen" &&
          data.rows.length > WORKSPACE_REFERENCE_LIMIT,
    ) ?? [];

  return (
    <section className={sectionClass} data-testid="workspace-grid-section">
      <div className="flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-2">
          <Table2
            aria-hidden="true"
            className="size-5 shrink-0 text-fd-primary"
          />
          <h2 className="text-xl font-bold tracking-[-0.03em]">Records</h2>
        </div>
        {tables.length > 0 ? (
          <label className="flex items-center gap-2 text-sm" htmlFor={selectId}>
            Table
            <select
              className="rounded-lg border bg-fd-card px-3 py-2 text-sm outline-none focus-visible:ring-2 focus-visible:ring-fd-ring"
              data-testid="workspace-table-select"
              disabled={locked}
              id={selectId}
              onChange={(event) => setChosen(event.target.value)}
              value={selected ?? ""}
            >
              {tables.map((name) => (
                <option key={name} value={name}>
                  {name}
                </option>
              ))}
            </select>
          </label>
        ) : null}
      </div>

      {tables.length === 0 ? (
        <p
          className="mt-4 text-sm text-fd-muted-foreground"
          data-testid="workspace-grid-empty"
        >
          This workspace has no tables yet
        </p>
      ) : (
        <p className="mt-3 text-sm text-fd-muted-foreground">
          Click a cell to edit it. Each edit is written to the workspace as you
          make it, and Save writes the workspace back to its file. The Record ID
          is assigned once and cannot be edited
        </p>
      )}

      {loading ? (
        <p
          className="mt-4 text-sm text-fd-muted-foreground"
          data-testid="workspace-grid-loading"
        >
          Reading the workspace
        </p>
      ) : null}

      <div
        className="mt-4 empty:hidden"
        data-testid="workspace-grid"
        ref={containerRef}
      />

      {truncated.length > 0 ? (
        <p
          className="mt-3 text-xs text-fd-muted-foreground"
          data-testid="workspace-grid-reference-note"
        >
          {`The picker for ${truncated
            .map((column) => column.name)
            .join(
              ", ",
            )} lists the first records of the related table, not all of them. A Record ID beyond that list can still be typed in`}
        </p>
      ) : null}
    </section>
  );
}
