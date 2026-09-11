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
 *
 * ## The spreadsheet gestures
 *
 * Range selection, the clipboard, and the fill handle are the grid's Excel-grade
 * half, and they follow one rule: a gesture is one movement, so it is one
 * command. Every one of them is planned whole before anything is sent
 * (`lib/workspace-gesture`), written by one batched `updateCells`
 * (`lib/workspace-protocol`), and reported to the shell as exactly one
 * `editSent` and one `editSettled`, because it is one unit of work in flight
 * however many cells it covers.
 *
 * Three pieces of it are ours rather than Tabulator's, and each is here for a
 * reason found in its source at 6.5.2 rather than by preference:
 *
 * - **The fill handle**, entirely: Tabulator has none, not even a copying one.
 * - **The paste**, because its built-in `range` paste action writes through
 *   `row.updateData`, which reaches cells by `setValueProcessData` and therefore
 *   never dispatches `cellEdited`. The grid's only path to the database is
 *   `cellEdited`, so that action would repaint the grid and save nothing. Ours
 *   plans the block and sends the one command. Its built-in `range` paste parser
 *   is not used either: it splits on "
" and leaves the "
" of a Windows copy
 *   on the last field of every row.
 * - **The copy**, because `generatePlainContent` quotes nothing, and a text
 *   column here may hold a tab or a newline, so such a cell would corrupt the
 *   clipboard exactly as the stray "
" corrupts a paste. It also emits a header
 *   row, which a range copy should not. Ours is `lib/workspace-tsv`, one grammar
 *   read in both directions.
 *
 * Two Tabulator options are deliberately left off. `selectableRangeClearCells`
 * would let Delete clear a range through `cell.setValue` per cell, which is one
 * gesture becoming as many commands as it has cells. Header sorting is off
 * because a header click now selects the column, and one click with two meanings
 * is how a visitor loses a selection they were building.
 */

import "tabulator-tables/dist/css/tabulator.min.css";

import {
  compactButtonClass,
  describeFailure,
  noticeClass,
  sectionClass,
} from "@/components/tool-kit";
import type {
  ReportEditorClosed,
  ReportEditorOpened,
  ReportEditSent,
  ReportEditSettled,
  ReportGridDetached,
} from "@/components/workspace-tool";
import {
  attachFillHandle,
  rangeRect,
  type FillHandleDrag,
} from "@/components/workspace-fill-handle";
import {
  answerFailure,
  answerFailures,
  cellKey,
  dismissFailures,
  failureReport,
  failuresAt,
  NOTHING_REFUSED,
  recordFailure,
  recordFailures,
  tableKey,
  type RecordedFailures,
} from "@/lib/workspace-cell-errors";
import {
  busyRefusal,
  ONE_RECTANGLE_ONLY,
  planFill,
  planPaste,
  type GestureGrid,
  type GestureKind,
  type GestureRect,
  type PlannedWrite,
} from "@/lib/workspace-gesture";
import { cellText, encodeTsv, parseTsv } from "@/lib/workspace-tsv";
import { referenceLabel } from "@/lib/workspace-labels";
import {
  WORKSPACE_REFERENCE_LIMIT,
  WORKSPACE_STALE_READ,
  type WorkspaceColumn,
  type WorkspaceReference,
  type WorkspaceSummary,
  type WorkspaceTable,
} from "@/lib/workspace-protocol";
import type {
  WorkspaceCellOutcome,
  WorkspaceClient,
} from "@/lib/workspace-worker";
import { isConsultChimpsError } from "@consultchimps/core";
import { RECORD_ID_COLUMN } from "@consultchimps/db";
import type { CellValue, TableRow } from "@consultchimps/tabular";
import { Table2 } from "lucide-react";
import { useEffect, useId, useLayoutEffect, useRef, useState } from "react";
import type {
  CellComponent,
  ColumnDefinition,
  Formatter,
  ListEditorParams,
  Options,
  RangeComponent,
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
        const text = cellText(value);
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

/**
 * Tabulator's paste pipeline, wired to our own parser and our own action.
 *
 * Only the paste half of its clipboard is enabled: `clipboard: "paste"`
 * registers the paste listener and leaves the copy to the listener this grid
 * adds itself. The paste is worth keeping because of what surrounds it, the
 * origin check that leaves a paste inside an open editor alone, not because of
 * what it does with the text.
 *
 * Both functions are cast past the published types, which describe the parser as
 * returning rows and the action as one of four names, while Tabulator accepts a
 * function for either (`Clipboard.setPasteParser`, `setPasteAction`). The parser
 * deliberately returns an object rather than an array: the clipboard module runs
 * an array through the mutator module before handing it on, and a block of text
 * is not row data for a mutator to touch.
 */
function pasteOptions(apply: (block: string[][]) => void): Partial<Options> {
  return {
    clipboard: "paste",
    clipboardPasteParser: (text: string) => ({ block: parseTsv(text) }),
    clipboardPasteAction: (parsed: { block: string[][] }) => {
      apply(parsed.block);
      // What Tabulator would report as the pasted rows. Ours are reported by the
      // worker's reply instead, so there is nothing to hand back here.
      return [];
    },
  } as unknown as Partial<Options>;
}

export interface WorkspaceGridProps {
  /**
   * Hands back the client that owns the workspace database. It is a function
   * rather than the client itself because the page holds it in a ref, and a ref
   * is only safe to read where this component reads it: inside an effect.
   */
  readonly getClient: () => WorkspaceClient;
  /**
   * Whether a command is in flight, so the workspace is being held still.
   * While it is, the grid opens no editor, cancels one that is open, and
   * refuses an edit that commits in the same instant, so a visitor is never
   * invited to make an edit the worker would refuse and none can slip in
   * between a save's snapshot and its notice. The grid keeps no notion of busy
   * of its own.
   *
   * A confirmation waiting for an answer deliberately does not lock: it is an
   * inline question rather than a modal, and locking for it cancelled the very
   * editor the question had been raised about, which is #174.
   */
  readonly locked: boolean;
  /**
   * An edit has been sent to the worker. It is only unsaved once the worker
   * accepts it, so between sending and the reply the workspace still reads as
   * clean; this is what lets the shell hold New, Open, and leaving for an edit
   * that is still on its way. The shell counts them, so the grid keeps no
   * count of its own.
   */
  readonly onEditSent: ReportEditSent;
  /**
   * That edit came back, and whether the worker took it. One report for both,
   * so the shell marks the workspace unsaved and releases the edit's hold in
   * one move and the hold cannot lapse between them.
   */
  readonly onEditSettled: ReportEditSettled;
  /**
   * Whether a cell is open for editing. What is typed into an editor and not
   * committed exists only in that input element, so this is the only way the
   * shell can know there is anything to lose. Reported from the moment the
   * editor opens, which is coarser than waiting for a keystroke and cannot miss
   * one; see `WorkspaceHoldState`.
   */
  readonly onEditorOpened: ReportEditorOpened;
  /**
   * That editor closed, and whether this grid closed it on its way out.
   *
   * A visitor's Escape and a teardown reach Tabulator as the same cancel, and
   * they mean opposite things to a confirmation waiting for an answer, so the
   * one thing that can tell them apart says which it was.
   */
  readonly onEditorClosed: ReportEditorClosed;
  /**
   * The grid is going away, taking whatever it was holding with it.
   *
   * Its own report rather than an editor-closed and a zeroed count, because
   * those are indistinguishable from a visitor pressing Escape, and the shell
   * has to tell them apart: a teardown caused by the very navigation a
   * confirmation is standing in front of must not be able to answer it. See the
   * #174 rule in `lib/workspace-state`.
   */
  readonly onDetached: ReportGridDetached;
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
  onDetached,
  onEditorClosed,
  onEditorOpened,
  onEditSent,
  onEditSettled,
  summary,
}: WorkspaceGridProps) {
  const containerRef = useRef<HTMLDivElement | null>(null);
  // The positioned element the fill handle and its outline live in. It has to be
  // outside the grid's own element: Tabulator takes the element it is given as
  // the table itself and lays out its children, so anything appended there is
  // swept away when it builds.
  const frameRef = useRef<HTMLDivElement | null>(null);
  // Read by Tabulator's editable callback whenever an editor is about to open,
  // so the lock takes effect without rebuilding the grid.
  const lockedRef = useRef(locked);
  // The cell whose editor is open, so the lock can close it.
  const editingRef = useRef<CellComponent | null>(null);
  const selectId = useId();

  // Which table the visitor picked. A preference, not a fact: which tables exist
  // is the summary's answer, and this is honoured only while it names one.
  const [chosen, setChosen] = useState<string | null>(null);
  const [loaded, setLoaded] = useState<LoadedTable | null>(null);
  // What has been refused and not yet dealt with, kept by the thing it is about
  // rather than as one slot the next success would clear, and stamped with the
  // workspace it happened in so a new one starts clean without anything having
  // to remember to clear it. See `lib/workspace-cell-errors`.
  const [recorded, setRecorded] = useState<RecordedFailures>(NOTHING_REFUSED);

  // Read by a gesture, which has to know whether a cell it would fill from is
  // one the database refused. A ref because the gesture handlers live in the
  // effect that built the grid, where the state of a later render is not in
  // scope, and it mirrors that state rather than being a second copy of it.
  const recordedRef = useRef(recorded);
  useEffect(() => {
    recordedRef.current = recorded;
  }, [recorded]);

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
        if (cancelled) {
          return;
        }
        setLoaded({ table: selected, generation, data: table });
        // The read this table was waiting on succeeded, so whatever was
        // standing about reading it is answered.
        setRecorded((previous) =>
          answerFailure(previous, generation, tableKey(selected)),
        );
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
        setRecorded((previous) =>
          recordFailure(
            previous,
            generation,
            tableKey(selected),
            describeFailure(caught),
          ),
        );
      });
    return () => {
      cancelled = true;
    };
  }, [generation, getClient, selected]);

  // A grid that goes away with edits still in flight, or with an editor open,
  // must not leave the shell holding for work that went with it.
  //
  // For an open editor this is the last resort, not the usual path: Tabulator
  // reports a cancel whenever an editor goes, including when the grid it is in
  // is torn down, so a table switch or a new workspace releases the hold
  // through the same handler a visitor's Escape does. What is left is the page
  // itself going, which reports nothing, and this covers it.
  //
  // Reported as a detach rather than as an editor closing and a count going to
  // zero, because the shell has to be able to tell this from a visitor pressing
  // Escape: a teardown is the navigation's own doing and may not answer the
  // confirmation that navigation raised.
  //
  // Deliberately not in the effect that builds the grid, which runs again
  // whenever the rows change. Releasing there would be a second answer to the
  // same question, and the one place the two differ is a rebuild caused by the
  // very navigation the hold exists to guard.
  useEffect(() => () => onDetached(), [onDetached]);

  // The lock is a ref so the grid does not have to be rebuilt to honour it, and
  // an effect so an editor already open is closed rather than left to commit
  // after the page has taken its snapshot.
  //
  // A layout effect, not a passive one: the ref is a mirror of the shell's one
  // derived answer, not a second opinion about it, and a mirror that updates
  // after the paint is a mirror that is wrong for a frame. The refusal in
  // `persist` and the `editable` callback both read it, so a frame of lag is a
  // frame in which an edit the page has already decided to refuse is accepted.
  useLayoutEffect(() => {
    lockedRef.current = locked;
    if (locked) {
      // Cancelling dispatches Tabulator's own cancel event, so the editor is
      // reported closed through the one handler rather than here as well.
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
    const frame = frameRef.current;
    if (data === null || element === null || frame === null) {
      return;
    }

    const client = getClient();
    let instance: Tabulator | null = null;
    let destroyed = false;
    let detachHandle: (() => void) | null = null;

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
      data.rows.map((row) => [cellText(row[RECORD_ID_COLUMN]), { ...row }]),
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
    // edit to the same cell reports its outcome without touching the grid. A
    // gesture stamps every cell it writes, so one cell of it can be overtaken
    // without the others losing their answer.
    const latest = new Map<string, number>();
    let sequence = 0;
    // A separator no identifier can hold, so two cells cannot share a key.
    const editKey = (recordId: string, column: string): string =>
      `${recordId}\u0000${column}`;
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
        const column = cell.getField();
        restore(cell, recordId);
        setRecorded((previous) =>
          recordFailure(
            previous,
            data.generation,
            cellKey(data.name, recordId, column),
            "That edit was not made because the workspace was busy. Make it again now the workspace is ready",
          ),
        );
        return;
      }
      const column = cell.getField();
      const key = editKey(recordId, column);
      sequence += 1;
      const mine = sequence;
      latest.set(key, mine);

      // An emptied cell means no value, not the empty string. Tabulator's input
      // editors hand back "" when a visitor clears one, and "" is a real text
      // value that a foreign key or a date column would then have to carry.
      const raw = cell.getValue() as CellValue | undefined;
      const value: CellValue = raw === "" || raw === undefined ? null : raw;

      // Reported before the command is posted, and while still inside the event
      // that committed the edit. The editor commits on blur, so a click on New
      // or Open is what commits the edit it is about to replace, and the shell
      // has to be holding by the time that click is handled.
      onEditSent();
      // What became of it, for the one report that settles it below.
      let accepted = false;
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
          // This cell is dealt with, so what was standing about it goes. Only
          // this cell's: a refusal somewhere else is still true and still the
          // only thing saying why that cell reads as it does.
          setRecorded((previous) =>
            answerFailure(
              previous,
              data.generation,
              cellKey(data.name, recordId, column),
            ),
          );
          // The worker took it, so the workspace differs from its file. The
          // shell owns that flag; the grid only says what happened, and only
          // here, after the reply, never on the way out. A refused edit falls
          // to the catch below and settles as refused.
          accepted = true;
        })
        .catch((caught: unknown) => {
          // Against the cell it was about, which is what stops the next
          // success elsewhere erasing it. A stale-generation refusal and a
          // value the database will not hold arrive here alike.
          setRecorded((previous) =>
            recordFailure(
              previous,
              data.generation,
              cellKey(data.name, recordId, column),
              describeFailure(caught),
            ),
          );
        })
        .finally(() => {
          // One report for both outcomes, after the branches above have said
          // which it was. The shell marks the workspace unsaved and releases
          // this edit's hold in the one move, so the hold never lapses between
          // the two reasons for it.
          onEditSettled(accepted);
          if (destroyed || latest.get(key) !== mine) {
            return;
          }
          latest.delete(key);
          // One path for both outcomes: the cell is set to what the database
          // holds, which a success has just updated and a refusal has not.
          restore(cell, recordId);
        });
    };

    /* ---------------------------------------------------------------------
     * The spreadsheet gestures: the clipboard and the fill handle.
     *
     * Each is planned whole, sent as one command, and reported to the shell as
     * one unit of work. Nothing below decides whether a value is acceptable, and
     * nothing below writes to a cell except through the `applying` guard, so the
     * one path from a gesture to the database stays the one path.
     * ------------------------------------------------------------------ */

    /** The columns as a gesture sees them, in the order they are shown. */
    const fields = [
      RECORD_ID_COLUMN,
      ...data.columns.map((column) => column.name),
    ];
    const writableColumn = (field: string): boolean =>
      field !== RECORD_ID_COLUMN;
    // A foreign key holds Record IDs, so a fill copies them: reading the
    // trailing integer of one as a counter would point rows at records nobody
    // chose, and those records may well exist.
    const referenceColumns = new Set(
      data.columns
        .filter((column) => column.references !== null)
        .map((column) => column.name),
    );

    /**
     * Record IDs in the order the rows are on screen, which is what a rectangle
     * indexes. Read from Tabulator rather than from `committed`, so the grid's
     * own idea of where a row is stays the only one.
     */
    const displayedRecords = (): string[] =>
      instance === null
        ? []
        : instance.getRows("active").map((row) => String(row.getIndex()));

    const gridForGesture = (): GestureGrid => ({
      columns: fields.map((field) => ({
        field,
        writable: writableColumn(field),
        series: !referenceColumns.has(field),
      })),
      recordIds: displayedRecords(),
      // What the database confirmed, never what Tabulator holds: a cell carrying
      // an edit nobody has answered yet is not a value to copy or to fill from.
      text: (recordId, field) => cellText(committed.get(recordId)?.[field]),
      refused: (recordId, field) =>
        failuresAt(recordedRef.current, data.generation).has(
          cellKey(data.name, recordId, field),
        ),
    });

    /**
     * Every selected rectangle, as grid indices, in the order they were made, or
     * null when one of them names a row or a column this grid is not showing.
     *
     * Null rather than the ones that did resolve: dropping one would turn a
     * selection of two rectangles into a gesture on one, which is the opposite
     * of the rule that a gesture acts on a single rectangle or on none.
     */
    const selectedRects = (): GestureRect[] | null => {
      if (instance === null) {
        return null;
      }
      const geometry = { columns: fields, rows: displayedRecords() };
      const rects = instance
        .getRanges()
        .map((range: RangeComponent) => rangeRect(range, geometry));
      return rects.every((rect): rect is GestureRect => rect !== null)
        ? rects
        : null;
    };

    /**
     * The cell a gesture's explanation belongs to: the top left of what it was
     * about. A gesture refused as a whole has no single cell of its own, and this
     * is the cell the visitor started from, so the sentence sits where they are
     * looking and is cleared by their next attempt there.
     */
    const anchorAt = (
      rect: GestureRect | undefined,
    ): { recordId: string; column: string } | null => {
      if (rect === undefined) {
        return null;
      }
      const recordId = displayedRecords()[rect.top];
      const column = fields[rect.left];
      return recordId === undefined || column === undefined
        ? null
        : { recordId, column };
    };

    const reportGesture = (
      anchor: { recordId: string; column: string } | null,
      message: string,
    ): void => {
      setRecorded((previous) =>
        recordFailure(
          previous,
          data.generation,
          // Against the cell the gesture started from, or against the table when
          // there is no such cell: a selection this grid cannot read has no cell
          // to name, and a refusal nobody is told about is worse than one filed
          // against the table it happened in.
          anchor === null
            ? tableKey(data.name)
            : cellKey(data.name, anchor.recordId, anchor.column),
          message,
        ),
      );
    };

    /**
     * Show what the database stored, without that counting as an edit.
     *
     * The value comes from the reply rather than from `committed`, and a record
     * with no baseline is left alone rather than blanked, which is the rule
     * `restore` keeps for the same reason: a missing entry must never turn a
     * reply into data loss.
     */
    const paint = (
      recordId: string,
      column: string,
      value: CellValue,
      stamp: number,
    ): void => {
      if (destroyed || instance === null || !committed.has(recordId)) {
        return;
      }
      if (latest.get(editKey(recordId, column)) !== stamp) {
        return;
      }
      // Tabulator answers `false` for a row or a cell it does not have, which
      // the published types do not say, so both are read as what they can be: a
      // record can have been reloaded away from under a reply.
      const row = instance.getRow(recordId) as RowComponent | false;
      if (row === false) {
        return;
      }
      const cell = row.getCell(column) as CellComponent | false;
      if (cell === false) {
        return;
      }
      applying += 1;
      try {
        cell.setValue(value);
      } finally {
        applying -= 1;
      }
    };

    /** What came back for each cell of a gesture, applied in one move. */
    const settleGesture = (
      outcomes: readonly WorkspaceCellOutcome[],
      stamp: number,
    ): void => {
      const answered: string[] = [];
      const refusals: Array<{ key: string; message: string }> = [];
      let renamed = false;
      for (const outcome of outcomes) {
        const key = cellKey(data.name, outcome.recordId, outcome.column);
        if (!outcome.accepted) {
          // Against the cell it was about, so a refusal elsewhere in the same
          // gesture is still the only thing saying why that cell reads as it
          // does.
          refusals.push({ key, message: describeFailure(outcome.error) });
          continue;
        }
        const row = committed.get(outcome.recordId);
        if (row !== undefined) {
          row[outcome.column] = outcome.value;
        }
        answered.push(key);
        if (namingColumns.has(outcome.column)) {
          renamed = true;
        }
        paint(outcome.recordId, outcome.column, outcome.value, stamp);
      }
      // One state update for the whole gesture: the cells it dealt with and the
      // cells it could not, in the order they happened.
      setRecorded((previous) =>
        recordFailures(
          answerFailures(previous, data.generation, answered),
          data.generation,
          refusals,
        ),
      );
      if (renamed) {
        refreshLabels();
      }
    };

    /**
     * Send one gesture: one command, one `editSent`, one `editSettled`.
     *
     * Nothing is painted before the reply. The single-cell path shows the typed
     * value only because the editor has already put it there; here there is no
     * editor, so painting first would be a block of values that some cells then
     * take back.
     */
    const applyGesture = (
      writes: readonly PlannedWrite[],
      anchor: { recordId: string; column: string } | null,
    ): void => {
      if (writes.length === 0) {
        return;
      }
      sequence += 1;
      const stamp = sequence;
      const keys = writes.map((write) => editKey(write.recordId, write.column));
      // Every cell of the gesture is stamped, so a later edit to one of them
      // overtakes this reply for that cell alone.
      for (const key of keys) {
        latest.set(key, stamp);
      }
      // Before the command is posted, and inside the event that caused it: the
      // shell has to be holding for this work by the time anything else is
      // clicked.
      onEditSent();
      let accepted = false;
      void client
        .updateCells({
          generation: data.generation,
          table: data.name,
          writes: writes.map((write) => ({
            recordId: write.recordId,
            column: write.column,
            // An empty cell means no value rather than the empty string, the
            // same rule a cleared editor follows.
            value: write.value === "" ? null : write.value,
          })),
        })
        .then(
          (outcomes) => {
            // Unsaved once, for the gesture, if the database took any of it.
            accepted = outcomes.some((outcome) => outcome.accepted);
            settleGesture(outcomes, stamp);
          },
          // The rejection handler rather than a `catch`, so a fault while
          // applying the reply is not reported as the gesture having been
          // refused: by then it has not been, and the cells are already painted.
          (caught: unknown) => {
            // The gesture was refused whole: a workspace that has moved on, more
            // cells than one step applies, or a worker that has gone. Nothing
            // was written, so every cell still shows what the database holds.
            reportGesture(anchor, describeFailure(caught));
          },
        )
        .finally(() => {
          onEditSettled(accepted);
          for (const key of keys) {
            if (latest.get(key) === stamp) {
              latest.delete(key);
            }
          }
        });
    };

    /**
     * Leave the selection on what a gesture just did, the way a spreadsheet
     * does, so the handle is on the corner of it and a second drag carries on
     * from there.
     *
     * On the movement rather than on the reply: it is the shape the visitor drew
     * and they should see it at once. A gesture the worker then refuses as a
     * whole leaves every value untouched and says so, so the selection is the
     * only thing that moved.
     *
     * A row the grid has not rendered has no cells to bound a range with, so a
     * fill or a paste reaching past the rendered window leaves the selection
     * where it was rather than guessing.
     */
    const selectCovered = (rect: GestureRect | null): void => {
      if (rect === null || instance === null) {
        return;
      }
      const rows = instance.getRows("active");
      const range = instance.getRanges()[0];
      const start = rows[rect.top]?.getCells()[rect.left];
      const end = rows[rect.bottom]?.getCells()[rect.right];
      if (range === undefined || start === undefined || end === undefined) {
        return;
      }
      range.setBounds(start, end);
    };

    /** Refuse a gesture the page is too busy for, without sending anything. */
    const refuseWhileBusy = (
      kind: GestureKind,
      anchor: { recordId: string; column: string } | null,
    ): boolean => {
      if (!lockedRef.current) {
        return false;
      }
      reportGesture(anchor, busyRefusal(kind));
      return true;
    };

    /** Put the selected rectangle on the clipboard as the text Excel reads. */
    const copySelection = (event: ClipboardEvent): void => {
      // An open editor has its own text selection, and copying inside it is the
      // browser's business rather than the grid's.
      if (editingRef.current !== null || instance === null) {
        return;
      }
      const rects = selectedRects();
      // Nothing selected is nothing to copy, and nothing to explain either: the
      // browser does whatever it would have done.
      if (rects === null || rects.length === 0) {
        return;
      }
      const rect = rects[0] as GestureRect;
      if (rects.length !== 1) {
        // Nothing is copied, rather than whichever rectangle came last.
        event.preventDefault();
        reportGesture(anchorAt(rect), ONE_RECTANGLE_ONLY);
        return;
      }
      const forGesture = gridForGesture();
      const rows: string[][] = [];
      for (let row = rect.top; row <= rect.bottom; row += 1) {
        const recordId = forGesture.recordIds[row];
        if (recordId === undefined) {
          return;
        }
        const line: string[] = [];
        for (let column = rect.left; column <= rect.right; column += 1) {
          const field = forGesture.columns[column]?.field;
          if (field === undefined) {
            return;
          }
          line.push(forGesture.text(recordId, field));
        }
        rows.push(line);
      }
      if (event.clipboardData === null) {
        return;
      }
      event.clipboardData.setData("text/plain", encodeTsv(rows));
      event.preventDefault();
    };

    /** Write a pasted block into the selection, or say why it was not written. */
    const pasteBlock = (block: string[][]): void => {
      const rects = selectedRects();
      const anchor = anchorAt(rects?.[0]);
      if (refuseWhileBusy("paste", anchor)) {
        return;
      }
      const plan = planPaste({
        grid: gridForGesture(),
        // An unreadable selection is not one rectangle, and the planner says so
        // in the one sentence that covers both.
        ranges: rects ?? [],
        block,
      });
      if (plan.kind === "refused") {
        reportGesture(anchor, plan.reason);
        return;
      }
      applyGesture(plan.writes, anchor);
      selectCovered(plan.covered);
    };

    /** Extend the selection's values into the cells the drag covered. */
    const fillFrom = (drag: FillHandleDrag): void => {
      const anchor = anchorAt(drag.ranges[0]);
      if (refuseWhileBusy("fill", anchor)) {
        return;
      }
      const plan = planFill({
        grid: gridForGesture(),
        ranges: drag.ranges,
        pointer: drag.pointer,
      });
      if (plan.kind === "refused") {
        reportGesture(anchor, plan.reason);
        return;
      }
      applyGesture(plan.writes, anchor);
      selectCovered(plan.covered);
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
        // A rectangle by drag or shift, several by ctrl-click, a whole column by
        // its header, and keyboard extension with shift-arrow and
        // ctrl-shift-arrow. All of it Tabulator's own.
        selectableRange: true,
        selectableRangeColumns: true,
        // Deliberately not selectableRangeRows: it makes the first column
        // Tabulator's row header, which is excluded from every cell range, and
        // the Record ID has to stay selectable so that it can be copied.
        //
        // Deliberately not selectableRangeClearCells either: Delete would clear
        // a range through one setValue per cell, which turns one gesture into one
        // command per cell.
        //
        // A double click opens the editor now that a drag selects, which is
        // Tabulator's own advice: the two would otherwise share the mouse. Enter
        // on the active cell opens it too.
        editTriggerEvent: "dblclick",
        // A header click selects the column, so it must not also sort. One click
        // with two meanings is how a visitor loses the selection they were
        // building, and the order rows are shown in is what a rectangle indexes.
        columnDefaults: { headerSort: false },
        ...pasteOptions(pasteBlock),
        placeholder: "This table has no records yet",
        rowFormatter: (row: RowComponent) => {
          // The Record ID on the row element, so a test or an assistive tool
          // can find a record without counting rows.
          row.getElement().dataset["recordId"] = cellText(
            row.getData()[RECORD_ID_COLUMN] as CellValue | undefined,
          );
        },
      });
      // Every time an editor closes or a cell is clicked, range selection puts
      // focus back on the rows, and a focus that scrolls moves the page under
      // whatever the visitor is reaching for: the click that closed the editor
      // on its way down lands somewhere else on its way up, so New, Open, a link
      // out of the page, and Back never see a click at all. The focus is wanted,
      // the scrolling is not.
      //
      // Every focus of this element, not only those two: it is a 60vh scroll
      // container that is nearly always on screen already, and a focus that
      // jumps the page to it is unhelpful wherever it comes from.
      //
      // On `tableBuilt`, because the table is built in a later task than the
      // constructor (Tabulator defers `_create` so that these very handlers can
      // be registered first) and building empties this element and appends the
      // row area to it. Looking for the row area now would find nothing.
      instance.on("tableBuilt", () => {
        const rowArea = element.querySelector(".tabulator-tableholder");
        if (rowArea instanceof HTMLElement) {
          rowArea.focus = function focusWithoutScrolling(
            options?: FocusOptions,
          ): void {
            HTMLElement.prototype.focus.call(this, {
              ...options,
              preventScroll: true,
            });
          };
        }
      });
      // Registered after construction because a cell edit is a table event in
      // Tabulator 6, not a table option.
      instance.on("cellEdited", persist);
      // Which cell is being edited, so the lock can close an editor that is
      // already open rather than let it commit after the page has moved on.
      instance.on("cellEditing", (cell: CellComponent) => {
        editingRef.current = cell;
        onEditorOpened();
      });
      // Every way an editor closes: committed, cancelled with Escape, cancelled
      // by the lock engaging, and this grid being destroyed, all of which
      // Tabulator reports as a cancel alike. Registered after `persist`, so by
      // the time these run the edit has been handed on and the editor is
      // closed: a refresh that was waiting for it can go ahead.
      //
      // `destroyed` is what separates the last of those from the rest, and it
      // is why the report carries it. An editor that went because this grid was
      // torn down is the navigation's own doing and may not answer a standing
      // confirmation; a visitor's Escape is an answer. Only the grid knows
      // which happened, so only the grid can say. See the #174 rule in
      // `lib/workspace-state`.
      //
      // A rebuild for new rows destroys the instance too, and reports the same
      // cause, which is deliberate but not free: the draft really has gone, so a
      // question standing over it would go on naming a cell that is no longer
      // there. It cannot arise as things stand, because every change of rows
      // comes from a command, and a command voids the drafts on its way through.
      // Splitting the two would need this grid to know it is unmounting, which
      // is the one thing this cleanup cannot tell.
      const closed = (): void => {
        // A value put back is a cell change like any other, and Tabulator
        // dispatches `cellEdited` for it, so this runs for a reply painting a
        // cell as well as for an editor closing. The guard is what tells the two
        // apart: without it a gesture would report an editor closing once per
        // cell it wrote, and an editor closing is an answer the page acts on.
        if (applying > 0) {
          return;
        }
        editingRef.current = null;
        onEditorClosed(destroyed);
        if (labelsStale) {
          refreshLabels();
        }
      };
      instance.on("cellEdited", closed);
      instance.on("cellEditCancelled", closed);

      // Copy is ours: see the note at the top of this file. Registered on the
      // frame, which every copy inside the grid bubbles up to.
      frame.addEventListener("copy", copySelection);
      // The fill handle, which Tabulator does not have at all.
      detachHandle = attachFillHandle({
        table: instance,
        container: frame,
        geometry: () => ({
          rows: displayedRecords(),
          columns: fields,
          writable: writableColumn,
        }),
        onFill: fillFrom,
      });
    })();

    return () => {
      destroyed = true;
      editingRef.current = null;
      frame.removeEventListener("copy", copySelection);
      detachHandle?.();
      instance?.destroy();
    };
  }, [
    data,
    getClient,
    onEditorClosed,
    onEditorOpened,
    onEditSent,
    onEditSettled,
  ]);

  // Only a snapshot can be short of records: a table referring to itself offers
  // the rows on screen, and the note below says so only when there are more of
  // those than the picker takes.
  // Everything standing in the workspace now held, newest in full. The cells it
  // names are above it, which is the other half of why the grid explains its own
  // refusals rather than handing them to a slot at the foot of the page. A set
  // recorded under a workspace that has been replaced is not shown at all: none
  // of its keys mean the same thing here, and a Record ID from one file is very
  // likely to exist in the next.
  const report = failureReport(failuresAt(recorded, generation));

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
          Double-click a cell to edit it. Click and drag to select a range,
          ctrl-click to add another, and copy or paste a range the way a
          spreadsheet does. Drag the corner of a selection to fill from it. Each
          change is written to the workspace as you make it, and Save writes the
          workspace back to its file. The Record ID is assigned once and cannot
          be edited
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

      {/* Positioned, and outside the grid's own element, because the fill handle
          and its outline are placed in it and Tabulator owns everything inside
          the element it is given. */}
      <div className="relative" ref={frameRef}>
        <div
          className="mt-4 empty:hidden"
          data-testid="workspace-grid"
          ref={containerRef}
        />
      </div>

      {report === null ? null : (
        <div className="mt-4">
          <pre
            aria-live="polite"
            className={`${noticeClass} mt-0`}
            data-testid="workspace-grid-error"
          >
            {report}
          </pre>
          <button
            className={`${compactButtonClass} mt-2`}
            data-testid="workspace-grid-error-dismiss"
            onClick={() => setRecorded(dismissFailures(generation))}
            type="button"
          >
            Dismiss
          </button>
        </div>
      )}

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
