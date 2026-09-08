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
import type { WorkspaceColumn, WorkspaceTable } from "@/lib/workspace-protocol";
import type { WorkspaceClient } from "@/lib/workspace-worker";
import { RECORD_ID_COLUMN } from "@consultchimps/db";
import type { CellValue, TableRow } from "@consultchimps/tabular";
import { Table2 } from "lucide-react";
import { useEffect, useId, useRef, useState } from "react";
import type {
  CellComponent,
  ColumnDefinition,
  RowComponent,
  Tabulator,
} from "tabulator-tables";

/** Render text as a node rather than markup, so no stored value can be HTML. */
function textElement(text: string): HTMLElement {
  const element = document.createElement("span");
  element.textContent = text;
  return element;
}

/** How a value reads in a cell when nothing prettier applies. */
function asText(value: CellValue | undefined): string {
  return value === null || value === undefined ? "" : String(value);
}

/**
 * The editor a column gets. Each is Tabulator's own: the point is to hand the
 * visitor an input that suits the column, not to police what they put in it.
 */
function gridColumn(column: WorkspaceColumn): ColumnDefinition {
  const base: ColumnDefinition = {
    title: column.name,
    field: column.name,
    minWidth: 120,
  };

  // A foreign key stores the referenced Record ID and shows a readable label.
  // The list editor searches the labels and writes the id behind them.
  if (column.references !== null) {
    const labels = new Map(
      column.references.map((reference) => [reference.value, reference.label]),
    );
    return {
      ...base,
      editor: "list",
      editorParams: {
        values: column.references.map((reference) => ({
          label: reference.label,
          value: reference.value,
        })),
        autocomplete: true,
        listOnEmpty: true,
        clearable: column.nullable,
        // Past the option cap the list cannot name every record, so a Record ID
        // can still be typed. Whether it exists is the database's answer, not
        // this editor's.
        freetext: column.referencesTruncated,
        placeholderEmpty: "No matching record",
      },
      formatter: (cell: CellComponent) => {
        const value = cell.getValue() as CellValue | undefined;
        const text = asText(value);
        return textElement(labels.get(text) ?? text);
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

function gridColumns(table: WorkspaceTable): ColumnDefinition[] {
  return [
    {
      title: "Record ID",
      field: RECORD_ID_COLUMN,
      headerTooltip:
        "Assigned once when the record is created, and never changes",
      width: 150,
    },
    ...table.columns.map(gridColumn),
  ];
}

export interface WorkspaceGridProps {
  /** The client that owns the workspace database. */
  /**
   * Hands back the client that owns the workspace database. It is a function
   * rather than the client itself because the page holds it in a ref, and a ref
   * is only safe to read where this component reads it: inside an effect.
   */
  readonly getClient: () => WorkspaceClient;
  /** Where a refusal is shown. Called with null once an edit succeeds. */
  readonly onError: (message: string | null) => void;
}

/** What the last completed read produced, and which table it was for. */
interface LoadedTable {
  readonly table: string;
  /** The table, or null when the read was refused. */
  readonly data: WorkspaceTable | null;
}

/**
 * Mount this with a key that changes whenever a workspace is created or opened.
 * Every piece of state here belongs to one database, so a new one gets a new
 * component rather than a reset path that has to remember each field.
 */
export function WorkspaceGrid({ getClient, onError }: WorkspaceGridProps) {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const selectId = useId();

  const [tables, setTables] = useState<readonly string[] | null>(null);
  const [selected, setSelected] = useState<string | null>(null);
  const [loaded, setLoaded] = useState<LoadedTable | null>(null);

  // Which tables the workspace holds. Read once: this component belongs to one
  // database, and no path in this page adds or drops a table.
  useEffect(() => {
    let cancelled = false;
    void getClient()
      .listTables()
      .then((names) => {
        if (cancelled) {
          return;
        }
        setTables(names);
        setSelected(names[0] ?? null);
      })
      .catch((caught: unknown) => {
        if (cancelled) {
          return;
        }
        setTables([]);
        onError(describeFailure(caught));
      });
    return () => {
      cancelled = true;
    };
  }, [getClient, onError]);

  // The selected table's columns and rows. The reply records which table it was
  // for, so one that arrives after the visitor switched away is recognised as
  // stale below rather than rendered under the wrong heading.
  useEffect(() => {
    if (selected === null) {
      return;
    }
    let cancelled = false;
    void getClient()
      .readTable(selected)
      .then((table) => {
        if (!cancelled) {
          setLoaded({ table: selected, data: table });
        }
      })
      .catch((caught: unknown) => {
        if (cancelled) {
          return;
        }
        setLoaded({ table: selected, data: null });
        onError(describeFailure(caught));
      });
    return () => {
      cancelled = true;
    };
  }, [getClient, onError, selected]);

  // Derived rather than stored, so switching tables needs no state reset: the
  // grid shows a table only while it is the selected one, and is reading until
  // an answer for that table has come back.
  const showing = loaded !== null && loaded.table === selected;
  const data = showing ? loaded.data : null;
  const loading = tables === null || (selected !== null && !showing);

  // Build the grid for the table now loaded, and tear it down when the table
  // changes or the component unmounts.
  useEffect(() => {
    const element = containerRef.current;
    if (data === null || element === null) {
      return;
    }

    const workspace = getClient();
    let instance: Tabulator | null = null;
    let destroyed = false;

    // The last value the database confirmed for each record, which is what a
    // refused edit reverts to. Reverting to the cell's previous value would be
    // wrong the moment two edits to one cell overlap: the second one's
    // "previous" is the first one's unstored text.
    const committed = new Map<string, TableRow>(
      data.rows.map((row) => [asText(row[RECORD_ID_COLUMN]), { ...row }]),
    );
    // Edits are numbered per cell so a reply that has been overtaken by a newer
    // edit to the same cell reports its outcome without touching the grid.
    const latest = new Map<string, number>();
    let sequence = 0;
    // Putting a value back is itself a cell change, so the edit handler ignores
    // anything written while this is raised.
    let applying = 0;

    const persist = (cell: CellComponent): void => {
      if (applying > 0) {
        return;
      }
      const recordId = String(cell.getRow().getIndex());
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

      void workspace
        .updateCell({ table: data.name, recordId, column, value })
        .then((stored) => {
          const row = committed.get(recordId);
          if (row !== undefined) {
            row[column] = stored;
          }
          onError(null);
        })
        .catch((caught: unknown) => {
          onError(describeFailure(caught));
        })
        .finally(() => {
          if (destroyed || latest.get(key) !== mine) {
            return;
          }
          latest.delete(key);
          // One path for both outcomes: the cell is set to what the database
          // holds, which a success has just updated and a refusal has not. A
          // record with no baseline is left alone rather than blanked, so a
          // missing entry could never turn a refusal into data loss.
          const row = committed.get(recordId);
          if (row === undefined) {
            return;
          }
          applying += 1;
          try {
            cell.setValue(row[column] ?? null);
          } finally {
            applying -= 1;
          }
        });
    };

    void (async () => {
      const { TabulatorFull } = await import("tabulator-tables");
      if (destroyed) {
        return;
      }
      instance = new TabulatorFull(element, {
        // Copies, because Tabulator writes into the row objects it is given and
        // these are the baseline a refused edit reverts to.
        data: data.rows.map((row) => ({ ...row })),
        columns: gridColumns(data),
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
    })();

    return () => {
      destroyed = true;
      instance?.destroy();
    };
  }, [data, getClient, onError]);

  const truncated =
    data?.columns.filter((column) => column.referencesTruncated) ?? [];

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
        {tables !== null && tables.length > 0 ? (
          <label className="flex items-center gap-2 text-sm" htmlFor={selectId}>
            Table
            <select
              className="rounded-lg border bg-fd-card px-3 py-2 text-sm outline-none focus-visible:ring-2 focus-visible:ring-fd-ring"
              data-testid="workspace-table-select"
              id={selectId}
              onChange={(event) => setSelected(event.target.value)}
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

      {tables !== null && tables.length === 0 ? (
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
