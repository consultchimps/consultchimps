"use client";

/**
 * Bringing a spreadsheet into an open workspace.
 *
 * The flow is two steps on purpose. Choosing a file only asks the worker what
 * the file holds, so nothing is written while the visitor is still deciding;
 * the answer becomes a small form, one row per worksheet, carrying a suggested
 * table name and Record ID prefix that are visible and editable before
 * anything is created. Import then sends the file back with the names that were
 * settled on, and `@consultchimps/db` creates every chosen table in one unit.
 *
 * The component owns the whole interaction and reports only the new workspace
 * summary upward, which keeps the page's mount point to a single line and the
 * two files independent of each other.
 */

import {
  describeFailure,
  inputClass,
  primaryButtonClass,
  readUploads,
  secondaryButtonClass,
  sectionClass,
} from "@/components/tool-kit";
import {
  WORKSPACE_IMPORT_FILES,
  workspaceImportKind,
  type WorkspaceImportKind,
} from "@/lib/accepted-files";
import {
  cellCountText,
  importBlockers,
  isImportable,
  type ImportBlocker,
} from "@/lib/workspace-import-blockers";
import type {
  ImportSourceDescription,
  ImportTableChoice,
  WorkspaceSummary,
} from "@/lib/workspace-protocol";
import type { WorkspaceBusy } from "@/lib/workspace-state";
import type { WorkspaceClient } from "@/lib/workspace-worker";
import { identifierKey, MAX_RECORD_ID_PADDING } from "@consultchimps/db/schema";
import { FileUp, LoaderCircle, Upload, X } from "lucide-react";
import { useCallback, useRef, useState } from "react";

/** The default zero-padding a new table numbers its records with. */
const DEFAULT_PADDING = "4";

/** The file the visitor chose, held so import can send the same bytes back. */
interface ChosenFile {
  readonly name: string;
  /**
   * Which family it belongs to, settled here where the browser's media type is
   * still available, and carried from here on. Nothing downstream sees the
   * `File`, so nothing downstream can decide this again.
   */
  readonly kind: WorkspaceImportKind;
  readonly bytes: Uint8Array;
  readonly sources: readonly ImportSourceDescription[];
}

/** One row of the form: a source, and what the visitor wants to call it. */
interface SourceChoice {
  selected: boolean;
  tableName: string;
  recordIdPrefix: string;
  /** Held as text so the field can be emptied while it is being retyped. */
  recordIdPadding: string;
}

/**
 * What to say about a condition that stops this source being imported.
 *
 * The wording is the page's own - the visitor is looking at the row already, so
 * nothing here repeats the file or worksheet name - while which conditions
 * exist, and in which order, comes from the rule the worker refuses by. The
 * form therefore never offers a tick the import will not honour.
 *
 * `describeImportSources` lists a source only when it has a table or has one of
 * these conditions, and the worker refuses a source with no table regardless,
 * so a change to that listing shows up as a refusal rather than as a table of
 * values nobody entered.
 */
function blockerText(blocker: ImportBlocker): string {
  const held = cellCountText(blocker.cells);
  const hold = blocker.cells === 1 ? "holds" : "hold";
  return blocker.kind === "uncalculated-formulas"
    ? `${held} ${hold} a formula this workbook carries no calculated value for, so importing would leave those values empty. Open the workbook in Excel, let it calculate, save it, and choose the file again`
    : `${held} ${hold} an error value, which this workbook stores as a number, so importing would put numbers nobody entered in those cells. Fix or clear the errors in Excel, save the workbook, and choose the file again`;
}

function initialChoices(
  sources: readonly ImportSourceDescription[],
): SourceChoice[] {
  return sources.map((source) => ({
    // A source with no rows can still be worth creating as an empty table, so
    // it is offered; everything that carries data and can be imported is ticked
    // to begin with.
    selected: isImportable(source) && source.rowCount > 0,
    tableName: source.suggestedTableName,
    recordIdPrefix: source.suggestedRecordIdPrefix,
    recordIdPadding: DEFAULT_PADDING,
  }));
}

/**
 * What is wrong with the form, as one message, or null when it is ready. Table
 * names are compared with the database package's own identifier fold, so the
 * page refuses exactly the names the workspace would refuse rather than a
 * near-enough approximation of them.
 */
function formProblem(
  choices: readonly SourceChoice[],
  existingTableNames: readonly string[],
): string | null {
  const chosen = choices.filter((choice) => choice.selected);
  if (chosen.length === 0) {
    return "Choose at least one table to import";
  }
  const taken = new Map(
    existingTableNames.map((name) => [identifierKey(name), name]),
  );
  for (const choice of chosen) {
    const name = choice.tableName.trim();
    if (name === "") {
      return "Give every table you are importing a name";
    }
    if (choice.recordIdPrefix.trim() === "") {
      return `Give "${name}" a Record ID prefix`;
    }
    // The text is judged, not just the number it parses to: "1e3" is a whole
    // number to `Number` and a refusal to `assertRecordIdConfig`, and a page
    // that enables Import on a value the library will reject has checked the
    // wrong thing.
    const padding = Number(choice.recordIdPadding);
    if (
      !/^\d+$/u.test(choice.recordIdPadding.trim()) ||
      padding > MAX_RECORD_ID_PADDING
    ) {
      return `The padding for "${name}" must be a whole number from 0 to ${MAX_RECORD_ID_PADDING}`;
    }
    const key = identifierKey(name);
    const clash = taken.get(key);
    if (clash !== undefined) {
      return clash === name
        ? `A table called "${name}" would be created twice, so give one of them another name`
        : `"${name}" cannot be used, because the workspace already reads "${clash}" as the same name`;
    }
    taken.set(key, name);
  }
  return null;
}

export interface WorkspaceImportProps {
  /**
   * The page's one busy state, read rather than kept here. Reading a file and
   * running an import are page-wide commands like a save, so they belong in the
   * shell's busy state: while either runs, New, Open, and Save are held back
   * with the rest of the page, and a click on New cannot queue itself behind an
   * import and discard what the import created.
   */
  readonly busy: WorkspaceBusy;
  /** The worker client, created lazily by the page that owns it. */
  readonly client: () => WorkspaceClient;
  /** The tables the workspace already holds, so a clash is caught early. */
  readonly existingTableNames: readonly string[];
  /**
   * A file was chosen and is being read so its sources can be described.
   *
   * This and the four below report what this section did, as it happens, rather
   * than setting a busy value from here. The page's state model decides what
   * each of them means; this section only says which happened, so one place
   * knows what the page is doing.
   */
  readonly onReading: () => void;
  /** That read ended, whether it described the file or refused it. */
  readonly onReadFinished: () => void;
  /** The import itself has been sent to the worker. */
  readonly onRunning: () => void;
  /** It landed: the workspace as it stands now, and what to say about it. */
  readonly onImported: (summary: WorkspaceSummary, notice: string) => void;
  /**
   * It did not. The explanation stays here, against the form it belongs to, so
   * the page says nothing about a failure whose context is on this section.
   */
  readonly onFailed: () => void;
}

export function WorkspaceImport({
  busy,
  client,
  existingTableNames,
  onFailed,
  onImported,
  onReadFinished,
  onReading,
  onRunning,
}: WorkspaceImportProps) {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const [file, setFile] = useState<ChosenFile | null>(null);
  const [choices, setChoices] = useState<SourceChoice[]>([]);
  const [error, setError] = useState<string | null>(null);

  const reset = useCallback(() => {
    setFile(null);
    setChoices([]);
    setError(null);
  }, []);

  const onFiles = useCallback(
    (files: readonly File[]) => {
      void (async () => {
        onReading();
        setError(null);
        try {
          // The kind and the acceptance are one question, asked once. Picking
          // the file by its kind is what makes them the same question.
          let chosen: { file: File; kind: WorkspaceImportKind } | undefined;
          for (const candidate of files) {
            const kind = workspaceImportKind(candidate);
            if (kind !== undefined) {
              chosen = { file: candidate, kind };
              break;
            }
          }
          if (chosen === undefined) {
            throw new Error(
              `That file is not ${WORKSPACE_IMPORT_FILES.description}, so nothing was read`,
            );
          }
          const [first] = await readUploads([chosen.file], () => true);
          if (first === undefined) {
            throw new Error(
              `That file is not ${WORKSPACE_IMPORT_FILES.description}, so nothing was read`,
            );
          }
          const sources = await client().describeImport(
            first.name,
            chosen.kind,
            first.bytes,
          );
          setFile({
            name: first.name,
            kind: chosen.kind,
            bytes: first.bytes,
            sources,
          });
          setChoices(initialChoices(sources));
        } catch (caught) {
          setFile(null);
          setChoices([]);
          setError(describeFailure(caught));
        } finally {
          onReadFinished();
        }
      })();
    },
    [client, onReadFinished, onReading],
  );

  const update = useCallback((index: number, change: Partial<SourceChoice>) => {
    setChoices((previous) =>
      previous.map((choice, position) =>
        position === index ? { ...choice, ...change } : choice,
      ),
    );
  }, []);

  const onImport = useCallback(() => {
    void (async () => {
      if (file === null) {
        return;
      }
      const tables: ImportTableChoice[] = [];
      file.sources.forEach((source, index) => {
        const choice = choices[index];
        if (choice === undefined || !choice.selected) {
          return;
        }
        tables.push({
          source: source.name,
          tableName: choice.tableName.trim(),
          recordIdPrefix: choice.recordIdPrefix.trim(),
          recordIdPadding: Number(choice.recordIdPadding),
        });
      });

      onRunning();
      setError(null);
      try {
        const result = await client().importFile(
          file.name,
          file.kind,
          file.bytes,
          tables,
        );
        const rows = result.tables.reduce(
          (total, table) => total + table.rowCount,
          0,
        );
        // A source that already carried a Record ID column had it left out,
        // because identifiers are generated. Saying so is the difference
        // between a report and a claim: the alternative is a notice that counts
        // the rows and never mentions the column that did not arrive.
        const ignored = [
          ...new Set(
            result.tables.flatMap((table) => [...table.ignoredColumns]),
          ),
        ];
        // A header too long to be a column name, or one that collided with
        // another once it had been shortened, is stored under a name the file
        // did not write. The values are all there, so this is a note rather
        // than a warning, but it is not something to leave unsaid.
        const renamed = result.tables.flatMap((table) => [
          ...table.renamedColumns,
        ]);
        onImported(
          result.summary,
          `Imported ${result.tables.length === 1 ? "1 table" : `${result.tables.length} tables`} with ${rows === 1 ? "1 row" : `${rows} rows`}${
            ignored.length === 0
              ? ""
              : `. The ${ignored.join(" and ")} column${ignored.length === 1 ? " was" : "s were"} left out, because a Record ID is always generated`
          }${
            renamed.length === 0
              ? ""
              : `. ${renamed.length === 1 ? "1 column was" : `${renamed.length} columns were`} stored under a shorter name, because the header was longer than a column name can be`
          }`,
        );
        reset();
      } catch (caught) {
        // Reported after the message is on screen, and only on this path: the
        // import that landed has already told the page what it did, and there
        // is no third outcome for a `finally` to cover.
        setError(describeFailure(caught));
        onFailed();
      }
    })();
  }, [choices, client, file, onFailed, onImported, onRunning, reset]);

  const problem =
    file === null ? null : formProblem(choices, existingTableNames);
  const isBusy = busy !== null;

  return (
    <section className={sectionClass} data-testid="workspace-import">
      <h2 className="text-xl font-bold tracking-[-0.03em]">Import data</h2>
      <p className="mt-3 text-sm text-fd-muted-foreground">
        Add the rows of a worksheet or a .csv file to this workspace as a table.
        Every row is given a Record ID, and each column takes the type its own
        values agree on, or text when they do not. A workbook&apos;s hidden
        worksheets are not offered
      </p>

      <div className="mt-5 flex flex-wrap gap-3">
        <button
          className={secondaryButtonClass}
          data-testid="workspace-import-choose"
          disabled={isBusy}
          onClick={() => inputRef.current?.click()}
          type="button"
        >
          {busy === "reading" ? (
            <LoaderCircle aria-hidden="true" className="size-4 animate-spin" />
          ) : (
            <FileUp aria-hidden="true" className="size-4" />
          )}
          Choose a file to import
        </button>
        <input
          accept={WORKSPACE_IMPORT_FILES.accept}
          aria-label="Choose a file to import"
          className="hidden"
          data-testid="workspace-import-input"
          onChange={(event) => {
            const files = [...(event.target.files ?? [])];
            event.target.value = "";
            onFiles(files);
          }}
          ref={inputRef}
          type="file"
        />
      </div>

      {file !== null ? (
        <div className="mt-6" data-testid="workspace-import-form">
          <p className="text-sm">
            <span className="font-mono">{file.name}</span> holds{" "}
            {file.sources.length === 1
              ? "one table"
              : `${file.sources.length} worksheets`}
            . Choose what to import and what to call it
          </p>

          <ul className="mt-4 space-y-4">
            {file.sources.map((source, index) => {
              const choice = choices[index];
              if (choice === undefined) {
                return null;
              }
              return (
                <li
                  className="rounded-lg border bg-fd-background/60 p-4"
                  data-testid="workspace-import-source"
                  key={source.name}
                >
                  <label className="flex items-center gap-2 text-sm font-semibold">
                    <input
                      checked={choice.selected}
                      data-testid="workspace-import-selected"
                      disabled={isBusy || !isImportable(source)}
                      onChange={(event) =>
                        update(index, { selected: event.target.checked })
                      }
                      type="checkbox"
                    />
                    <span data-testid="workspace-import-source-name">
                      {source.name}
                    </span>
                    <span className="font-normal text-fd-muted-foreground">
                      {source.rowCount === 1
                        ? "1 row"
                        : `${source.rowCount} rows`}
                      ,{" "}
                      {source.columnCount === 1
                        ? "1 column"
                        : `${source.columnCount} columns`}
                    </span>
                  </label>

                  {importBlockers(source).map((blocker) => (
                    <p
                      className="mt-2 text-sm text-fd-muted-foreground"
                      data-testid="workspace-import-blocked"
                      key={blocker.kind}
                    >
                      {blockerText(blocker)}
                    </p>
                  ))}

                  <div className="mt-3 grid gap-3 sm:grid-cols-[2fr_1fr_1fr]">
                    <label className="block text-xs font-semibold uppercase tracking-[0.12em] text-fd-muted-foreground">
                      Table name
                      <input
                        className={`mt-1 ${inputClass} font-normal normal-case tracking-normal`}
                        data-testid="workspace-import-name"
                        disabled={
                          isBusy || !choice.selected || !isImportable(source)
                        }
                        onChange={(event) =>
                          update(index, { tableName: event.target.value })
                        }
                        type="text"
                        value={choice.tableName}
                      />
                    </label>
                    <label className="block text-xs font-semibold uppercase tracking-[0.12em] text-fd-muted-foreground">
                      Record ID prefix
                      <input
                        className={`mt-1 ${inputClass} font-normal normal-case tracking-normal`}
                        data-testid="workspace-import-prefix"
                        disabled={
                          isBusy || !choice.selected || !isImportable(source)
                        }
                        onChange={(event) =>
                          update(index, { recordIdPrefix: event.target.value })
                        }
                        type="text"
                        value={choice.recordIdPrefix}
                      />
                    </label>
                    <label className="block text-xs font-semibold uppercase tracking-[0.12em] text-fd-muted-foreground">
                      Padding
                      <input
                        className={`mt-1 ${inputClass} font-normal normal-case tracking-normal`}
                        data-testid="workspace-import-padding"
                        disabled={
                          isBusy || !choice.selected || !isImportable(source)
                        }
                        max={MAX_RECORD_ID_PADDING}
                        min={0}
                        onChange={(event) =>
                          update(index, { recordIdPadding: event.target.value })
                        }
                        type="number"
                        value={choice.recordIdPadding}
                      />
                    </label>
                  </div>
                </li>
              );
            })}
          </ul>

          {problem === null ? null : (
            <p
              className="mt-4 text-sm text-fd-muted-foreground"
              data-testid="workspace-import-problem"
            >
              {problem}
            </p>
          )}

          <div className="mt-5 flex flex-wrap gap-3">
            <button
              className={primaryButtonClass}
              data-testid="workspace-import-run"
              disabled={isBusy || problem !== null}
              onClick={onImport}
              type="button"
            >
              {busy === "importing" ? (
                <LoaderCircle
                  aria-hidden="true"
                  className="size-4 animate-spin"
                />
              ) : (
                <Upload aria-hidden="true" className="size-4" />
              )}
              Import
            </button>
            <button
              className={secondaryButtonClass}
              data-testid="workspace-import-cancel"
              disabled={isBusy}
              onClick={reset}
              type="button"
            >
              <X aria-hidden="true" className="size-4" />
              Cancel
            </button>
          </div>
        </div>
      ) : null}

      {error ? (
        <pre
          aria-live="polite"
          className="mt-4 overflow-x-auto whitespace-pre-wrap rounded-lg border border-fd-primary/40 bg-fd-accent/30 px-4 py-3 text-xs leading-6 text-fd-accent-foreground"
          data-testid="workspace-import-error"
        >
          {error}
        </pre>
      ) : null}
    </section>
  );
}
