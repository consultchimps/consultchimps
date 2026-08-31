"use client";

/**
 * The interactive half of the /shortcuts page: a word box, a key capture area,
 * a palette of key buttons, and the grouped results.
 *
 * Two things make this more than a filtered list. The first is that the key
 * search is driven by real presses: the capture area listens for keydown and
 * keyup, so a visitor who half remembers a shortcut can hold what they do
 * remember and read the rest off the screen. The second is that the same
 * sequence can be built with the mouse, because a key search that only works
 * for people with a physical keyboard would leave out the visitors most likely
 * to be looking a shortcut up on a phone.
 *
 * Both inputs write to the same query, and everything about what a press means
 * to the list lives in `@/lib/shortcut-search`, which is tested without a
 * browser. This file owns events, focus, and rendering.
 *
 * The listening is deliberately scoped to the capture area. Key handling is
 * attached to that element rather than to the document, and preventDefault is
 * only called while it holds focus, so pressing Ctrl+P anywhere else on the
 * page still prints. Alt gets the same treatment on both keydown and keyup:
 * on Windows a bare Alt keyup is what opens the browser's menu bar, so the
 * default has to be suppressed on the release as well as the press. Tab is
 * the deliberate exception, because an area that swallows it is a trap for
 * anyone navigating by keyboard.
 */

import { EXCEL_SHORTCUTS } from "@/lib/excel-shortcuts-data";
import {
  keyTokenLabel,
  keyTokenName,
  type ExcelShortcut,
  type KeyToken,
} from "@/lib/excel-shortcuts";
import {
  appendKeyToken,
  beginKeyStep,
  EMPTY_KEY_QUERY,
  filterShortcuts,
  formatKeySequence,
  groupShortcutsByCategory,
  isEmptyKeyQuery,
  isMatchedKeyToken,
  keyTokenFromPress,
  nextKeyOptions,
  removeLastKeyEntry,
  type KeyQuery,
} from "@/lib/shortcut-search";
import { CornerDownRight, Delete, Keyboard, Search, X } from "lucide-react";
import {
  Fragment,
  useCallback,
  useId,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
} from "react";

const chipClass =
  "inline-flex items-center justify-center rounded-md border bg-fd-card px-2.5 py-1.5 font-mono text-xs font-semibold transition-colors hover:bg-fd-accent disabled:pointer-events-none disabled:opacity-40";
const actionButtonClass =
  "inline-flex items-center gap-1.5 rounded-md border bg-fd-card px-2.5 py-1.5 text-xs font-semibold transition-colors hover:bg-fd-accent disabled:pointer-events-none disabled:opacity-40";

/** One shortcut's keys, with the part the visitor entered picked out. */
function KeySequence({
  keys,
  query,
}: {
  readonly keys: ExcelShortcut["keys"];
  readonly query: KeyQuery;
}) {
  return (
    <span className="flex flex-wrap items-center gap-1">
      {keys.map((step, stepIndex) => (
        <Fragment key={stepIndex}>
          {stepIndex > 0 ? (
            <span className="px-0.5 font-mono text-[0.65rem] uppercase tracking-[0.1em] text-fd-muted-foreground">
              then
            </span>
          ) : null}
          {step.map((token, tokenIndex) => (
            <Fragment key={token}>
              {tokenIndex > 0 ? (
                <span className="text-xs text-fd-muted-foreground">+</span>
              ) : null}
              <kbd
                className="shortcut-key"
                data-matched={isMatchedKeyToken(query, stepIndex, token)}
              >
                {keyTokenLabel(token)}
              </kbd>
            </Fragment>
          ))}
        </Fragment>
      ))}
    </span>
  );
}

export function ShortcutFinder() {
  const [words, setWords] = useState("");
  const [query, setQuery] = useState<KeyQuery>(EMPTY_KEY_QUERY);
  /** Set once every key is released: the next key opens a new step. */
  const [stepClosed, setStepClosed] = useState(false);
  const [isListening, setIsListening] = useState(false);
  const heldCodes = useRef<Set<string>>(new Set());
  const wordsId = useId();
  const hintId = useId();

  const results = useMemo(
    () => filterShortcuts(EXCEL_SHORTCUTS, { keys: query, words }),
    [query, words],
  );
  const groups = useMemo(() => groupShortcutsByCategory(results), [results]);

  // Where the next key lands. Releasing the keys does not narrow the list on
  // its own, so the pending step only exists for the palette until a key or a
  // button actually opens it.
  const pendingQuery = useMemo(
    () => (stepClosed ? beginKeyStep(query) : query),
    [query, stepClosed],
  );
  const paletteOptions = useMemo(() => {
    const matches = stepClosed
      ? filterShortcuts(EXCEL_SHORTCUTS, { keys: pendingQuery, words })
      : results;
    return nextKeyOptions(matches, pendingQuery);
  }, [pendingQuery, results, stepClosed, words]);
  const canBeginNextStep = useMemo(
    () => !stepClosed && nextKeyOptions(results, query).canBeginNextStep,
    [query, results, stepClosed],
  );

  const addToken = useCallback(
    (token: KeyToken) => {
      setQuery((previous) =>
        appendKeyToken(
          // The pending step is materialised here, at the moment a key is
          // actually added to it.
          stepClosed ? beginKeyStep(previous) : previous,
          token,
        ),
      );
      setStepClosed(false);
    },
    [stepClosed],
  );

  const clearSequence = useCallback(() => {
    heldCodes.current.clear();
    setQuery(EMPTY_KEY_QUERY);
    setStepClosed(false);
  }, []);

  const removeLast = useCallback(() => {
    heldCodes.current.clear();
    setQuery((previous) => removeLastKeyEntry(previous));
    setStepClosed(false);
  }, []);

  const handleKeyDown = useCallback(
    (event: ReactKeyboardEvent<HTMLDivElement>) => {
      // Tab is the one key that keeps its default. Capturing it would leave a
      // visitor who reached this area by keyboard with no way out of it, which
      // is a worse failure than not being able to press Excel's Tab shortcuts
      // here: those are on the buttons below, as the hint says.
      if (event.key === "Tab") {
        return;
      }
      // Otherwise scoped to this element: the page's other keyboard
      // behaviour, and the browser's, are untouched everywhere else.
      event.preventDefault();
      event.stopPropagation();
      if (event.repeat) {
        return;
      }
      if (event.key === "Escape") {
        clearSequence();
        return;
      }
      if (event.key === "Backspace") {
        removeLast();
        return;
      }
      const token = keyTokenFromPress(event);
      if (token === null) {
        return;
      }
      heldCodes.current.add(event.code);
      addToken(token);
    },
    [addToken, clearSequence, removeLast],
  );

  const handleKeyUp = useCallback(
    (event: ReactKeyboardEvent<HTMLDivElement>) => {
      if (event.key === "Tab") {
        return;
      }
      // The release matters as much as the press: a bare Alt keyup is what
      // opens the browser menu bar on Windows.
      event.preventDefault();
      event.stopPropagation();
      // Only a key this area recorded as held can close the step. Backspace
      // and Escape are handled on the way down and never join that set, so
      // their release must not close the chord they just edited: Ctrl+Shift+L,
      // Backspace, O has to end as the chord Ctrl+Shift+O rather than as two
      // steps. The same guard covers a key pressed before the area had focus,
      // whose release arrives here on its own.
      if (!heldCodes.current.delete(event.code)) {
        return;
      }
      if (heldCodes.current.size === 0) {
        setStepClosed(true);
      }
    },
    [],
  );

  const handleBlur = useCallback(() => {
    // A key released while the focus was elsewhere never reaches the handler,
    // so leaving the area drops whatever it thought was held and treats the
    // step as finished. On an empty query that closing is a no-op.
    heldCodes.current.clear();
    setIsListening(false);
    setStepClosed(true);
  }, []);

  const hasSequence = !isEmptyKeyQuery(query);
  const countLabel =
    results.length === 1
      ? "1 shortcut matches"
      : `${results.length} shortcuts match`;

  return (
    <div className="flex flex-col gap-6">
      <div className="rounded-xl border bg-fd-card/80 p-5 shadow-[0_12px_36px_hsl(15_10%_11%/6%)] sm:p-6">
        <label className="block text-sm font-semibold" htmlFor={wordsId}>
          Search the descriptions
        </label>
        <div className="relative mt-2">
          <Search
            aria-hidden="true"
            className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-fd-muted-foreground"
          />
          <input
            autoComplete="off"
            className="w-full rounded-lg border bg-fd-background py-2 pl-9 pr-3 text-sm outline-none focus-visible:ring-2 focus-visible:ring-fd-ring"
            data-testid="word-search"
            id={wordsId}
            onChange={(event) => setWords(event.target.value)}
            placeholder="filter, freeze, paste values"
            type="text"
            value={words}
          />
        </div>
      </div>

      <div className="rounded-xl border bg-fd-card/80 p-5 shadow-[0_12px_36px_hsl(15_10%_11%/6%)] sm:p-6">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <span className="inline-flex items-center gap-2 text-sm font-semibold">
            <Keyboard aria-hidden="true" className="size-4 text-fd-primary" />
            Search by pressing the keys
          </span>
          <span
            className="font-mono text-[0.65rem] uppercase tracking-[0.14em] text-fd-muted-foreground"
            data-testid="capture-state"
          >
            {isListening ? "Listening" : "Click to listen"}
          </span>
        </div>

        <div
          aria-describedby={hintId}
          aria-label="Key sequence capture"
          className={`mt-3 min-h-20 rounded-lg border-2 border-dashed px-4 py-4 transition-colors ${
            isListening
              ? "border-fd-primary bg-fd-accent/30"
              : "bg-fd-background/60"
          }`}
          data-testid="key-capture"
          onBlur={handleBlur}
          onFocus={() => setIsListening(true)}
          onKeyDown={handleKeyDown}
          onKeyUp={handleKeyUp}
          role="group"
          tabIndex={0}
        >
          {hasSequence ? (
            <div
              className="flex flex-wrap items-center gap-1.5"
              data-testid="entered-sequence"
            >
              {query.steps.map((step, stepIndex) => (
                <Fragment key={stepIndex}>
                  {stepIndex > 0 ? (
                    <span className="px-0.5 font-mono text-[0.65rem] uppercase tracking-[0.1em] text-fd-muted-foreground">
                      then
                    </span>
                  ) : null}
                  {step.length === 0 ? (
                    <span className="font-mono text-xs text-fd-muted-foreground">
                      next key
                    </span>
                  ) : null}
                  {step.map((token, tokenIndex) => (
                    <Fragment key={token}>
                      {tokenIndex > 0 ? (
                        <span className="text-xs text-fd-muted-foreground">
                          +
                        </span>
                      ) : null}
                      <kbd className="shortcut-key" data-matched="true">
                        {keyTokenLabel(token)}
                      </kbd>
                    </Fragment>
                  ))}
                </Fragment>
              ))}
              {stepClosed ? (
                <span className="font-mono text-[0.65rem] uppercase tracking-[0.1em] text-fd-muted-foreground">
                  released
                </span>
              ) : null}
            </div>
          ) : (
            <p className="text-sm text-fd-muted-foreground">
              Click here, then press a shortcut. Hold the keys of a chord
              together, and press the keys of a ribbon route one after another
            </p>
          )}
        </div>

        <p
          className="mt-3 text-xs leading-5 text-fd-muted-foreground"
          id={hintId}
        >
          Backspace removes the last key, Escape clears the sequence, and Tab
          moves on to the buttons below, so search for those three keys with the
          buttons. A few combinations belong to the browser or to Windows and
          never reach this page, Ctrl+W and Alt+F4 among them
        </p>

        <div className="mt-4 flex flex-wrap gap-1.5" data-testid="key-palette">
          {paletteOptions.modifiers.map((token) => (
            <button
              aria-label={keyTokenName(token)}
              className={`${chipClass} border-fd-primary/40 text-fd-primary`}
              key={token}
              onClick={() => addToken(token)}
              type="button"
            >
              {keyTokenLabel(token)}
            </button>
          ))}
          {paletteOptions.keys.map((token) => (
            <button
              aria-label={keyTokenName(token)}
              className={chipClass}
              key={token}
              onClick={() => addToken(token)}
              type="button"
            >
              {keyTokenLabel(token)}
            </button>
          ))}
        </div>

        <div className="mt-4 flex flex-wrap items-center gap-2 border-t pt-4">
          <button
            className={actionButtonClass}
            data-testid="next-step"
            disabled={!canBeginNextStep}
            onClick={() => {
              setQuery((previous) => beginKeyStep(previous));
              setStepClosed(false);
            }}
            type="button"
          >
            <CornerDownRight aria-hidden="true" className="size-3.5" />
            Then
          </button>
          <button
            className={actionButtonClass}
            data-testid="remove-last"
            disabled={!hasSequence}
            onClick={removeLast}
            type="button"
          >
            <Delete aria-hidden="true" className="size-3.5" />
            Remove last
          </button>
          <button
            className={actionButtonClass}
            data-testid="clear-sequence"
            disabled={!hasSequence}
            onClick={clearSequence}
            type="button"
          >
            <X aria-hidden="true" className="size-3.5" />
            Clear
          </button>
          <span className="ml-auto font-mono text-xs text-fd-muted-foreground">
            {hasSequence
              ? formatKeySequence(query.steps.filter((step) => step.length > 0))
              : "No keys entered"}
          </span>
        </div>
      </div>

      <p
        className="font-mono text-xs font-semibold uppercase tracking-[0.16em] text-fd-muted-foreground"
        data-testid="result-count"
        role="status"
      >
        {countLabel}
      </p>

      {groups.length === 0 ? (
        <div
          className="rounded-xl border border-dashed px-5 py-8 text-center"
          data-testid="empty-state"
        >
          <p className="text-sm font-semibold">No shortcut matches</p>
          <p className="mt-1 text-sm text-fd-muted-foreground">
            Remove the last key, clear the sequence, or try another word
          </p>
        </div>
      ) : (
        <div className="flex flex-col gap-8">
          {groups.map(({ category, shortcuts }) => (
            <section
              data-testid={`group-${category.replaceAll(" ", "-")}`}
              key={category}
            >
              <h2 className="font-mono text-xs font-semibold uppercase tracking-[0.16em] text-fd-primary">
                {category} · {shortcuts.length}
              </h2>
              <ul className="mt-3 divide-y rounded-xl border bg-fd-card/60">
                {shortcuts.map((shortcut) => (
                  <li
                    className="flex flex-col gap-2 px-4 py-3 sm:flex-row sm:items-baseline sm:gap-5"
                    data-testid="shortcut-row"
                    key={shortcut.id}
                  >
                    <div className="sm:w-64 sm:shrink-0">
                      <KeySequence keys={shortcut.keys} query={query} />
                    </div>
                    <div className="min-w-0">
                      <p className="text-sm leading-6">{shortcut.action}</p>
                      {shortcut.note === undefined ? null : (
                        <p className="mt-0.5 text-xs leading-5 text-fd-muted-foreground">
                          {shortcut.note}
                        </p>
                      )}
                    </div>
                  </li>
                ))}
              </ul>
            </section>
          ))}
        </div>
      )}
    </div>
  );
}
