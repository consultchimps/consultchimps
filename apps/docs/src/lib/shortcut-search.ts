/**
 * The matching engine behind the shortcut page: a key sequence being built up,
 * a word filter, and the prefix rule that decides which shortcuts survive.
 *
 * The engine is a set of pure functions over plain data so it can be tested
 * without a browser. The page owns the keyboard and pointer events; everything
 * about what a press means to the list lives here.
 *
 * The prefix rule, stated once:
 *
 * - A query is a list of steps, the same shape a shortcut's keys have.
 * - Every step before the last must match the shortcut's step at that index
 *   exactly, as a set: the order the modifiers were pressed in does not
 *   matter, but a step the visitor has moved on from is finished.
 * - The last step is a partial chord, so it only has to be a subset of the
 *   shortcut's step at that index. Holding Shift therefore keeps Ctrl+Shift+L,
 *   and adding Alt to the same chord keeps Ctrl+Shift+Alt+... instead.
 * - A shortcut with fewer steps than the query cannot match, so committing a
 *   second step drops the chords and leaves the ribbon routes.
 */

import {
  isKeyToken,
  isModifierKeyToken,
  MODIFIER_KEY_TOKENS,
  SHORTCUT_CATEGORIES,
  type ExcelShortcut,
  type KeyToken,
  type ShortcutCategory,
  type ShortcutKeys,
  type ShortcutStep,
} from "./excel-shortcuts";

/**
 * Punctuation by physical key, so a shifted key reports the cap rather than
 * the character it produces: the database stores Ctrl+Shift+` and the visitor
 * pressing it produces Ctrl+Shift+`, not Ctrl+Shift+~.
 */
const PUNCTUATION_BY_CODE: Readonly<Record<string, KeyToken>> = {
  Backquote: "`",
  Backslash: "\\",
  BracketLeft: "[",
  BracketRight: "]",
  Comma: ",",
  Equal: "=",
  Minus: "-",
  Period: ".",
  Quote: "'",
  Semicolon: ";",
  Slash: "/",
};

/** The part of a keyboard event this mapping reads. */
export interface KeyPress {
  /** The physical key, such as KeyL or Digit8. */
  readonly code: string;
  /** The character or key name the browser reports, such as l or ArrowUp. */
  readonly key: string;
}

/**
 * The canonical token for a press, or null for a key the database does not
 * model. Letters and digits come from the physical key so that Shift+3 is
 * recorded as Shift+3 rather than as the # it types.
 */
export function keyTokenFromPress(press: KeyPress): KeyToken | null {
  const { code, key } = press;
  switch (key) {
    case "Control":
      return "Ctrl";
    case "Shift":
      return "Shift";
    case "Alt":
      return "Alt";
    case "Meta":
    case "OS":
      return "Win";
    case " ":
      return "Space";
    default:
      break;
  }

  const letter = /^Key([A-Z])$/u.exec(code)?.[1];
  if (letter !== undefined && isKeyToken(letter)) {
    return letter;
  }
  const digit = /^Digit([0-9])$/u.exec(code)?.[1];
  if (digit !== undefined && isKeyToken(digit)) {
    return digit;
  }
  const punctuation = PUNCTUATION_BY_CODE[code];
  if (punctuation !== undefined) {
    return punctuation;
  }
  // The function keys and the named keys report a usable `key` directly.
  if (isKeyToken(key)) {
    return key;
  }
  const upper = key.toUpperCase();
  return isKeyToken(upper) ? upper : null;
}

/**
 * The sequence entered so far. The last step is the one still being built,
 * whether the visitor is holding its keys or has released them; a released
 * step only stops accepting keys, which `beginKeyStep` records by opening an
 * empty step after it.
 */
export interface KeyQuery {
  readonly steps: readonly ShortcutStep[];
}

export const EMPTY_KEY_QUERY: KeyQuery = { steps: [] };

function lastStepOf(query: KeyQuery): ShortcutStep {
  return query.steps.at(-1) ?? [];
}

/** True when nothing has been entered, so the query filters nothing out. */
export function isEmptyKeyQuery(query: KeyQuery): boolean {
  return query.steps.every((step) => step.length === 0);
}

export function keyQueryTokenCount(query: KeyQuery): number {
  return query.steps.reduce((total, step) => total + step.length, 0);
}

/**
 * Add a key to the step being built. A key already in that step is ignored,
 * which is what the browser's auto-repeat produces while a key is held.
 */
export function appendKeyToken(query: KeyQuery, token: KeyToken): KeyQuery {
  if (query.steps.length === 0) {
    return { steps: [[token]] };
  }
  const last = lastStepOf(query);
  if (last.includes(token)) {
    return query;
  }
  return { steps: [...query.steps.slice(0, -1), [...last, token]] };
}

/**
 * Finish the current step and open the next one. Called when the visitor
 * releases every key, and by the "Then" button for pointer input.
 */
export function beginKeyStep(query: KeyQuery): KeyQuery {
  if (lastStepOf(query).length === 0) {
    return query;
  }
  return { steps: [...query.steps, []] };
}

/**
 * Undo one entry: the last key of the step being built, or the step itself
 * once it is empty. Removing the last key of the only step clears the query.
 */
export function removeLastKeyEntry(query: KeyQuery): KeyQuery {
  if (query.steps.length === 0) {
    return query;
  }
  const last = lastStepOf(query);
  if (last.length === 0) {
    return { steps: query.steps.slice(0, -1) };
  }
  const shortened = last.slice(0, -1);
  if (shortened.length === 0 && query.steps.length === 1) {
    return EMPTY_KEY_QUERY;
  }
  return { steps: [...query.steps.slice(0, -1), shortened] };
}

function isSameTokenSet(left: ShortcutStep, right: ShortcutStep): boolean {
  return (
    left.length === right.length && left.every((token) => right.includes(token))
  );
}

function isTokenSubset(subset: ShortcutStep, step: ShortcutStep): boolean {
  return subset.every((token) => step.includes(token));
}

/** Whether a shortcut's key sequence starts with the query. */
export function matchesKeyQuery(keys: ShortcutKeys, query: KeyQuery): boolean {
  if (query.steps.length > keys.length) {
    return false;
  }
  const lastIndex = query.steps.length - 1;
  return query.steps.every((queryStep, index) => {
    const shortcutStep = keys[index] ?? [];
    return index === lastIndex
      ? isTokenSubset(queryStep, shortcutStep)
      : isSameTokenSet(queryStep, shortcutStep);
  });
}

/**
 * Whether one of a shortcut's keys was part of the match, so the page can
 * highlight the prefix the visitor entered. A token counts as matched when the
 * query names it at the same step, which is exactly the rule above: earlier
 * steps match as a whole, and the last step matches the keys it holds.
 */
export function isMatchedKeyToken(
  query: KeyQuery,
  stepIndex: number,
  token: KeyToken,
): boolean {
  return query.steps[stepIndex]?.includes(token) ?? false;
}

/** Case-insensitive substring search over the action, category, and note. */
export function matchesWordQuery(
  shortcut: ExcelShortcut,
  words: string,
): boolean {
  const needle = words.trim().toLowerCase();
  if (needle === "") {
    return true;
  }
  const haystack =
    `${shortcut.action} ${shortcut.category} ${shortcut.note ?? ""}`.toLowerCase();
  return haystack.includes(needle);
}

export interface ShortcutFilter {
  readonly keys: KeyQuery;
  readonly words: string;
}

/** The word filter and the key filter compose: a shortcut must pass both. */
export function filterShortcuts(
  shortcuts: readonly ExcelShortcut[],
  filter: ShortcutFilter,
): readonly ExcelShortcut[] {
  return shortcuts.filter(
    (shortcut) =>
      matchesWordQuery(shortcut, filter.words) &&
      matchesKeyQuery(shortcut.keys, filter.keys),
  );
}

export interface ShortcutGroup {
  readonly category: ShortcutCategory;
  readonly shortcuts: readonly ExcelShortcut[];
}

/**
 * Group results by category, in the declared category order, dropping the
 * categories nothing matched.
 */
export function groupShortcutsByCategory(
  shortcuts: readonly ExcelShortcut[],
): readonly ShortcutGroup[] {
  return SHORTCUT_CATEGORIES.map((category) => ({
    category,
    shortcuts: shortcuts.filter((shortcut) => shortcut.category === category),
  })).filter((group) => group.shortcuts.length > 0);
}

export interface NextKeyOptions {
  /** Modifiers that would keep at least one of the current results. */
  readonly modifiers: readonly KeyToken[];
  /** Other keys that would keep at least one of the current results. */
  readonly keys: readonly KeyToken[];
  /** Whether any result continues into a further step from here. */
  readonly canBeginNextStep: boolean;
}

// Modifiers sort ahead of the rest, in the order Excel writes them.
const modifierRank = new Map<string, number>(
  MODIFIER_KEY_TOKENS.map((token, index) => [token, index]),
);

function compareTokens(left: KeyToken, right: KeyToken): number {
  const leftRank = modifierRank.get(left) ?? Number.MAX_SAFE_INTEGER;
  const rightRank = modifierRank.get(right) ?? Number.MAX_SAFE_INTEGER;
  if (leftRank !== rightRank) {
    return leftRank - rightRank;
  }
  return left.localeCompare(right, "en");
}

/**
 * The keys worth offering next, taken from the shortcuts that still match.
 * Offering only these means a chip can never empty the results, and the
 * palette narrows as the sequence grows.
 */
export function nextKeyOptions(
  matches: readonly ExcelShortcut[],
  query: KeyQuery,
): NextKeyOptions {
  const stepIndex = Math.max(0, query.steps.length - 1);
  const current = query.steps[stepIndex] ?? [];
  const available = new Set<KeyToken>();
  let canBeginNextStep = false;

  for (const shortcut of matches) {
    const step = shortcut.keys[stepIndex];
    if (step === undefined) {
      continue;
    }
    for (const token of step) {
      if (!current.includes(token)) {
        available.add(token);
      }
    }
    if (
      current.length > 0 &&
      shortcut.keys.length > stepIndex + 1 &&
      isSameTokenSet(current, step)
    ) {
      canBeginNextStep = true;
    }
  }

  const sorted = [...available].sort(compareTokens);
  return {
    canBeginNextStep,
    keys: sorted.filter((token) => !isModifierKeyToken(token)),
    modifiers: sorted.filter((token) => isModifierKeyToken(token)),
  };
}

/** Plain-text rendering of a sequence, used for titles and test assertions. */
export function formatKeySequence(keys: ShortcutKeys): string {
  return keys.map((step) => step.join("+")).join(", ");
}
