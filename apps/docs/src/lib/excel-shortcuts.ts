/**
 * The shape of the Excel keyboard shortcut database, and the vocabulary its
 * entries are written in.
 *
 * A shortcut is a sequence of steps, and a step is the set of keys held down
 * together. That one model covers both kinds of Excel shortcut: a chord such
 * as Ctrl+Shift+L is a single step, and a ribbon route such as Alt, H, O, I is
 * four steps of one key each. Matching can then treat "starts with" the same
 * way for both.
 *
 * Tokens name physical keys, not the characters a shifted key produces.
 * Microsoft's own lists write Ctrl+Shift+* and Ctrl+Shift+~, but the keys a
 * visitor actually presses are Ctrl+Shift+8 and Ctrl+Shift+`, and the search
 * on this site is driven by real key presses. The published spelling goes in
 * the entry's note where it differs.
 */

/** Keys that only ever modify another key. */
export const MODIFIER_KEY_TOKENS = ["Ctrl", "Shift", "Alt", "Win"] as const;

export const FUNCTION_KEY_TOKENS = [
  "F1",
  "F2",
  "F3",
  "F4",
  "F5",
  "F6",
  "F7",
  "F8",
  "F9",
  "F10",
  "F11",
  "F12",
] as const;

/** Keys whose name is a word rather than the character they type. */
export const NAMED_KEY_TOKENS = [
  "ArrowUp",
  "ArrowDown",
  "ArrowLeft",
  "ArrowRight",
  "Backspace",
  "Delete",
  "End",
  "Enter",
  "Escape",
  "Home",
  "Insert",
  "PageDown",
  "PageUp",
  "Space",
  "Tab",
] as const;

/**
 * Keys that type one character. Letters are stored uppercase and punctuation
 * is stored unshifted, so the token is the key cap rather than the character
 * a particular combination produces.
 */
export const CHARACTER_KEY_TOKENS = [
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P",
  "Q",
  "R",
  "S",
  "T",
  "U",
  "V",
  "W",
  "X",
  "Y",
  "Z",
  "0",
  "1",
  "2",
  "3",
  "4",
  "5",
  "6",
  "7",
  "8",
  "9",
  "`",
  "-",
  "=",
  "[",
  "]",
  "\\",
  ";",
  "'",
  ",",
  ".",
  "/",
] as const;

/** Every key token an entry in the database may name. */
export const CANONICAL_KEY_TOKENS = [
  ...MODIFIER_KEY_TOKENS,
  ...FUNCTION_KEY_TOKENS,
  ...NAMED_KEY_TOKENS,
  ...CHARACTER_KEY_TOKENS,
] as const;

export type ModifierKeyToken = (typeof MODIFIER_KEY_TOKENS)[number];
export type KeyToken = (typeof CANONICAL_KEY_TOKENS)[number];

/** The keys held down together in one step of a shortcut. */
export type ShortcutStep = readonly KeyToken[];

/** The steps of a shortcut, in the order they are pressed. */
export type ShortcutKeys = readonly ShortcutStep[];

/**
 * The categories an entry may declare, in the order the results are grouped.
 * A shortcut is filed by what it does, so a ribbon route such as Alt, H, O, I
 * sits under formatting with the other column-width shortcuts. The ribbon
 * sequences category is for the shortcuts that drive the ribbon itself: the
 * key tips, the tabs, and the File menu.
 */
export const SHORTCUT_CATEGORIES = [
  "navigation",
  "selection",
  "data entry",
  "formatting",
  "formulas",
  "tables and filtering",
  "pivot tables",
  "workbook and window",
  "ribbon sequences",
  "other",
] as const;

export type ShortcutCategory = (typeof SHORTCUT_CATEGORIES)[number];

export interface ExcelShortcut {
  /** Stable identifier, unique across the database. */
  readonly id: string;
  /** The Windows key sequence. */
  readonly keys: ShortcutKeys;
  /**
   * The macOS key sequence. Reserved for a later pass over the database: no
   * entry declares one yet, and the page filters the Windows sequences only.
   */
  readonly mac?: ShortcutKeys;
  /** What the shortcut does, in one line without a terminal full stop. */
  readonly action: string;
  readonly category: ShortcutCategory;
  /** An extra condition, caveat, or alternative spelling, when one applies. */
  readonly note?: string;
}

const canonicalTokens: ReadonlySet<string> = new Set(CANONICAL_KEY_TOKENS);

export function isKeyToken(value: string): value is KeyToken {
  return canonicalTokens.has(value);
}

export function isModifierKeyToken(token: KeyToken): boolean {
  return (MODIFIER_KEY_TOKENS as readonly string[]).includes(token);
}

/** Compact chip text. Arrows read better as glyphs than as their token. */
const KEY_TOKEN_LABELS: Readonly<Partial<Record<KeyToken, string>>> = {
  ArrowDown: "↓",
  ArrowLeft: "←",
  ArrowRight: "→",
  ArrowUp: "↑",
  Escape: "Esc",
  PageDown: "Page Down",
  PageUp: "Page Up",
};

/** Spoken name, for the accessible name of a chip a glyph would hide. */
const KEY_TOKEN_NAMES: Readonly<Partial<Record<KeyToken, string>>> = {
  ArrowDown: "Down arrow",
  ArrowLeft: "Left arrow",
  ArrowRight: "Right arrow",
  ArrowUp: "Up arrow",
  "`": "Backtick",
  "\\": "Backslash",
  "'": "Apostrophe",
  ",": "Comma",
  "-": "Hyphen",
  ".": "Full stop",
  "/": "Slash",
  ";": "Semicolon",
  "=": "Equals",
  "[": "Left bracket",
  "]": "Right bracket",
};

export function keyTokenLabel(token: KeyToken): string {
  return KEY_TOKEN_LABELS[token] ?? token;
}

export function keyTokenName(token: KeyToken): string {
  return KEY_TOKEN_NAMES[token] ?? token;
}
