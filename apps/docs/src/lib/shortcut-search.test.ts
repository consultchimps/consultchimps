import { describe, expect, it } from "vitest";

import type { ExcelShortcut } from "./excel-shortcuts";
import {
  appendKeyToken,
  beginKeyStep,
  EMPTY_KEY_QUERY,
  filterShortcuts,
  formatKeySequence,
  groupShortcutsByCategory,
  isEmptyKeyQuery,
  isMatchedKeyToken,
  keyQueryTokenCount,
  keyTokenFromPress,
  matchesKeyQuery,
  matchesWordQuery,
  nextKeyOptions,
  removeKeyTokens,
  removeLastKeyEntry,
  type KeyQuery,
} from "./shortcut-search";

/** A query written the way a test reads it, rather than key press by press. */
const query = (...steps: string[][]): KeyQuery => ({
  steps: steps as KeyQuery["steps"],
});

const shortcut = (
  id: string,
  keys: string[][],
  overrides: Partial<ExcelShortcut> = {},
): ExcelShortcut =>
  ({
    id,
    keys,
    action: `Do ${id}`,
    category: "other",
    ...overrides,
  }) as ExcelShortcut;

const TOGGLE_FILTERS = shortcut("filters", [["Ctrl", "Shift", "L"]], {
  action: "Add or remove the filter row",
  category: "tables and filtering",
});
const GROUP_ITEMS = shortcut("group", [["Alt", "Shift", "ArrowRight"]], {
  action: "Group the selected PivotTable items",
  category: "pivot tables",
});
const AUTOFIT = shortcut("autofit", [["Alt"], ["H"], ["O"], ["I"]], {
  action: "Fit the column width to the widest entry",
  category: "formatting",
});
const NEW_LINE = shortcut("new-line", [["Alt", "Enter"]], {
  action: "Start a new line inside the same cell",
  category: "data entry",
});
const DATABASE = [TOGGLE_FILTERS, GROUP_ITEMS, AUTOFIT, NEW_LINE];

describe("matchesKeyQuery", () => {
  it("matches everything while the query is empty", () => {
    for (const entry of DATABASE) {
      expect(matchesKeyQuery(entry.keys, EMPTY_KEY_QUERY)).toBe(true);
    }
  });

  it("keeps a chord whose step merely starts with the keys held", () => {
    // The core of the page: one modifier narrows to the chords that contain
    // it, rather than to the chords that are only that modifier.
    expect(matchesKeyQuery(TOGGLE_FILTERS.keys, query(["Shift"]))).toBe(true);
    expect(matchesKeyQuery(GROUP_ITEMS.keys, query(["Shift"]))).toBe(true);
    expect(matchesKeyQuery(AUTOFIT.keys, query(["Shift"]))).toBe(false);
  });

  it("ignores the order the keys of a chord were pressed in", () => {
    // Shift then Alt, against a chord stored as Alt+Shift+ArrowRight.
    expect(matchesKeyQuery(GROUP_ITEMS.keys, query(["Shift", "Alt"]))).toBe(
      true,
    );
    expect(matchesKeyQuery(GROUP_ITEMS.keys, query(["Alt", "Shift"]))).toBe(
      true,
    );
    // A key the chord does not hold still rules it out, whatever the order.
    expect(matchesKeyQuery(GROUP_ITEMS.keys, query(["Shift", "Ctrl"]))).toBe(
      false,
    );
  });

  it("narrows as more keys join the same chord", () => {
    expect(matchesKeyQuery(TOGGLE_FILTERS.keys, query(["Ctrl"]))).toBe(true);
    expect(matchesKeyQuery(TOGGLE_FILTERS.keys, query(["Ctrl", "Shift"]))).toBe(
      true,
    );
    expect(
      matchesKeyQuery(TOGGLE_FILTERS.keys, query(["Ctrl", "Shift", "L"])),
    ).toBe(true);
    expect(
      matchesKeyQuery(TOGGLE_FILTERS.keys, query(["Ctrl", "Shift", "K"])),
    ).toBe(false);
  });

  it("walks a ribbon sequence step by step", () => {
    expect(matchesKeyQuery(AUTOFIT.keys, query(["Alt"]))).toBe(true);
    expect(matchesKeyQuery(AUTOFIT.keys, query(["Alt"], ["H"]))).toBe(true);
    expect(matchesKeyQuery(AUTOFIT.keys, query(["Alt"], ["H"], ["O"]))).toBe(
      true,
    );
    expect(matchesKeyQuery(AUTOFIT.keys, query(["Alt"], ["H"], ["P"]))).toBe(
      false,
    );
  });

  it("treats a second step as proof the shortcut is not a chord", () => {
    // Alt alone keeps Alt+Enter, because a chord may still grow.
    expect(matchesKeyQuery(NEW_LINE.keys, query(["Alt"]))).toBe(true);
    // Committing a second step says the first one is finished, so a
    // single-step chord can no longer match.
    expect(matchesKeyQuery(NEW_LINE.keys, query(["Alt"], []))).toBe(false);
    expect(matchesKeyQuery(NEW_LINE.keys, query(["Alt"], ["H"]))).toBe(false);
    expect(matchesKeyQuery(AUTOFIT.keys, query(["Alt"], []))).toBe(true);
  });

  it("requires a finished step to match as a whole", () => {
    // Alt+Shift is a subset of the first step of nothing here, and as a
    // finished step it must equal the step it sits against.
    expect(matchesKeyQuery(AUTOFIT.keys, query(["Alt", "Shift"], ["H"]))).toBe(
      false,
    );
  });

  it("rejects a query longer than the shortcut", () => {
    expect(matchesKeyQuery(NEW_LINE.keys, query(["Alt", "Enter"], ["H"]))).toBe(
      false,
    );
  });
});

describe("building a key query", () => {
  it("starts empty", () => {
    expect(isEmptyKeyQuery(EMPTY_KEY_QUERY)).toBe(true);
    expect(keyQueryTokenCount(EMPTY_KEY_QUERY)).toBe(0);
  });

  it("adds keys to the step being built", () => {
    const held = appendKeyToken(
      appendKeyToken(EMPTY_KEY_QUERY, "Ctrl"),
      "Shift",
    );
    expect(held.steps).toEqual([["Ctrl", "Shift"]]);
    expect(keyQueryTokenCount(held)).toBe(2);
    expect(isEmptyKeyQuery(held)).toBe(false);
  });

  it("ignores a key that is already held, as auto-repeat sends it", () => {
    const held = appendKeyToken(EMPTY_KEY_QUERY, "Ctrl");
    expect(appendKeyToken(held, "Ctrl")).toBe(held);
  });

  it("opens a new step once the keys are released", () => {
    const first = appendKeyToken(EMPTY_KEY_QUERY, "Alt");
    const second = appendKeyToken(beginKeyStep(first), "H");
    expect(second.steps).toEqual([["Alt"], ["H"]]);
  });

  it("does not open a second empty step in a row", () => {
    const opened = beginKeyStep(appendKeyToken(EMPTY_KEY_QUERY, "Alt"));
    expect(beginKeyStep(opened)).toBe(opened);
    expect(beginKeyStep(EMPTY_KEY_QUERY)).toBe(EMPTY_KEY_QUERY);
  });

  it("removes the last key, then the step it sat in", () => {
    const built = appendKeyToken(
      beginKeyStep(appendKeyToken(EMPTY_KEY_QUERY, "Alt")),
      "H",
    );
    const withoutH = removeLastKeyEntry(built);
    expect(withoutH.steps).toEqual([["Alt"], []]);
    const withoutStep = removeLastKeyEntry(withoutH);
    expect(withoutStep.steps).toEqual([["Alt"]]);
    expect(removeLastKeyEntry(withoutStep)).toEqual(EMPTY_KEY_QUERY);
  });

  it("leaves an empty query alone when there is nothing to remove", () => {
    expect(removeLastKeyEntry(EMPTY_KEY_QUERY)).toBe(EMPTY_KEY_QUERY);
  });

  it("takes the keys of a navigation gesture back out", () => {
    const held = appendKeyToken(EMPTY_KEY_QUERY, "Shift");
    // Shift held only to press Shift+Tab leaves nothing filtering the list.
    expect(removeKeyTokens(held, ["Shift"])).toEqual(EMPTY_KEY_QUERY);

    // A key entered earlier, with the buttons, survives the gesture.
    const mixed = appendKeyToken(
      appendKeyToken(EMPTY_KEY_QUERY, "Ctrl"),
      "Shift",
    );
    expect(removeKeyTokens(mixed, ["Shift"]).steps).toEqual([["Ctrl"]]);

    // Nothing held, nothing to give back.
    expect(removeKeyTokens(mixed, [])).toBe(mixed);
    expect(removeKeyTokens(mixed, ["Alt"])).toBe(mixed);
    expect(removeKeyTokens(EMPTY_KEY_QUERY, ["Shift"])).toBe(EMPTY_KEY_QUERY);
  });

  it("drops the step a gesture empties, leaving the query it interrupted", () => {
    // Shift pressed after a finished chord opens a step to hold it. Undoing
    // the gesture has to take that step with it: an empty trailing step is a
    // second step, and the one-step chord would stop matching.
    const finished = query(["Ctrl", "Shift", "L"]);
    const withGesture = appendKeyToken(beginKeyStep(finished), "Shift");
    expect(withGesture.steps).toEqual([["Ctrl", "Shift", "L"], ["Shift"]]);
    expect(removeKeyTokens(withGesture, ["Shift"]).steps).toEqual([
      ["Ctrl", "Shift", "L"],
    ]);
    expect(
      matchesKeyQuery(
        TOGGLE_FILTERS.keys,
        removeKeyTokens(withGesture, ["Shift"]),
      ),
    ).toBe(true);

    const secondStep = appendKeyToken(
      beginKeyStep(appendKeyToken(EMPTY_KEY_QUERY, "Alt")),
      "Shift",
    );
    expect(removeKeyTokens(secondStep, ["Shift"]).steps).toEqual([["Alt"]]);
  });

  it("removes one key of a chord at a time", () => {
    const chord = appendKeyToken(
      appendKeyToken(appendKeyToken(EMPTY_KEY_QUERY, "Ctrl"), "Shift"),
      "L",
    );
    expect(removeLastKeyEntry(chord).steps).toEqual([["Ctrl", "Shift"]]);
  });
});

describe("matchesWordQuery", () => {
  it("matches nothing away when the box is empty", () => {
    expect(matchesWordQuery(AUTOFIT, "   ")).toBe(true);
  });

  it("searches the action without regard to case", () => {
    expect(matchesWordQuery(AUTOFIT, "COLUMN width")).toBe(true);
    expect(matchesWordQuery(AUTOFIT, "pivot")).toBe(false);
  });

  it("searches the category and the note as well", () => {
    expect(matchesWordQuery(GROUP_ITEMS, "pivot tables")).toBe(true);
    const noted = shortcut("noted", [["Ctrl", "9"]], {
      action: "Hide the selected rows",
      note: "Published as something else",
    });
    expect(matchesWordQuery(noted, "published")).toBe(true);
  });
});

describe("filterShortcuts", () => {
  it("requires both filters to pass", () => {
    const matches = filterShortcuts(DATABASE, {
      keys: query(["Alt"]),
      words: "pivot",
    });
    expect(matches.map((entry) => entry.id)).toEqual(["group"]);
  });

  it("applies the word filter on its own", () => {
    const matches = filterShortcuts(DATABASE, {
      keys: EMPTY_KEY_QUERY,
      words: "filter row",
    });
    expect(matches.map((entry) => entry.id)).toEqual(["filters"]);
  });

  it("applies the key filter on its own", () => {
    const matches = filterShortcuts(DATABASE, {
      keys: query(["Alt"]),
      words: "",
    });
    expect(matches.map((entry) => entry.id)).toEqual([
      "group",
      "autofit",
      "new-line",
    ]);
  });

  it("returns nothing when the two filters disagree", () => {
    expect(
      filterShortcuts(DATABASE, {
        keys: query(["Ctrl"]),
        words: "pivot",
      }),
    ).toEqual([]);
  });
});

describe("groupShortcutsByCategory", () => {
  it("groups in the declared category order and drops empty categories", () => {
    const groups = groupShortcutsByCategory(DATABASE);
    expect(groups.map((group) => group.category)).toEqual([
      "data entry",
      "formatting",
      "tables and filtering",
      "pivot tables",
    ]);
    expect(groups.at(0)?.shortcuts.map((entry) => entry.id)).toEqual([
      "new-line",
    ]);
  });

  it("returns nothing for an empty result list", () => {
    expect(groupShortcutsByCategory([])).toEqual([]);
  });
});

describe("isMatchedKeyToken", () => {
  it("marks the keys the query named, step by step", () => {
    const entered = query(["Alt"], ["H"]);
    expect(isMatchedKeyToken(entered, 0, "Alt")).toBe(true);
    expect(isMatchedKeyToken(entered, 1, "H")).toBe(true);
    expect(isMatchedKeyToken(entered, 2, "O")).toBe(false);
  });

  it("marks only the held keys of a partial chord", () => {
    const entered = query(["Ctrl", "Shift"]);
    expect(isMatchedKeyToken(entered, 0, "Ctrl")).toBe(true);
    expect(isMatchedKeyToken(entered, 0, "L")).toBe(false);
  });
});

describe("nextKeyOptions", () => {
  it("offers the first keys of the database, modifiers in their usual order", () => {
    const options = nextKeyOptions(DATABASE, EMPTY_KEY_QUERY);
    expect(options.modifiers).toEqual(["Ctrl", "Shift", "Alt"]);
    expect(options.keys).toEqual(["ArrowRight", "Enter", "L"]);
    expect(options.canBeginNextStep).toBe(false);
  });

  it("offers only the keys that keep a result", () => {
    const matches = filterShortcuts(DATABASE, {
      keys: query(["Ctrl"]),
      words: "",
    });
    const options = nextKeyOptions(matches, query(["Ctrl"]));
    expect(options.modifiers).toEqual(["Shift"]);
    expect(options.keys).toEqual(["L"]);
  });

  it("reports when a finished step could be followed by another", () => {
    const matches = filterShortcuts(DATABASE, {
      keys: query(["Alt"]),
      words: "",
    });
    expect(nextKeyOptions(matches, query(["Alt"])).canBeginNextStep).toBe(true);
    // A chord that is nobody's complete first step cannot start a second one.
    const chordMatches = filterShortcuts(DATABASE, {
      keys: query(["Alt", "Shift"]),
      words: "",
    });
    expect(
      nextKeyOptions(chordMatches, query(["Alt", "Shift"])).canBeginNextStep,
    ).toBe(false);
  });

  it("offers the keys of the step now being built", () => {
    const matches = filterShortcuts(DATABASE, {
      keys: query(["Alt"], []),
      words: "",
    });
    const options = nextKeyOptions(matches, query(["Alt"], []));
    expect(options.keys).toEqual(["H"]);
  });
});

describe("keyTokenFromPress", () => {
  it("names the modifiers the way the database spells them", () => {
    expect(keyTokenFromPress({ code: "ControlLeft", key: "Control" })).toBe(
      "Ctrl",
    );
    expect(keyTokenFromPress({ code: "ShiftRight", key: "Shift" })).toBe(
      "Shift",
    );
    expect(keyTokenFromPress({ code: "AltLeft", key: "Alt" })).toBe("Alt");
    expect(keyTokenFromPress({ code: "MetaLeft", key: "Meta" })).toBe("Win");
  });

  it("reads a letter as the letter, in either case", () => {
    expect(keyTokenFromPress({ code: "KeyL", key: "l" })).toBe("L");
    expect(keyTokenFromPress({ code: "KeyL", key: "L" })).toBe("L");
  });

  it("takes the letter the layout typed, not the key's US position", () => {
    // On a French layout the key marked A sits where a US keyboard has Q, so
    // the position would record the wrong letter.
    expect(keyTokenFromPress({ code: "KeyQ", key: "a" })).toBe("A");
    expect(keyTokenFromPress({ code: "KeyA", key: "q" })).toBe("Q");
    // A layout whose character is not one the database spells still falls
    // back to the position rather than dropping the press.
    expect(keyTokenFromPress({ code: "KeyL", key: "ł" })).toBe("L");
  });

  it("reads a digit as the digit, not the character Shift types", () => {
    // Ctrl+Shift+8 is published as Ctrl+Shift+*, and the press reports "*".
    expect(keyTokenFromPress({ code: "Digit8", key: "*" })).toBe("8");
    expect(keyTokenFromPress({ code: "Digit3", key: "#" })).toBe("3");
  });

  it("reads punctuation as the unshifted key cap", () => {
    expect(keyTokenFromPress({ code: "Backquote", key: "~" })).toBe("`");
    expect(keyTokenFromPress({ code: "Semicolon", key: ":" })).toBe(";");
    expect(keyTokenFromPress({ code: "Backslash", key: "|" })).toBe("\\");
  });

  it("passes the named and function keys through", () => {
    expect(keyTokenFromPress({ code: "ArrowRight", key: "ArrowRight" })).toBe(
      "ArrowRight",
    );
    expect(keyTokenFromPress({ code: "F5", key: "F5" })).toBe("F5");
    expect(keyTokenFromPress({ code: "Space", key: " " })).toBe("Space");
    expect(keyTokenFromPress({ code: "PageDown", key: "PageDown" })).toBe(
      "PageDown",
    );
  });

  it("returns nothing for a key the database does not model", () => {
    expect(keyTokenFromPress({ code: "CapsLock", key: "CapsLock" })).toBeNull();
    expect(
      keyTokenFromPress({ code: "ScrollLock", key: "ScrollLock" }),
    ).toBeNull();
  });
});

describe("formatKeySequence", () => {
  it("writes a chord with plus signs and a sequence with commas", () => {
    expect(formatKeySequence(TOGGLE_FILTERS.keys)).toBe("Ctrl+Shift+L");
    expect(formatKeySequence(AUTOFIT.keys)).toBe("Alt, H, O, I");
  });
});
