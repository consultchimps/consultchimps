import { describe, expect, it } from "vitest";

import { EXCEL_SHORTCUTS } from "./excel-shortcuts-data";
import {
  isKeyToken,
  SHORTCUT_CATEGORIES,
  type ShortcutKeys,
} from "./excel-shortcuts";
import { formatKeySequence } from "./shortcut-search";

/**
 * The database is data, so these are the rules a reviewer would otherwise have
 * to hold in their head while reading 261 entries: the identifiers are usable
 * as React keys and anchors, the tokens are ones the page can render and the
 * keyboard capture can produce, and no two entries claim the same keys on the
 * same platform.
 */

/** A sequence reduced to a comparable string, ignoring the order within a chord. */
function normalizeSequence(keys: ShortcutKeys): string {
  return keys.map((step) => [...step].sort().join("+")).join(", ");
}

describe("the Excel shortcut database", () => {
  it("carries enough entries to be worth searching", () => {
    expect(EXCEL_SHORTCUTS.length).toBeGreaterThanOrEqual(200);
  });

  it("gives every category at least one entry", () => {
    for (const category of SHORTCUT_CATEGORIES) {
      expect(
        EXCEL_SHORTCUTS.filter((entry) => entry.category === category),
      ).not.toHaveLength(0);
    }
  });

  it("declares a category from the fixed list", () => {
    for (const entry of EXCEL_SHORTCUTS) {
      expect(SHORTCUT_CATEGORIES).toContain(entry.category);
    }
  });

  it("uses an identifier that is unique and safe in a URL", () => {
    const seen = new Set<string>();
    for (const entry of EXCEL_SHORTCUTS) {
      expect(entry.id).toMatch(/^[a-z0-9]+(?:-[a-z0-9]+)*$/u);
      expect(seen.has(entry.id)).toBe(false);
      seen.add(entry.id);
    }
    expect(seen.size).toBe(EXCEL_SHORTCUTS.length);
  });

  it("spells every key with a canonical token", () => {
    for (const entry of EXCEL_SHORTCUTS) {
      expect(entry.keys.length).toBeGreaterThan(0);
      for (const step of entry.keys) {
        expect(step.length).toBeGreaterThan(0);
        for (const token of step) {
          expect(isKeyToken(token)).toBe(true);
        }
        // A key held twice in one chord is a typo, not a shortcut.
        expect(new Set(step).size).toBe(step.length);
      }
    }
  });

  it("holds no duplicate key sequence within a platform", () => {
    const windowsSequences = new Map<string, string>();
    const macSequences = new Map<string, string>();
    for (const entry of EXCEL_SHORTCUTS) {
      const windows = normalizeSequence(entry.keys);
      expect(windowsSequences.get(windows) ?? entry.id).toBe(entry.id);
      windowsSequences.set(windows, entry.id);

      if (entry.mac !== undefined) {
        const mac = normalizeSequence(entry.mac);
        expect(macSequences.get(mac) ?? entry.id).toBe(entry.id);
        macSequences.set(mac, entry.id);
      }
    }
    expect(windowsSequences.size).toBe(EXCEL_SHORTCUTS.length);
  });

  it("leaves the mac field for the later pass, and validates it if one appears", () => {
    // The field is part of the schema today and unused: an entry that starts
    // declaring one is held to the same token rules as the Windows sequence.
    for (const entry of EXCEL_SHORTCUTS) {
      if (entry.mac === undefined) {
        continue;
      }
      expect(entry.mac.length).toBeGreaterThan(0);
      for (const step of entry.mac) {
        expect(step.length).toBeGreaterThan(0);
        for (const token of step) {
          expect(isKeyToken(token)).toBe(true);
        }
      }
    }
  });

  it("writes the rendered strings without a terminal full stop", () => {
    for (const entry of EXCEL_SHORTCUTS) {
      // The site's rule for interface copy: a sentence keeps its internal
      // stops and loses the last one.
      expect(entry.action.trim()).toBe(entry.action);
      expect(entry.action).not.toBe("");
      expect(entry.action.endsWith(".")).toBe(false);
      if (entry.note !== undefined) {
        expect(entry.note.trim()).toBe(entry.note);
        expect(entry.note).not.toBe("");
        expect(entry.note.endsWith(".")).toBe(false);
      }
    }
  });

  it("starts each action with a capital letter", () => {
    for (const entry of EXCEL_SHORTCUTS) {
      expect(entry.action.slice(0, 1)).toBe(
        entry.action.slice(0, 1).toUpperCase(),
      );
    }
  });

  it("models chords as one step and ribbon routes as several", () => {
    const chord = EXCEL_SHORTCUTS.find(
      (entry) => entry.id === "tbl-toggle-filters",
    );
    expect(chord && formatKeySequence(chord.keys)).toBe("Ctrl+Shift+L");
    const sequence = EXCEL_SHORTCUTS.find(
      (entry) => entry.id === "fmt-autofit-column-width",
    );
    expect(sequence && formatKeySequence(sequence.keys)).toBe("Alt, H, O, I");
  });
});
