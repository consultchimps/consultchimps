import fc from "fast-check";
import { describe, expect, it } from "vitest";

import { safeNameFragment, truncateToUtf8Bytes } from "../src/index.js";

const FALLBACK = "fallback";
const MAX_FRAGMENT_BYTES = 80;
// A reserved device name is guarded with a "_" prefix after the byte cap has
// already been applied, so a guarded fragment can be one byte over the cap.
const MAX_GUARDED_BYTES = MAX_FRAGMENT_BYTES + 1;

const RESERVED_DEVICE_NAMES = [
  "con",
  "prn",
  "aux",
  "nul",
  ...Array.from({ length: 9 }, (_, index) => `com${index + 1}`),
  ...Array.from({ length: 9 }, (_, index) => `lpt${index + 1}`),
];
const RESERVED_WITH_EXTENSION =
  /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\..*)?$/iu;
const UNSAFE_CHARACTERS = /[<>:"/\\|?*]/u;

const encoder = new TextEncoder();

function encodedLength(value: string): number {
  return encoder.encode(value).length;
}

function hasControlCharacters(value: string): boolean {
  for (const character of value) {
    const codePoint = character.codePointAt(0);
    if (codePoint !== undefined && (codePoint < 32 || codePoint === 127)) {
      return true;
    }
  }
  return false;
}

/** Strip the reserved-device-name guard so the byte cap can be checked. */
function withoutReservedGuard(value: string): string {
  return value.startsWith("_") ? value.slice(1) : value;
}

/** The complete portability guarantee, asserted as one unit. */
function expectPortableFragment(fragment: string): void {
  expect(fragment.length).toBeGreaterThan(0);
  expect(encodedLength(withoutReservedGuard(fragment))).toBeLessThanOrEqual(
    MAX_FRAGMENT_BYTES,
  );
  expect(encodedLength(fragment)).toBeLessThanOrEqual(MAX_GUARDED_BYTES);
  expect(hasControlCharacters(fragment)).toBe(false);
  expect(UNSAFE_CHARACTERS.test(fragment)).toBe(false);
  expect(fragment.endsWith(".")).toBe(false);
  expect(fragment.endsWith(" ")).toBe(false);
  expect(RESERVED_DEVICE_NAMES).not.toContain(fragment.toLowerCase());
  expect(RESERVED_WITH_EXTENSION.test(fragment)).toBe(false);
}

// Arbitrary text including control characters, astral-plane code points and
// compatibility characters, so normalization and stripping both get exercised.
const anyText = fc.string({ unit: "binary", maxLength: 200 });
// Printable text only: control characters are the one input class that makes a
// single sanitizing pass non-idempotent (see the documented exception below).
const printableText = fc.string({ unit: "grapheme", maxLength: 200 });

// A fixed seed keeps the suite deterministic, matching the repository's
// promise that identical inputs always produce identical results.
const runs = { numRuns: 1000, seed: 20260805 };

describe("safeNameFragment properties", () => {
  it("always produces a non-empty fragment", () => {
    fc.assert(
      fc.property(anyText, (value) => {
        expect(safeNameFragment(value, FALLBACK).length).toBeGreaterThan(0);
      }),
      runs,
    );
  });

  it("keeps the fragment inside the UTF-8 byte cap", () => {
    fc.assert(
      fc.property(anyText, (value) => {
        const fragment = safeNameFragment(value, FALLBACK);
        expect(
          encodedLength(withoutReservedGuard(fragment)),
        ).toBeLessThanOrEqual(MAX_FRAGMENT_BYTES);
        expect(encodedLength(fragment)).toBeLessThanOrEqual(MAX_GUARDED_BYTES);
      }),
      runs,
    );
  });

  it("never emits control characters", () => {
    fc.assert(
      fc.property(anyText, (value) => {
        expect(hasControlCharacters(safeNameFragment(value, FALLBACK))).toBe(
          false,
        );
      }),
      runs,
    );
  });

  it("never emits characters that are unsafe in a filename", () => {
    fc.assert(
      fc.property(anyText, (value) => {
        expect(UNSAFE_CHARACTERS.test(safeNameFragment(value, FALLBACK))).toBe(
          false,
        );
      }),
      runs,
    );
  });

  it("never ends with a dot or a space", () => {
    fc.assert(
      fc.property(anyText, (value) => {
        const fragment = safeNameFragment(value, FALLBACK);
        expect(fragment.endsWith(".")).toBe(false);
        expect(fragment.endsWith(" ")).toBe(false);
      }),
      runs,
    );
  });

  it("never produces a Windows reserved device name", () => {
    fc.assert(
      fc.property(anyText, (value) => {
        const fragment = safeNameFragment(value, FALLBACK);
        expect(RESERVED_DEVICE_NAMES).not.toContain(fragment.toLowerCase());
        expect(RESERVED_WITH_EXTENSION.test(fragment)).toBe(false);
      }),
      runs,
    );
  });

  it("is idempotent for printable names", () => {
    fc.assert(
      fc.property(printableText, (value) => {
        fc.pre(!hasControlCharacters(value));
        const fragment = safeNameFragment(value, FALLBACK);
        fc.pre(encodedLength(fragment) <= MAX_FRAGMENT_BYTES);
        expect(safeNameFragment(fragment, FALLBACK)).toBe(fragment);
      }),
      runs,
    );
  });

  it("reaches a fixed point after a second pass for any input", () => {
    fc.assert(
      fc.property(anyText, (value) => {
        const second = safeNameFragment(
          safeNameFragment(value, FALLBACK),
          FALLBACK,
        );
        expect(safeNameFragment(second, FALLBACK)).toBe(second);
      }),
      runs,
    );
  });
});

// The fallback is part of the public API surface, so a dynamic or untrusted
// fallback must not be able to bypass the rules the value goes through.
describe("safeNameFragment sanitizes the fallback", () => {
  it("holds the whole guarantee for arbitrary value and fallback pairs", () => {
    fc.assert(
      fc.property(anyText, anyText, (value, fallback) => {
        expectPortableFragment(safeNameFragment(value, fallback));
      }),
      runs,
    );
  });

  it("holds the whole guarantee when the value sanitizes away", () => {
    fc.assert(
      fc.property(anyText, (fallback) => {
        expectPortableFragment(safeNameFragment("", fallback));
        expectPortableFragment(safeNameFragment("   ", fallback));
        expectPortableFragment(safeNameFragment("...", fallback));
      }),
      runs,
    );
  });

  it("leaves the fallbacks used by the operations byte-identical", () => {
    // Every fallback passed inside @consultchimps/pdf and @consultchimps/xlsx
    // is already a clean fragment, so no existing output name changes.
    for (const fallback of ["document", "combined", "split", "value"]) {
      expect(safeNameFragment("", fallback)).toBe(fallback);
      expect(safeNameFragment("\u0007", fallback)).toBe(fallback);
    }
  });

  it("guards a reserved fallback", () => {
    expect(safeNameFragment("", "con")).toBe("_con");
    expect(safeNameFragment("", "  NUL.  ")).toBe("_NUL");
  });

  it("caps and cleans an unsafe fallback", () => {
    expect(safeNameFragment("", 'a<b>c:"d/e')).toBe("a-b-c-d-e");
    expect(safeNameFragment("", "trailing. ")).toBe("trailing");
    expect(safeNameFragment("", "with\u0007control")).toBe("withcontrol");

    const capped = safeNameFragment("", "\u4e2d".repeat(100));
    expect(encodedLength(capped)).toBe(78);
    expect([...capped]).toHaveLength(26);
  });

  it("returns a hardcoded name when both inputs sanitize away", () => {
    expect(safeNameFragment("", "")).toBe("file");
    expect(safeNameFragment("\u0007", "  ...  ")).toBe("file");
    expect(safeNameFragment("...", "\u0000")).toBe("file");
  });
});

// The sanitizer is deliberately a single pass; these two inputs are the only
// known cases where re-sanitizing its output changes it. They are recorded
// here so the behaviour is pinned rather than accidental, and because both
// already produce names that are valid on every supported platform.
describe("safeNameFragment documented single-pass exceptions", () => {
  it("can exceed the byte cap by the reserved-name guard", () => {
    const fragment = safeNameFragment(`con.${"a".repeat(76)}`, FALLBACK);

    expect(fragment.startsWith("_con.")).toBe(true);
    expect(encodedLength(fragment)).toBe(MAX_GUARDED_BYTES);
    // A second pass no longer sees a reserved name and re-applies the cap.
    expect(encodedLength(safeNameFragment(fragment, FALLBACK))).toBe(
      MAX_FRAGMENT_BYTES,
    );
  });

  it("can leave a decomposed sequence when a control character is removed", () => {
    // NFKC runs before control characters are stripped, so a BEL between a
    // base letter and its combining mark survives normalization and leaves
    // the pair uncomposed.
    const fragment = safeNameFragment("u\u0007\u0301", FALLBACK);

    // "u" plus a combining acute accent, still decomposed.
    expect(fragment).toBe("u\u0301");
    // A second pass composes it into the single precomposed code point.
    expect(safeNameFragment(fragment, FALLBACK)).toBe("\u00fa");
  });
});

describe("truncateToUtf8Bytes properties", () => {
  it("never exceeds the requested byte budget", () => {
    fc.assert(
      fc.property(anyText, fc.integer({ min: 0, max: 200 }), (value, max) => {
        expect(
          encodedLength(truncateToUtf8Bytes(value, max)),
        ).toBeLessThanOrEqual(max);
      }),
      runs,
    );
  });

  it("returns a prefix that cuts only at code point boundaries", () => {
    fc.assert(
      fc.property(anyText, fc.integer({ min: 0, max: 200 }), (value, max) => {
        const truncated = truncateToUtf8Bytes(value, max);
        expect(value.startsWith(truncated)).toBe(true);
        expect([...truncated]).toEqual(
          [...value].slice(0, [...truncated].length),
        );
      }),
      runs,
    );
  });

  it("returns the whole value when the budget covers it", () => {
    fc.assert(
      fc.property(anyText, (value) => {
        expect(truncateToUtf8Bytes(value, encodedLength(value))).toBe(value);
      }),
      runs,
    );
  });
});
