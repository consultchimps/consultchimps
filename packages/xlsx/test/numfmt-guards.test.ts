/**
 * The two corrections `src/operations/numfmt-guards.ts` makes to numfmt.
 * The "upstream" tests pin numfmt's own wrong answers: when one fails after a
 * numfmt upgrade, upstream has fixed that case, so delete its guard and the
 * pin together.
 */
import { format } from "numfmt";
import { describe, expect, it } from "vitest";

import { formatDisplayText } from "../src/operations/display-text.js";

const OPTIONS = { locale: "en-US", nbsp: false, leap1900: true } as const;

describe("guard (a): General rounds 10 and 11 digit numbers", () => {
  it("upstream: numfmt 3.2.6 still truncates them", () => {
    // Fails once numfmt rounds: delete generalRoundingGuard then.
    expect(format("General", 4403928373.5, OPTIONS)).toBe("4403928373");
    expect(format("General", 12345678901.5, OPTIONS)).toBe("12345678901");
  });

  it.each([
    [4403928373.5, "4403928374"],
    [-4511987749.6, "-4511987750"],
    [1234567890.4, "1234567890"],
    [9999999999.6, "10000000000"],
    [12345678901.5, "12345678902"],
    [99999999999.6, "1E+11"],
    [-99999999999.6, "-1E+11"],
    // Untouched: integers, and numbers numfmt already formats as Excel does.
    [12345670000, "12345670000"],
    [123456789.5, "123456789.5"],
  ])("shows %s as %s", (value, text) => {
    expect(formatDisplayText("General", value, false)).toBe(text);
  });
});

describe("guard (b): a time rounding up to midnight moves to the next day", () => {
  it("upstream: numfmt 3.2.6 still keeps the old date", () => {
    // Fails once numfmt carries the day: delete midnightCarryGuard then.
    expect(format("m/d/yy h:mm", 45292.9999999, OPTIONS)).toBe("1/1/24 0:00");
  });

  it.each([
    ["m/d/yy h:mm", 45292.9999999, false, "1/2/24 0:00"],
    [
      "yyyy-mm-dd hh:mm:ss",
      45292 + 86399.6 / 86400,
      false,
      "2024-01-02 00:00:00",
    ],
    ["dddd h:mm", 45292.9999999, false, "Tuesday 0:00"],
    ["m/d/yy h:mm", 45292.9999999 - 1462, true, "1/2/24 0:00"],
    // No carry: the second rounds down, or the shown tenths absorb it.
    [
      "yyyy-mm-dd hh:mm:ss",
      45292 + 86399.4 / 86400,
      false,
      "2024-01-01 23:59:59",
    ],
    [
      "yyyy-mm-dd hh:mm:ss.0",
      45292 + 86399.94 / 86400,
      false,
      "2024-01-01 23:59:59.9",
    ],
    ["[h]:mm", 0.9999999, false, "24:00"],
  ])("formats %s of %s (1904: %s) as %s", (code, value, date1904, text) => {
    expect(formatDisplayText(code, value, date1904)).toBe(text);
  });
});
