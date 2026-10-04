/**
 * L3: the text a cell shows, its number format applied (ADR 0006). The
 * formatting is numfmt's, a maintained implementation of the ECMA-376 format
 * language; this module only adapts a workbook to it.
 */
import { format, isDateFormat, isValidFormat } from "numfmt";

import { generalRoundingGuard, midnightCarryGuard } from "./numfmt-guards.js";

/** Days between the 1900 and 1904 date systems' day zero. */
const DATE_1904_OFFSET = 1462;

const OPTIONS = {
  // Excel's own English defaults, with ordinary spaces.
  locale: "en-US",
  nbsp: false,
  // Serial 60 is 1900-02-29, as Excel counts it.
  leap1900: true,
  dateSpanLarge: false,
  // A date outside Excel's range shows as its number rather than failing.
  dateErrorNumber: true,
  throws: false,
} as const;

/**
 * A currency or locale tag such as `[$€-407]`. Excel keeps the reader's own
 * separators for these, but numfmt would switch to the tagged locale's, so
 * the locale part is dropped and the symbol kept.
 */
const LOCALE_TAG = /\[\$([^\]-]*)-[0-9A-Fa-f]+\]/gu;

/**
 * An elapsed-time token such as `[h]`. It shows a duration, which is the same
 * in both date systems, so its number is never moved between them.
 */
const ELAPSED_TOKEN = /\[(?:h+|m+|s+)\]/iu;

/**
 * The text `value` shows under the format `code`, or undefined when the code
 * cannot be read. A date format reads a number in the workbook's date system.
 */
export function formatDisplayText(
  code: string,
  value: number | string | boolean,
  date1904: boolean,
): string | undefined {
  const pattern = code.replace(LOCALE_TAG, (_, symbol: string) =>
    symbol === "" ? "" : `[$${symbol}]`,
  );
  if (!isValidFormat(pattern)) return undefined;
  const general = generalRoundingGuard(pattern, value);
  if (general !== undefined) return general;
  if (typeof value !== "number" || !isDateFormat(pattern)) {
    return format(pattern, value, OPTIONS);
  }
  const shifted =
    date1904 && !ELAPSED_TOKEN.test(pattern.replace(/"[^"]*"/gu, ""))
      ? value + DATE_1904_OFFSET
      : value;
  return format(
    pattern,
    midnightCarryGuard(pattern, shifted) ?? shifted,
    OPTIONS,
  );
}
