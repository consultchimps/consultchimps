/**
 * L2: the typing opt-ins a CSV read can ask for (ADR 0007). A CSV field is text
 * unless the caller opts in, and an opt-in converts only a field that matches
 * its whole pattern; anything else stays the text it was.
 */
import {
  calendarIsoParts,
  calendarIsoText,
  serial1900,
} from "../model/calendar.js";

/** The orders a date may be read in. `iso` reads only `yyyy-mm-dd`. */
export type CsvDateOrder = "iso" | "dmy" | "mdy" | "ymd";

/** The most significant digits a double holds exactly. */
const MAX_SIGNIFICANT_DIGITS = 15;

function escapeForPattern(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
}

/**
 * Reads numbers written with one decimal separator and, optionally, one
 * thousands separator grouping the whole part in threes.
 */
export class CsvNumberReader {
  readonly #pattern: RegExp;
  readonly #thousands: string;

  constructor(decimal: string, thousands: string) {
    const point = escapeForPattern(decimal);
    // A whole part is a lone zero, or digits with no leading zero, grouped in
    // threes or not grouped at all.
    const whole =
      thousands === ""
        ? "0|[1-9]\\d*"
        : `0|[1-9]\\d{0,2}(?:${escapeForPattern(thousands)}\\d{3})+|[1-9]\\d*`;
    this.#pattern = new RegExp(`^([+-]?)(${whole})(?:${point}(\\d+))?$`, "u");
    this.#thousands = thousands;
  }

  /** The number the field holds, or undefined when it is not only a number. */
  read(field: string): number | undefined {
    const matched = this.#pattern.exec(field);
    if (!matched) return undefined;
    const sign = matched[1] ?? "";
    const whole =
      this.#thousands === ""
        ? (matched[2] ?? "")
        : (matched[2] ?? "").split(this.#thousands).join("");
    const fraction = matched[3] ?? "";
    const significant = `${whole}${fraction}`.replace(/^0+/u, "");
    if (significant.length > MAX_SIGNIFICANT_DIGITS) return undefined;
    const value = Number(
      fraction === "" ? `${sign}${whole}` : `${sign}${whole}.${fraction}`,
    );
    if (!Number.isFinite(value)) return undefined;
    // A nonzero field too small for a double reads as zero; it stays text.
    if (value === 0 && significant !== "") return undefined;
    // Negative zero writes as zero everywhere, so it is read as zero.
    return value === 0 ? 0 : value;
  }
}

const TIME = String.raw`(?:[T ](\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?`;
const ISO_DATE = new RegExp(
  String.raw`^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?$`,
  "u",
);
const DAY_FIRST = new RegExp(
  String.raw`^(\d{1,2})([/.-])(\d{1,2})\2(\d{4})${TIME}$`,
  "u",
);
const YEAR_FIRST = new RegExp(
  String.raw`^(\d{4})([/.-])(\d{1,2})\2(\d{1,2})${TIME}$`,
  "u",
);

/**
 * The workbook date spelling for a field holding one date in `order`, or
 * undefined when it does not, including a day the calendar does not have and
 * one before 1900, which a workbook cannot store as a date.
 */
export function csvDate(
  field: string,
  order: CsvDateOrder,
): string | undefined {
  let year: string | undefined;
  let month: string | undefined;
  let day: string | undefined;
  let time: ReadonlyArray<string | undefined>;
  const iso = ISO_DATE.exec(field);
  if (iso) {
    [year, month, day] = [iso[1], iso[2], iso[3]];
    time = iso.slice(4, 8);
  } else if (order === "dmy" || order === "mdy") {
    const matched = DAY_FIRST.exec(field);
    if (!matched) return undefined;
    year = matched[4];
    [day, month] =
      order === "dmy" ? [matched[1], matched[3]] : [matched[3], matched[1]];
    time = matched.slice(5, 9);
  } else if (order === "ymd") {
    const matched = YEAR_FIRST.exec(field);
    if (!matched) return undefined;
    [year, month, day] = [matched[1], matched[3], matched[4]];
    time = matched.slice(5, 9);
  } else {
    return undefined;
  }
  const [hour, minute, second, fraction] = time;
  const parts = {
    year: Number(year),
    month: Number(month),
    day: Number(day),
    hour: Number(hour ?? 0),
    minute: Number(minute ?? 0),
    second: Number(second ?? 0),
    millisecond: Number((fraction ?? "").padEnd(3, "0")),
  };
  if (parts.hour > 23 || parts.minute > 59 || parts.second > 59) {
    return undefined;
  }
  const text = calendarIsoText(parts);
  // The spelling's own reading checks the day exists.
  const checked = calendarIsoParts(text);
  if (checked === undefined || serial1900(checked) === undefined) {
    return undefined;
  }
  return text;
}
