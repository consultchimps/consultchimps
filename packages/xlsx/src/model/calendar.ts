/**
 * L1: calendar components, and the one spelling every date in this package
 * gets.
 *
 * A workbook names a moment in two ways: a serial counted from the workbook's
 * epoch, and, for a `t="d"` cell, ISO 8601 text with no zone on it. Both arrive
 * as components, and both leave as the same characters, because a value read
 * from a file has to be a function of the file.
 *
 * Nothing here builds a `Date` from components. `Date.UTC` and `new Date(text)`
 * carry rules of their own that quietly rewrite what they are given: a year
 * from 0 to 99 is remapped into the twentieth century, so 0099 becomes 1999,
 * and an out-of-range field is normalised rather than refused, so month 13
 * becomes January of the next year. Those rules turn a malformed value into a
 * plausible wrong one, which is the worst outcome available. The arithmetic
 * below has no such rules: it converts, and the caller decides beforehand
 * whether the components are worth converting.
 */

/** A moment named by its calendar components, exactly as they were written. */
export interface CalendarParts {
  /** The year as written, from 0 to 9999. Never remapped. */
  readonly year: number;
  /** 1 to 12. */
  readonly month: number;
  /** 1 to 31. */
  readonly day: number;
  readonly hour: number;
  readonly minute: number;
  readonly second: number;
  readonly millisecond: number;
}

const MILLISECONDS_PER_DAY = 86_400_000;

/** Days in a month of the proleptic Gregorian calendar. */
const MONTH_LENGTHS = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31] as const;

/**
 * Whether every component is inside the range its field allows.
 *
 * The day is checked against 31 rather than against the month, because a serial
 * decoded from a workbook may legitimately name 29 February 1900: the 1900 date
 * system deliberately reproduces a spreadsheet-era bug in which that day
 * exists, and refusing it would refuse a day Excel itself shows. Text that
 * claims to be ISO 8601 is held to the calendar as well, through
 * `isRealCalendarDay`.
 *
 * Hour 24 and second 60 are out of range here. ISO 8601 allows both, as the end
 * of a day and as a leap second, and a worksheet writes neither.
 */
export function isComponentsInRange(parts: CalendarParts): boolean {
  return (
    Number.isInteger(parts.year) &&
    parts.year >= 0 &&
    parts.year <= 9999 &&
    Number.isInteger(parts.month) &&
    parts.month >= 1 &&
    parts.month <= 12 &&
    Number.isInteger(parts.day) &&
    parts.day >= 1 &&
    parts.day <= 31 &&
    Number.isInteger(parts.hour) &&
    parts.hour >= 0 &&
    parts.hour <= 23 &&
    Number.isInteger(parts.minute) &&
    parts.minute >= 0 &&
    parts.minute <= 59 &&
    Number.isInteger(parts.second) &&
    parts.second >= 0 &&
    parts.second <= 59 &&
    Number.isInteger(parts.millisecond) &&
    parts.millisecond >= 0 &&
    parts.millisecond <= 999
  );
}

/** Whether that day exists in that month of that year. */
export function isRealCalendarDay(
  year: number,
  month: number,
  day: number,
): boolean {
  if (month < 1 || month > 12 || day < 1) {
    return false;
  }
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const length =
    month === 2 && leap ? 29 : (MONTH_LENGTHS[month - 1] as number);
  return day <= length;
}

/**
 * Days from 1 January 1970 to the given civil date, by Howard Hinnant's
 * `days_from_civil`. Exact for every year the components allow, and free of the
 * remapping a date constructor would apply.
 */
function daysFromCivil(year: number, month: number, day: number): number {
  const shifted = month <= 2 ? year - 1 : year;
  const era = Math.floor(shifted / 400);
  const yearOfEra = shifted - era * 400;
  const dayOfYear =
    Math.floor((153 * (month + (month > 2 ? -3 : 9)) + 2) / 5) + day - 1;
  const dayOfEra =
    yearOfEra * 365 +
    Math.floor(yearOfEra / 4) -
    Math.floor(yearOfEra / 100) +
    dayOfYear;
  return era * 146_097 + dayOfEra - 719_468;
}

/**
 * The moment these components name, as milliseconds from the epoch, read as
 * UTC. `offsetMinutes` is what the text declared, if it declared one; the
 * moment moves back by it, because a written offset says how far ahead of UTC
 * the clock that wrote it was.
 */
export function calendarEpochMs(
  parts: CalendarParts,
  offsetMinutes = 0,
): number {
  return (
    daysFromCivil(parts.year, parts.month, parts.day) * MILLISECONDS_PER_DAY +
    (parts.hour * 3600 + parts.minute * 60 + parts.second) * 1000 +
    parts.millisecond -
    offsetMinutes * 60_000
  );
}

/**
 * The components a moment wears on its UTC face.
 *
 * That is the face this package means whenever it holds a `Date` for a
 * workbook value: a moment composed from a workbook's epoch and whole days
 * carries the calendar date there, and the local face is that same instant
 * shifted by wherever the reader is sitting. Written once, so no caller has to
 * remember which of the two faces it wanted.
 */
export function utcCalendarParts(value: Date): CalendarParts {
  return {
    year: value.getUTCFullYear(),
    month: value.getUTCMonth() + 1,
    day: value.getUTCDate(),
    hour: value.getUTCHours(),
    minute: value.getUTCMinutes(),
    second: value.getUTCSeconds(),
    millisecond: value.getUTCMilliseconds(),
  };
}

function pad(value: number, width: number): string {
  return String(value).padStart(width, "0");
}

/**
 * The one spelling: the full ISO 8601 timestamp, whether or not the moment
 * carries a time.
 *
 * A date column then reads the same way down its whole length rather than
 * changing shape at the first cell that happens to carry an hour, and a value
 * a workbook holds as a date stays distinguishable from text somebody typed,
 * which a bare `2024-01-01` would not be. `@consultchimps/db` accepts this
 * spelling in a `date` column.
 */
export function calendarIsoText(parts: CalendarParts): string {
  return `${pad(parts.year, 4)}-${pad(parts.month, 2)}-${pad(parts.day, 2)}T${pad(
    parts.hour,
    2,
  )}:${pad(parts.minute, 2)}:${pad(parts.second, 2)}.${pad(
    parts.millisecond,
    3,
  )}Z`;
}
