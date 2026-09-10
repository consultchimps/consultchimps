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
  /** The year as written, from 0000 to 9999. Never remapped. */
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

/** Milliseconds in a day. The unit every conversion here counts in. */
export const MILLISECONDS_PER_DAY = 86_400_000;

/**
 * How many digits a year is written with, and therefore which years exist here.
 *
 * ISO 8601 writes a year outside four digits only in its expanded form, with a
 * sign and six digits, and nothing downstream reads that form: not the `date`
 * column in `@consultchimps/db`, whose grammar spells a year `\d{4}`, and not
 * a split's output filename. So the range is 0000 to 9999, and it is one number
 * rather than two: the bound below and the width `calendarIsoText` pads to are
 * the same digit count, so a value this module spells is a value that grammar
 * accepts. Both ends are named in a test on either side of that seam, so
 * neither package can move its end of the range quietly.
 */
const YEAR_DIGITS = 4;
const MIN_YEAR = 0;
const MAX_YEAR = 10 ** YEAR_DIGITS - 1;

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
    parts.year >= MIN_YEAR &&
    parts.year <= MAX_YEAR &&
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
function isRealCalendarDay(year: number, month: number, day: number): boolean {
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
function calendarEpochMs(parts: CalendarParts, offsetMinutes = 0): number {
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

/**
 * The moment these components name, once the offset the text wrote is applied,
 * or undefined when they name none.
 *
 * Three things have to hold, and the third is the one that is easy to miss.
 * The components have to be in range; the day has to be one the calendar has;
 * and the *result* has to be in range too. An offset moves the moment, and it
 * can move it out of the years that can be written: `9999-12-31T23:30:00-01:00`
 * is a perfectly good timestamp whose UTC face is the year 10000, which
 * `calendarIsoText` cannot spell and no reader downstream accepts. Judging only
 * what was written would hand the formatter a value it has no spelling for, so
 * the adjusted components go through the same rule.
 *
 * `new Date` is reached only with a number here, which has no interpretation to
 * do; a moment too far out even for that comes back as an invalid date, whose
 * components are not integers and fail the same rule.
 */
export function calendarMoment(
  parts: CalendarParts,
  offsetMinutes: number,
): Date | undefined {
  if (
    !isComponentsInRange(parts) ||
    !isRealCalendarDay(parts.year, parts.month, parts.day)
  ) {
    return undefined;
  }
  const moment = new Date(calendarEpochMs(parts, offsetMinutes));
  return isComponentsInRange(utcCalendarParts(moment)) ? moment : undefined;
}

/** A time of day, and the whole days it carries into. */
export interface TimeOfDay {
  /** Whole days the milliseconds covered, which is 1 when they rounded up. */
  readonly days: number;
  readonly hour: number;
  readonly minute: number;
  readonly second: number;
  readonly millisecond: number;
}

/**
 * Split a whole number of milliseconds into a time of day and the days it
 * carries.
 *
 * Taking a whole number and dividing it is what keeps a carry honest. Rounding
 * a fraction into a millisecond field after the hour, minute and second have
 * already been decided cannot carry: 23:59:59 and a fraction that rounds to a
 * full second has to become 00:00:00 of the next day, and a field that is
 * clamped instead reports the second before, on the day before.
 */
export function timeOfDayFromMilliseconds(milliseconds: number): TimeOfDay {
  const days = Math.floor(milliseconds / MILLISECONDS_PER_DAY);
  const intoDay = milliseconds - days * MILLISECONDS_PER_DAY;
  return {
    days,
    hour: Math.floor(intoDay / 3_600_000),
    minute: Math.floor(intoDay / 60_000) % 60,
    second: Math.floor(intoDay / 1000) % 60,
    millisecond: intoDay % 1000,
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
  return `${pad(parts.year, YEAR_DIGITS)}-${pad(parts.month, 2)}-${pad(parts.day, 2)}T${pad(
    parts.hour,
    2,
  )}:${pad(parts.minute, 2)}:${pad(parts.second, 2)}.${pad(
    parts.millisecond,
    3,
  )}Z`;
}
