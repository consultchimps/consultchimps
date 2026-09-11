/**
 * What the fill handle writes: one line of source values, extended.
 *
 * Tabulator has no fill handle at all, so the whole gesture is ours, and this
 * is the half that decides the values. It works on text, because the text is
 * what a fill sends: the same canonical text the clipboard carries
 * (`workspace-tsv`), handed to the library's one conversion point, which is the
 * only thing that decides what a column will hold. So this module never asks
 * what type a column is, and it cannot disagree with the database about one.
 *
 * A fill runs one line at a time: a column of the source for a vertical fill, a
 * row of it for a horizontal one. The source occupies indices 0 to n-1 along
 * that line in grid order, and a target is asked for by its index, negative for
 * the cells before the source and n or greater for the cells after it. One
 * index means one rule extrapolated in both directions, rather than two
 * directions with two chances to disagree.
 *
 * The rules, tried in this order, each stated once:
 *
 * - A line the caller marks `copyOnly`, or one holding an empty cell, copies.
 *   A foreign key is the `copyOnly` case: its values are Record IDs, and
 *   reading the trailing integer as a series would point rows at records nobody
 *   chose, which may well exist.
 * - **Numbers**: a single number copies. Two or more with a constant difference
 *   extend linearly. The arithmetic is exact integer arithmetic on the digits,
 *   scaled to the most decimal places the source uses, so 0.1 and 0.2 extend to
 *   0.3 rather than to 0.30000000000000004, a long integer keeps every digit,
 *   and the number of decimals the source was written with is the number the
 *   fill writes. A number written in exponent notation copies.
 * - **Dates**: the value is a date the `date` column's grammar accepts, which
 *   is an ISO date or an ISO timestamp, and a series applies when every source
 *   value carries the same time part, all-midnight import timestamps included.
 *   The date part steps and the time part and the source's spelling are
 *   preserved exactly. A single date steps by one day, as a spreadsheet does.
 *   Two or more step by a constant month difference when every source date
 *   shares a day of the month of 28 or lower (so every month it lands on has
 *   that day and there is no clamping to invent), and otherwise by a constant
 *   day difference. Values with differing time parts copy. Every produced value
 *   is checked against the same grammar the column keeps, and a line that would
 *   produce one the column refuses copies instead, so a fill cannot write a
 *   value the database will only reject.
 * - **Text with a trailing integer**: the integer steps. A single value steps
 *   by one; two or more sharing a prefix step by their constant difference.
 *   Zero padding is kept, to the width of the source value the fill continues
 *   from, and digits grow past it rather than being truncated.
 * - **Everything else copies** the source block cyclically, which is what a
 *   spreadsheet does with a pattern it cannot read: mixed kinds, plain text,
 *   and booleans, which have no series to infer and would otherwise be
 *   alternated into data nobody entered.
 *
 * Every rule is a pure function of the source text and the index, with no
 * locale, no time zone (dates are counted in UTC), and no floating point, so
 * the same drag produces the same values everywhere.
 */
import { isIsoDateText } from "@consultchimps/db";

/** A plain decimal, which is the only numeric spelling a series is read from. */
const DECIMAL = /^([+-]?)(\d+)(?:\.(\d+))?$/u;

/** A date or timestamp, split into the date part and everything after it. */
const DATE_PREFIX = /^(\d{4})-(\d{2})-(\d{2})(.*)$/su;

/** Text ending in digits, split into what comes before them and the digits. */
const TRAILING_INTEGER = /^([\s\S]*?)(\d+)$/u;

/**
 * A number written with an exponent, which the trailing-integer rule must never
 * touch: the digits on the end of "1e+21" are its magnitude, and stepping them
 * would multiply the value by ten.
 */
const EXPONENT_NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)[eE][+-]?\d+$/u;

const MILLISECONDS_PER_DAY = 86_400_000;

export interface FillLineOptions {
  /**
   * Copy the source rather than reading a series from it, whatever it looks
   * like. Set for a column whose values name something rather than counting it.
   */
  readonly copyOnly?: boolean;
}

/** How one target index is answered, once a rule has been read from the source. */
type LineRule = (index: number) => string | null;

/**
 * Extend a line of source values to the given target indices.
 *
 * Indices are relative to the source's own start: 0 to n-1 are the source, a
 * negative index is a cell before it, and n or greater is a cell after it. The
 * answers come back in the order the indices were asked for.
 */
export function fillLine(
  source: readonly string[],
  indices: readonly number[],
  options: FillLineOptions = {},
): string[] {
  if (source.length === 0) {
    return indices.map(() => "");
  }
  const rule = options.copyOnly === true ? null : inferRule(source);
  const copy = cyclicRule(source);
  if (rule === null) {
    return indices.map((index) => copy(index) as string);
  }
  // A rule answers null for a value it cannot spell (a year past the grammar's
  // four digits). One such value condemns the line rather than the cell,
  // because a line half extended and half copied is a pattern nobody asked for.
  const values = indices.map(rule);
  return values.some((value) => value === null)
    ? indices.map((index) => copy(index) as string)
    : (values as string[]);
}

/** The rule the source reads as, or null when nothing but a copy fits. */
function inferRule(source: readonly string[]): LineRule | null {
  if (source.some((value) => value === "")) {
    return null;
  }
  // One number is a value rather than a series, so it copies. Said here because
  // the text rule below would otherwise read its digits as a counter: "5" and
  // "Item 5" look alike to a trailing-integer match, and only one of them is a
  // number.
  const only = source[0];
  if (source.length === 1 && only !== undefined && isPlainNumber(only)) {
    return null;
  }
  return numberRule(source) ?? dateRule(source) ?? textRule(source);
}

/** Repeat the source block, forwards and backwards alike. */
function cyclicRule(source: readonly string[]): LineRule {
  const length = source.length;
  return (index) => source[((index % length) + length) % length] as string;
}

/* -------------------------------------------------------------------------
 * Numbers
 * ---------------------------------------------------------------------- */

interface ScaledNumber {
  /** The value as an integer, scaled by 10 to the power of `scale`. */
  readonly units: bigint;
  /** How many decimal places the value was written with. */
  readonly scale: number;
}

function parseDecimal(value: string): ScaledNumber | null {
  const match = DECIMAL.exec(value);
  if (match === null) {
    return null;
  }
  const [, sign = "", whole = "", fraction = ""] = match;
  const units = BigInt(`${whole}${fraction}`);
  return { units: sign === "-" ? -units : units, scale: fraction.length };
}

function renderDecimal(units: bigint, scale: number): string {
  if (scale === 0) {
    return units.toString();
  }
  const negative = units < 0n;
  const digits = (negative ? -units : units)
    .toString()
    .padStart(scale + 1, "0");
  const whole = digits.slice(0, digits.length - scale);
  const fraction = digits.slice(digits.length - scale);
  return `${negative ? "-" : ""}${whole}.${fraction}`;
}

/**
 * Whether text is a number written the way this module writes one back.
 *
 * The round trip is the test, rather than the grammar alone, because a value
 * whose spelling carries information the arithmetic would drop is not a number
 * here. "007" is the case that matters: read as a number it extends to "8" and
 * loses the padding a text column was deliberately holding, so it belongs to
 * the trailing-integer rule, which keeps it.
 */
function isPlainNumber(value: string): boolean {
  const parsed = parseDecimal(value);
  return parsed !== null && renderDecimal(parsed.units, parsed.scale) === value;
}

function numberRule(source: readonly string[]): LineRule | null {
  // Two or more: a single number is answered in `inferRule`, before the rules
  // that would read its digits as something else.
  if (source.length < 2 || !source.every(isPlainNumber)) {
    return null;
  }
  const parsed: ScaledNumber[] = [];
  for (const value of source) {
    const number = parseDecimal(value);
    if (number === null) {
      return null;
    }
    parsed.push(number);
  }
  // Compared at the most decimal places any of them uses, so 1 and 1.50 are
  // still a step of 0.50 rather than two numbers that cannot be subtracted.
  const scale = Math.max(...parsed.map((number) => number.scale));
  const units = parsed.map((number) => rescale(number, scale));
  const step = (units[1] as bigint) - (units[0] as bigint);
  for (let position = 1; position < units.length; position += 1) {
    if (
      (units[position] as bigint) - (units[position - 1] as bigint) !==
      step
    ) {
      return null;
    }
  }
  const first = units[0] as bigint;
  return (index) => renderDecimal(first + BigInt(index) * step, scale);
}

function rescale(number: ScaledNumber, scale: number): bigint {
  return number.units * 10n ** BigInt(scale - number.scale);
}

/* -------------------------------------------------------------------------
 * Dates
 * ---------------------------------------------------------------------- */

interface DateParts {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  /** Everything after the date, the "T" included, or "" for a plain date. */
  readonly time: string;
}

function parseDate(value: string): DateParts | null {
  const match = DATE_PREFIX.exec(value);
  if (match === null || !isIsoDateText(value)) {
    return null;
  }
  const [, year = "", month = "", day = "", time = ""] = match;
  return {
    year: Number(year),
    month: Number(month),
    day: Number(day),
    time,
  };
}

/** Days since the epoch, in UTC, which is the only clock this reads. */
function dayNumber(date: DateParts): number {
  return Date.UTC(date.year, date.month - 1, date.day) / MILLISECONDS_PER_DAY;
}

/** The month as one number, so a month step is a subtraction. */
function monthNumber(date: DateParts): number {
  return date.year * 12 + (date.month - 1);
}

function renderDate(
  year: number,
  month: number,
  day: number,
  time: string,
): string | null {
  if (year < 0 || year > 9999) {
    return null;
  }
  const text = `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}${time}`;
  // The same grammar the column keeps, so a fill can never produce a value the
  // database would only turn away.
  return isIsoDateText(text) ? text : null;
}

function fromDayNumber(day: number, time: string): string | null {
  const moment = new Date(day * MILLISECONDS_PER_DAY);
  return renderDate(
    moment.getUTCFullYear(),
    moment.getUTCMonth() + 1,
    moment.getUTCDate(),
    time,
  );
}

function dateRule(source: readonly string[]): LineRule | null {
  const parsed: DateParts[] = [];
  for (const value of source) {
    const date = parseDate(value);
    if (date === null) {
      return null;
    }
    parsed.push(date);
  }
  const first = parsed[0] as DateParts;
  // The same instant of the day throughout, or these are not one series. An
  // import writes every date as a timestamp at midnight, so this is the usual
  // case rather than the exotic one.
  if (parsed.some((date) => date.time !== first.time)) {
    return null;
  }
  const { time } = first;

  // A single date steps by a day, the way a spreadsheet fills one.
  if (parsed.length === 1) {
    const start = dayNumber(first);
    return (index) => fromDayNumber(start + index, time);
  }

  const monthStep = constantMonthStep(parsed);
  if (monthStep !== null) {
    const startMonth = monthNumber(first);
    const { day } = first;
    return (index) => {
      const month = startMonth + monthStep * index;
      return renderDate(Math.floor(month / 12), (month % 12) + 1, day, time);
    };
  }

  const days = parsed.map(dayNumber);
  const step = (days[1] as number) - (days[0] as number);
  for (let position = 1; position < days.length; position += 1) {
    if ((days[position] as number) - (days[position - 1] as number) !== step) {
      return null;
    }
  }
  const start = days[0] as number;
  return (index) => fromDayNumber(start + step * index, time);
}

/**
 * The constant month step these dates walk, or null when they do not.
 *
 * Every date has to share a day of the month, and it has to be a day every
 * month has. February is why: stepping the 31st by a month has no answer that
 * is not an invented rule about where it lands, so those dates step by days,
 * which is what they actually state.
 */
function constantMonthStep(dates: readonly DateParts[]): number | null {
  const first = dates[0] as DateParts;
  if (first.day > 28) {
    return null;
  }
  if (dates.some((date) => date.day !== first.day)) {
    return null;
  }
  const months = dates.map(monthNumber);
  const step = (months[1] as number) - (months[0] as number);
  if (step === 0) {
    return null;
  }
  for (let position = 1; position < months.length; position += 1) {
    if (
      (months[position] as number) - (months[position - 1] as number) !==
      step
    ) {
      return null;
    }
  }
  return step;
}

/* -------------------------------------------------------------------------
 * Text with a trailing integer
 * ---------------------------------------------------------------------- */

interface NumberedText {
  readonly prefix: string;
  readonly digits: string;
  readonly value: bigint;
}

function parseNumbered(value: string): NumberedText | null {
  const match = TRAILING_INTEGER.exec(value);
  if (match === null) {
    return null;
  }
  const [, prefix = "", digits = ""] = match;
  return { prefix, digits, value: BigInt(digits) };
}

function textRule(source: readonly string[]): LineRule | null {
  // A number in exponent notation ends in digits and is not text. The number
  // rule has already declined it, so this is where it copies.
  if (source.some((value) => EXPONENT_NUMBER.test(value))) {
    return null;
  }
  const parsed: NumberedText[] = [];
  for (const value of source) {
    const numbered = parseNumbered(value);
    if (numbered === null) {
      return null;
    }
    parsed.push(numbered);
  }
  const first = parsed[0] as NumberedText;
  if (parsed.some((item) => item.prefix !== first.prefix)) {
    return null;
  }

  let step = 1n;
  if (parsed.length > 1) {
    step = (parsed[1] as NumberedText).value - first.value;
    for (let position = 1; position < parsed.length; position += 1) {
      if (
        (parsed[position] as NumberedText).value -
          (parsed[position - 1] as NumberedText).value !==
        step
      ) {
        return null;
      }
    }
  }

  const last = parsed[parsed.length - 1] as NumberedText;
  return (index) => {
    const value = first.value + BigInt(index) * step;
    // Padded to the width of the source value this target continues from, which
    // is the last one going forwards and the first one going backwards.
    const width = (index < 0 ? first.digits : last.digits).length;
    const negative = value < 0n;
    const digits = (negative ? -value : value).toString().padStart(width, "0");
    return `${first.prefix}${negative ? "-" : ""}${digits}`;
  };
}
