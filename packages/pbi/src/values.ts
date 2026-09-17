import type { PbiReasonCode } from "./errors.js";
import type { PbiColumnType, PbiValue } from "./model.js";

/** Excel's own grid limits. They exist nowhere else in the repository. */
export const MAX_WORKSHEET_ROWS = 1_048_576;
export const MAX_DATA_ROWS_PER_PART: number = MAX_WORKSHEET_ROWS - 1;
export const MAX_WORKSHEET_COLUMNS = 16_384;
export const MAX_CELL_UNITS = 32_767;
export const MAX_CELL_LINE_BREAKS = 253;
export const MAX_SHEET_NAME_UNITS = 31;

const MS_PER_DAY = 86_400_000;
const MS_PER_DAY_BIG = 86_400_000n;
/** Day count of 1900-03-01 from the model epoch 1899-12-30. */
const EXCEL_LEAP_BUG_DAYS = 61n;
/** 9999-12-31 as an Excel serial; a candidate at or past the next day is text. */
const MAX_EXCEL_SERIAL = 2_958_466n;
/** Excel's conservative numeric magnitude window. */
const MIN_ABSOLUTE = 2.2251e-308;
const MAX_ABSOLUTE = 9.99999999999999e307;
const MAX_SIGNIFICANT_DIGITS = 15;

/** Style indexes of the workbook's fixed style sheet. */
export const STYLE_GENERAL = 0;
export const STYLE_DATE = 1;
export const STYLE_CURRENCY = 2;

export type Cell =
  | { readonly kind: "blank" }
  | { readonly kind: "number"; readonly text: string; readonly style: number }
  | { readonly kind: "text"; readonly text: string }
  | { readonly kind: "boolean"; readonly value: boolean };

export interface CellResult {
  readonly cell: Cell;
  readonly reason?: PbiReasonCode;
}

const BLANK: CellResult = { cell: { kind: "blank" } };

/** value = mantissa * 2^exponent, exactly, for any finite double. */
function decompose(value: number): { mantissa: bigint; exponent: number } {
  const view = new DataView(new ArrayBuffer(8));
  view.setFloat64(0, value);
  const bits = view.getBigUint64(0);
  const negative = bits >> 63n === 1n;
  const exponentBits = Number((bits >> 52n) & 0x7ffn);
  const fraction = bits & 0xf_ffff_ffff_ffffn;
  const mantissa = exponentBits === 0 ? fraction : fraction | (1n << 52n);
  const exponent = exponentBits === 0 ? -1074 : exponentBits - 1075;
  return { mantissa: negative ? -mantissa : mantissa, exponent };
}

/**
 * The serial's exact rational value times 86,400,000, rounded once to an
 * integer millisecond, an exact half going to the later instant. Never staged
 * through nanoseconds or a host date.
 */
export function exactMilliseconds(serial: number): {
  milliseconds: bigint;
  exact: boolean;
} {
  const { mantissa, exponent } = decompose(serial);
  const numerator = mantissa * MS_PER_DAY_BIG;
  if (exponent >= 0)
    return { milliseconds: numerator << BigInt(exponent), exact: true };
  const denominator = 1n << BigInt(-exponent);
  let quotient = numerator / denominator;
  let remainder = numerator - quotient * denominator;
  if (remainder < 0n) {
    quotient -= 1n;
    remainder += denominator;
  }
  const exact = remainder === 0n;
  if (remainder * 2n >= denominator) quotient += 1n;
  return { milliseconds: quotient, exact };
}

/** Significant decimal digits of a coefficient, leading and trailing zeros excluded. */
export function significantDigits(decimal: string): number {
  const coefficient = decimal.split(/[eE]/)[0] ?? "";
  const digits = coefficient.replace(/^[+-]/, "").replace(".", "");
  const trimmed = digits.replace(/^0+/, "").replace(/0+$/, "");
  return trimmed.length;
}

function inNumericRange(value: number): boolean {
  if (Object.is(value, 0)) return true;
  const magnitude = Math.abs(value);
  return magnitude >= MIN_ABSOLUTE && magnitude <= MAX_ABSOLUTE;
}

/** Sign, integer digits and exactly four fractional digits, no exponent. */
export function currencyText(scaled: bigint): string {
  const negative = scaled < 0n;
  const absolute = negative ? -scaled : scaled;
  const whole = absolute / 10_000n;
  const fraction = (absolute % 10_000n).toString().padStart(4, "0");
  return `${negative ? "-" : ""}${whole.toString()}.${fraction}`;
}

function numberCell(text: string, style: number): CellResult {
  return { cell: { kind: "number", text, style } };
}

function textCell(text: string, reason: PbiReasonCode): CellResult {
  return { cell: { kind: "text", text }, reason };
}

function int64Cell(value: bigint): CellResult {
  const text = value.toString();
  if (significantDigits(text) > MAX_SIGNIFICANT_DIGITS)
    return textCell(text, "PBI_NUMERIC_AS_TEXT");
  const asNumber = Number(value);
  if (!inNumericRange(asNumber) || BigInt(asNumber) !== value)
    return textCell(text, "PBI_NUMERIC_AS_TEXT");
  // The number cell carries the same digits, so the read-back comparison is in
  // the original integer units.
  if (Number(asNumber.toString()) !== asNumber)
    return textCell(text, "PBI_NUMERIC_AS_TEXT");
  return numberCell(asNumber.toString(), STYLE_GENERAL);
}

function currencyCell(scaled: bigint): CellResult {
  const text = currencyText(scaled);
  if (significantDigits(text) > MAX_SIGNIFICANT_DIGITS)
    return textCell(text, "PBI_NUMERIC_AS_TEXT");
  const asNumber = Number(scaled) / 10_000;
  if (!inNumericRange(asNumber) && !Object.is(asNumber, 0))
    return textCell(text, "PBI_NUMERIC_AS_TEXT");
  if (Object.is(asNumber, -0)) return textCell(text, "PBI_NUMERIC_AS_TEXT");
  // Compare in the original scaled integer units after serialization.
  if (!scaledRoundTrips(Number(asNumber.toString()), scaled))
    return textCell(text, "PBI_NUMERIC_AS_TEXT");
  return numberCell(asNumber.toString(), STYLE_CURRENCY);
}

/**
 * True when reading the emitted cell back and returning it to its original
 * integer units recovers the stored count of ten-thousandths. The scaling uses
 * exact rational arithmetic, because a decimal such as 12.3456 is never an
 * exact double and an exact-equality test would send every currency to text.
 */
function scaledRoundTrips(value: number, scaled: bigint): boolean {
  const { mantissa, exponent } = decompose(value);
  const numerator = mantissa * 10_000n;
  if (exponent >= 0) return numerator << BigInt(exponent) === scaled;
  const denominator = 1n << BigInt(-exponent);
  let quotient = numerator / denominator;
  let remainder = numerator - quotient * denominator;
  if (remainder < 0n) {
    quotient -= 1n;
    remainder += denominator;
  }
  if (remainder * 2n >= denominator) quotient += 1n;
  return quotient === scaled;
}

function doubleCell(value: number): CellResult {
  if (Number.isNaN(value)) return textCell("NaN", "PBI_NONFINITE_AS_TEXT");
  if (value === Infinity) return textCell("Infinity", "PBI_NONFINITE_AS_TEXT");
  if (value === -Infinity)
    return textCell("-Infinity", "PBI_NONFINITE_AS_TEXT");
  if (Object.is(value, -0)) return textCell("-0", "PBI_NUMERIC_AS_TEXT");
  const text = value.toString();
  if (significantDigits(text) > MAX_SIGNIFICANT_DIGITS)
    return textCell(text, "PBI_NUMERIC_AS_TEXT");
  if (!inNumericRange(value)) return textCell(text, "PBI_NUMERIC_AS_TEXT");
  if (Number(text) !== value) return textCell(text, "PBI_NUMERIC_AS_TEXT");
  return numberCell(text, STYLE_GENERAL);
}

/** The fallback spelling that reconstructs the same finite double exactly. */
export function dateSerialText(serial: number): string {
  return `pbi-date-serial:${Object.is(serial, -0) ? "-0" : serial.toString()}`;
}

function dateCell(serial: number): CellResult {
  // Values below 2, meaning before 1900-01-01, select text without rounding
  // even where rounding would cross the boundary.
  if (!(serial >= 2))
    return textCell(dateSerialText(serial), "PBI_DATE_AS_TEXT");
  const { milliseconds, exact } = exactMilliseconds(serial);
  const excelMilliseconds =
    milliseconds < EXCEL_LEAP_BUG_DAYS * MS_PER_DAY_BIG
      ? milliseconds - MS_PER_DAY_BIG
      : milliseconds;
  if (
    excelMilliseconds < MS_PER_DAY_BIG ||
    excelMilliseconds >= MAX_EXCEL_SERIAL * MS_PER_DAY_BIG
  )
    return textCell(dateSerialText(serial), "PBI_DATE_AS_TEXT");
  const candidate = Number(excelMilliseconds) / MS_PER_DAY;
  const text = candidate.toString();
  // The cell must read back to the same integer millisecond.
  if (exactMilliseconds(Number(text)).milliseconds !== excelMilliseconds)
    return textCell(dateSerialText(serial), "PBI_DATE_AS_TEXT");
  return exact
    ? numberCell(text, STYLE_DATE)
    : {
        cell: { kind: "number", text, style: STYLE_DATE },
        reason: "PBI_DATE_ROUNDED",
      };
}

/**
 * Truncate to the longest prefix within both worksheet limits, never splitting
 * a surrogate pair or a CRLF and never reaching a 254th line break.
 */
export function truncateCellText(value: string): {
  text: string;
  truncated: boolean;
} {
  if (value.length <= MAX_CELL_UNITS && !/[\r\n]/.test(value))
    return { text: value, truncated: false };
  let index = 0;
  let breaks = 0;
  while (index < value.length) {
    const unit = value.charCodeAt(index);
    let step = 1;
    let isBreak = false;
    if (unit === 13) {
      isBreak = true;
      if (value.charCodeAt(index + 1) === 10) step = 2;
    } else if (unit === 10) isBreak = true;
    else if (unit >= 0xd800 && unit <= 0xdbff) {
      const low = value.charCodeAt(index + 1);
      if (low >= 0xdc00 && low <= 0xdfff) step = 2;
    }
    if (isBreak && breaks + 1 > MAX_CELL_LINE_BREAKS) break;
    if (index + step > MAX_CELL_UNITS) break;
    if (isBreak) breaks++;
    index += step;
  }
  return { text: value.slice(0, index), truncated: index < value.length };
}

const BASE64 =
  "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

/** Standard padded base64 with no line breaks, host-independent. */
export function toBase64(bytes: Uint8Array): string {
  let out = "";
  let index = 0;
  for (; index + 2 < bytes.length; index += 3) {
    const word =
      (bytes[index]! << 16) | (bytes[index + 1]! << 8) | bytes[index + 2]!;
    out +=
      BASE64[(word >> 18) & 63]! +
      BASE64[(word >> 12) & 63]! +
      BASE64[(word >> 6) & 63]! +
      BASE64[word & 63]!;
  }
  const remaining = bytes.length - index;
  if (remaining === 1) {
    const word = bytes[index]! << 16;
    out += `${BASE64[(word >> 18) & 63]!}${BASE64[(word >> 12) & 63]!}==`;
  } else if (remaining === 2) {
    const word = (bytes[index]! << 16) | (bytes[index + 1]! << 8);
    out += `${BASE64[(word >> 18) & 63]!}${BASE64[(word >> 12) & 63]!}${BASE64[(word >> 6) & 63]!}=`;
  }
  return out;
}

/** One typed value to one worksheet cell, with the conversion it needed. */
export function toCell(value: PbiValue, type: PbiColumnType): CellResult {
  if (value === null) return BLANK;
  switch (type) {
    case "string": {
      const { text, truncated } = truncateCellText(value as string);
      return truncated
        ? { cell: { kind: "text", text }, reason: "PBI_TEXT_TRUNCATED" }
        : { cell: { kind: "text", text } };
    }
    case "int64":
      return int64Cell(value as bigint);
    case "currency":
      return currencyCell(value as bigint);
    case "double":
      return doubleCell(value as number);
    case "dateTimeSerial":
      return dateCell(value as number);
    case "boolean":
      return { cell: { kind: "boolean", value: value as boolean } };
    case "binary":
      return {
        cell: { kind: "text", text: toBase64(value as Uint8Array) },
        reason: "PBI_BINARY_AS_BASE64",
      };
  }
}

/** A date column holding a non-null, non-finite serial cannot be read at all. */
export function dateColumnReadable(values: readonly PbiValue[]): boolean {
  for (const value of values) {
    if (value === null) continue;
    if (typeof value !== "number" || !Number.isFinite(value)) return false;
  }
  return true;
}

/** Oversized base64 excludes the whole column; base64 is never truncated. */
export function oversizedBinaryValues(values: readonly PbiValue[]): number {
  let count = 0;
  for (const value of values) {
    if (value === null) continue;
    // Four base64 characters per three bytes, padded.
    if (Math.ceil((value as Uint8Array).length / 3) * 4 > MAX_CELL_UNITS)
      count++;
  }
  return count;
}
