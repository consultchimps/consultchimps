import type { ColumnDefinition } from "../schema.js";
import { databaseError } from "../errors.js";
import { normalizeDecimal } from "./decimal.js";
import type { EngineRow, EngineValue } from "./engine.js";

function invalidValue(table: string, column: ColumnDefinition): never {
  throw databaseError(
    "DB_CONVERSION_INVALID_VALUE",
    `The source table "${table}" contains a value in column "${column.name}" that cannot be converted as ${column.type} without changing it. Repair the source value and prepare the conversion again.`,
    { table, column: column.name, type: column.type },
  );
}

function safeInteger(value: EngineValue): bigint | undefined {
  if (typeof value === "bigint") return value;
  if (typeof value === "number" && Number.isSafeInteger(value)) {
    return BigInt(value);
  }
  return undefined;
}

function floorDivide(value: bigint, divisor: bigint): bigint {
  const quotient = value / divisor;
  return value < 0n && value % divisor !== 0n ? quotient - 1n : quotient;
}

function calendarDate(
  displayYearText: string,
  monthText: string,
  dayText: string,
  beforeCommonEra: boolean,
): {
  readonly year: bigint;
  readonly month: number;
  readonly day: number;
} | null {
  if (displayYearText.replace(/^0*/u, "").length > 7) return null;
  const displayYear = BigInt(displayYearText);
  if (beforeCommonEra && displayYear < 1n) return null;
  const year = beforeCommonEra ? 1n - displayYear : displayYear;
  const month = Number(monthText);
  const day = Number(dayText);
  if (month < 1 || month > 12 || day < 1) return null;
  const leap = year % 4n === 0n && (year % 100n !== 0n || year % 400n === 0n);
  const monthDays = [
    31,
    leap ? 29 : 28,
    31,
    30,
    31,
    30,
    31,
    31,
    30,
    31,
    30,
    31,
  ];
  return day <= monthDays[month - 1]! ? { year, month, day } : null;
}

function epochDays(date: {
  readonly year: bigint;
  readonly month: number;
  readonly day: number;
}): bigint {
  const adjustedYear = date.month <= 2 ? date.year - 1n : date.year;
  const era = floorDivide(adjustedYear, 400n);
  const yearOfEra = adjustedYear - era * 400n;
  const shiftedMonth = BigInt(date.month + (date.month > 2 ? -3 : 9));
  const dayOfYear = (153n * shiftedMonth + 2n) / 5n + BigInt(date.day) - 1n;
  const dayOfEra =
    yearOfEra * 365n + yearOfEra / 4n - yearOfEra / 100n + dayOfYear;
  return era * 146_097n + dayOfEra - 719_468n;
}

function validDate(value: string): boolean {
  if (value === "infinity" || value === "-infinity") return true;
  const match = /^(\d{4,})-(\d{2})-(\d{2})(?: (\(BC\)))?$/u.exec(value);
  if (match === null) return false;
  const date = calendarDate(
    match[1]!,
    match[2]!,
    match[3]!,
    match[4] !== undefined,
  );
  if (date === null) return false;
  const days = epochDays(date);
  return days >= -2_147_483_647n && days <= 2_147_483_646n;
}

function validTimestamp(value: string): boolean {
  if (value === "infinity" || value === "-infinity") return true;
  const match =
    /^(\d{4,})-(\d{2})-(\d{2})(?: (\(BC\)))?[ T](\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?(?:Z)?$/u.exec(
      value,
    );
  if (match === null) return false;
  const date = calendarDate(
    match[1]!,
    match[2]!,
    match[3]!,
    match[4] !== undefined,
  );
  if (date === null) return false;
  const hour = Number(match[5]);
  const minute = Number(match[6]);
  const second = Number(match[7]);
  if (hour > 23 || minute > 59 || second > 59) return false;
  const fraction = match[8] ?? "";
  if (fraction.length > 6 && /[1-9]/u.test(fraction.slice(6))) return false;
  const micros = BigInt((fraction.slice(0, 6) + "000000").slice(0, 6));
  const timestamp =
    epochDays(date) * 86_400_000_000n +
    BigInt(hour) * 3_600_000_000n +
    BigInt(minute) * 60_000_000n +
    BigInt(second) * 1_000_000n +
    micros;
  return (
    timestamp >= -9_223_372_036_854_775_807n &&
    timestamp <= 9_223_372_036_854_775_806n
  );
}

function validatedValue(
  value: EngineValue,
  table: string,
  column: ColumnDefinition,
): EngineValue {
  if (value === null) {
    if (column.nullable !== false) return null;
    return invalidValue(table, column);
  }
  switch (column.type) {
    case "text":
      if (typeof value === "string") return value;
      break;
    case "integer": {
      const integer = safeInteger(value);
      if (integer !== undefined) return integer;
      break;
    }
    case "real":
      if (typeof value === "number" && Number.isFinite(value)) return value;
      break;
    case "decimal":
      if (
        typeof value === "string" &&
        column.precision !== undefined &&
        column.scale !== undefined
      ) {
        const decimal = normalizeDecimal(value, column.precision, column.scale);
        if (decimal !== undefined) return decimal;
      }
      break;
    case "boolean":
      if (typeof value === "boolean") return value;
      if (value === 0n || value === 0) return false;
      if (value === 1n || value === 1) return true;
      break;
    case "date":
      if (typeof value === "string" && validDate(value)) return value;
      break;
    case "timestamp":
      if (typeof value === "string" && validTimestamp(value)) return value;
      break;
  }
  return invalidValue(table, column);
}

export function validatedConversionRows(options: {
  readonly table: string;
  readonly columns: readonly ColumnDefinition[];
  readonly rows: readonly EngineRow[];
}): readonly (readonly EngineValue[])[] {
  return options.rows.map((row) =>
    options.columns.map((column) =>
      validatedValue(row[column.name] ?? null, options.table, column),
    ),
  );
}
