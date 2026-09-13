import { databaseError } from "../errors.js";
import type { ImportCell } from "./types.js";

const NUMERIC_SERIAL = /^-?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/u;
const WORKSHEET_DATE =
  /^(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})(?:[Tt ](?<hour>\d{2}):(?<minute>\d{2})(?::(?<second>\d{2})(?:\.(?<fraction>\d+))?)?)?(?:(?<zulu>[Zz])|(?<offsetSign>[+-])(?<offsetHour>\d{2}):(?<offsetMinute>\d{2}))?$/u;
const IMPORT_ISO =
  /^(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})(?:T(?<hour>\d{2}):(?<minute>\d{2}):(?<second>\d{2})(?:\.(?<fraction>\d+))?(?:(?<zulu>Z)|(?<offsetSign>[+-])(?<offsetHour>\d{2}):(?<offsetMinute>\d{2}))?)?$/u;

interface ParsedDate {
  readonly dateOnly: boolean;
  readonly epochMilliseconds: number;
  readonly canonical: string;
}

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

function realDay(year: number, month: number, day: number): boolean {
  if (month < 1 || month > 12 || day < 1) return false;
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
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
  return day <= monthDays[month - 1]!;
}

function milliseconds(fraction: string): number | undefined {
  if (fraction.length > 3 && /[1-9]/u.test(fraction.slice(3))) {
    return undefined;
  }
  return Number(fraction.slice(0, 3).padEnd(3, "0"));
}

function parsedDate(
  matched: RegExpExecArray,
  allowNonZeroOffset: boolean,
): ParsedDate | undefined {
  const parts = matched.groups as Record<string, string | undefined>;
  const year = Number(parts["year"]);
  const month = Number(parts["month"]);
  const day = Number(parts["day"]);
  if (!realDay(year, month, day)) return undefined;
  const dateOnly = parts["hour"] === undefined;
  const hour = Number(parts["hour"] ?? "0");
  const minute = Number(parts["minute"] ?? "0");
  const second = Number(parts["second"] ?? "0");
  const millisecond = milliseconds(parts["fraction"] ?? "");
  if (hour > 23 || minute > 59 || second > 59 || millisecond === undefined) {
    return undefined;
  }
  const offsetHour = Number(parts["offsetHour"] ?? "0");
  const offsetMinute = Number(parts["offsetMinute"] ?? "0");
  if (
    offsetHour > 23 ||
    offsetMinute > 59 ||
    (!allowNonZeroOffset && (offsetHour !== 0 || offsetMinute !== 0))
  ) {
    return undefined;
  }
  const offset =
    (parts["offsetSign"] === "-" ? -1 : 1) * (offsetHour * 60 + offsetMinute);
  const epochMilliseconds =
    daysFromCivil(year, month, day) * 86_400_000 +
    hour * 3_600_000 +
    minute * 60_000 +
    second * 1_000 +
    millisecond -
    offset * 60_000;
  const normalized = new Date(epochMilliseconds);
  const normalizedYear = normalized.getUTCFullYear();
  if (normalizedYear < 0 || normalizedYear > 9_999) return undefined;
  return {
    dateOnly,
    epochMilliseconds,
    canonical: dateOnly ? matched[0]! : normalized.toISOString(),
  };
}

function normalizedDateCell(
  cell: Extract<ImportCell, { kind: "date" }>,
): string | undefined {
  const isoMatch = IMPORT_ISO.exec(cell.iso);
  if (isoMatch === null) return undefined;
  const iso = parsedDate(isoMatch, false);
  if (iso === undefined) return undefined;

  if (NUMERIC_SERIAL.test(cell.raw)) {
    return Number.isFinite(Number(cell.raw)) ? iso.canonical : undefined;
  }
  const rawMatch = WORKSHEET_DATE.exec(cell.raw.trim());
  if (rawMatch === null) return undefined;
  const raw = parsedDate(rawMatch, true);
  if (raw === undefined || raw.epochMilliseconds !== iso.epochMilliseconds) {
    return undefined;
  }
  return raw.dateOnly || !iso.dateOnly ? iso.canonical : undefined;
}

export function validatedImportDateIso(
  input: Extract<ImportCell, { kind: "date" }>,
  owner: "source" | "prepared",
): string;
export function validatedImportDateIso(
  input: ImportCell,
  owner: "source" | "prepared",
): string | undefined;
export function validatedImportDateIso(
  input: ImportCell,
  owner: "source" | "prepared",
): string | undefined {
  const cell =
    input.kind === "formula" && input.cached.kind !== "missing"
      ? input.cached
      : input;
  if (cell.kind !== "date") return undefined;
  const normalized = normalizedDateCell(cell);
  if (normalized !== undefined) return normalized;
  if (owner === "prepared") {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import plan contains an invalid captured date cell.",
    );
  }
  throw databaseError(
    "DB_INVALID_SOURCE_DATE",
    "A source date cell is invalid. Re-read the source and prepare the import again.",
  );
}

export function assertValidImportDateCell(
  input: ImportCell,
  owner: "source" | "prepared",
): void {
  validatedImportDateIso(input, owner);
}
