import { readFileSync, existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import type { PbiColumnType, PbiValue } from "../src/model.js";

/**
 * Test-only projection of a decoded value into the pbixray oracle's canonical
 * token grammar. The library never produces these tokens: they exist so a cell
 * can be compared with an independent reader's dump without either side
 * depending on JavaScript coercion.
 */

export const ROOT: string = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);
export const FIXTURES: string = path.join(ROOT, "fixtures");
export const FETCHED: string = path.join(FIXTURES, "fetched");

export interface OracleTable {
  readonly name: string;
  readonly rows: number;
  readonly columns: readonly string[];
  readonly digest: string;
}

export interface OracleFixture {
  readonly name: string;
  readonly fileName: string;
  readonly bytes: number;
  readonly sha256: string;
  readonly url: string;
  readonly tables: readonly OracleTable[];
}

export interface OracleIndex {
  readonly commit: string;
  readonly baseUrl: string;
  readonly fixtures: readonly OracleFixture[];
}

export const COMMITTED: ReadonlySet<string> = new Set([
  "a-2018-fuzzy",
  "b-2018-profiling",
]);

export function oracleIndex(): OracleIndex {
  return JSON.parse(
    readFileSync(path.join(FIXTURES, "oracle-digests.json"), "utf8"),
  ) as OracleIndex;
}

/** The committed file, the fetched cache, or null when the fixture is absent. */
export function fixturePath(fixture: OracleFixture): string | null {
  const committed = path.join(FIXTURES, fixture.fileName);
  if (existsSync(committed)) return committed;
  const fetched = path.join(FETCHED, fixture.fileName);
  return existsSync(fetched) ? fetched : null;
}

export function readFixture(file: string): Uint8Array {
  return new Uint8Array(readFileSync(file));
}

const NULL_TOKEN = "~N~";
const EPOCH_OFFSET_DAYS = 25_569;
const NS_PER_DAY = 86_400_000_000_000n;

function formatFloat(value: number): string {
  if (Number.isNaN(value)) return NULL_TOKEN;
  if (Number.isInteger(value) && Math.abs(value) < 1e15)
    return `F:${String(value)}`;
  // Python repr and JavaScript String agree on the shortest round-trip digits;
  // only the exponent spelling differs (1e-07 against 1e-7).
  let text = String(value);
  const short = /^(-?[\d.]+)e([+-])(\d)$/.exec(text);
  if (short) text = `${short[1]}e${short[2]}0${short[3]}`;
  else if (/e[+-]?\d/.test(text) && !/e[+-]/.test(text))
    text = text.replace("e", "e+");
  return `F:${text}`;
}

/** numpy round-half-to-even at `decimals` places. */
function npRound(value: number, decimals: number): number {
  const scale = Math.pow(10, decimals);
  const scaled = value * scale;
  const floor = Math.floor(scaled);
  const difference = scaled - floor;
  const rounded =
    difference > 0.5
      ? floor + 1
      : difference < 0.5
        ? floor
        : floor % 2 === 0
          ? floor
          : floor + 1;
  return rounded / scale;
}

/**
 * The oracle prints dates as ISO strings through a pandas conversion while the
 * library keeps the day serial, so the projection converts at compare time.
 */
function daysToIso(days: number): string {
  const multiplier = Number(NS_PER_DAY);
  const shifted = days - EPOCH_OFFSET_DAYS;
  const base = Math.trunc(shifted);
  const fraction = npRound(shifted - base, 13);
  const nanoseconds =
    BigInt(base) * NS_PER_DAY + BigInt(Math.trunc(fraction * multiplier));
  const perSecond = 1_000_000_000n;
  let seconds = nanoseconds / perSecond;
  let remainder = nanoseconds % perSecond;
  if (remainder < 0n) {
    remainder += perSecond;
    seconds -= 1n;
  }
  const date = new Date(Number(seconds) * 1000);
  const pad = (value: number, width = 2): string =>
    String(value).padStart(width, "0");
  let text =
    `${pad(date.getUTCFullYear(), 4)}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}` +
    `T${pad(date.getUTCHours())}:${pad(date.getUTCMinutes())}:${pad(date.getUTCSeconds())}`;
  if (remainder !== 0n) text += `.${String(remainder).padStart(9, "0")}`;
  return `T:${text}`;
}

/** The oracle prints currency through Decimal.normalize, from the same integer. */
function currencyToken(scaled: bigint): string {
  const negative = scaled < 0n;
  const absolute = negative ? -scaled : scaled;
  const whole = absolute / 10_000n;
  const fraction = (absolute % 10_000n)
    .toString()
    .padStart(4, "0")
    .replace(/0+$/, "");
  if (whole === 0n && fraction === "") return `D:${negative ? "-0" : "0"}`;
  return `D:${negative ? "-" : ""}${whole.toString()}${fraction ? `.${fraction}` : ""}`;
}

export function canonical(value: PbiValue, type: PbiColumnType): string {
  if (value === null) return NULL_TOKEN;
  switch (type) {
    case "string":
      return `S:${value as string}`;
    case "int64":
      return `I:${(value as bigint).toString()}`;
    case "double":
      return formatFloat(value as number);
    case "dateTimeSerial":
      return daysToIso(value as number);
    case "currency":
      return currencyToken(value as bigint);
    case "boolean":
      return `B:${value ? "1" : "0"}`;
    case "binary":
      return `X:${[...(value as Uint8Array)]
        .map((byte) => byte.toString(16).padStart(2, "0"))
        .join("")}`;
  }
}

/** The escaping the oracle dumps apply before writing a TSV cell. */
export function escapeToken(token: string): string {
  return token
    .replace(/\t/g, "<TAB>")
    .replace(/\n/g, "<LF>")
    .replace(/\r/g, "<CR>");
}

/** FNV-1a over the canonical cells, one stable digest per table. */
export function tableDigest(
  columns: readonly string[],
  rows: readonly (readonly string[])[],
): string {
  let hash = 0x811c9dc5;
  const bump = (text: string): void => {
    for (let index = 0; index < text.length; index++) {
      hash ^= text.charCodeAt(index) & 0xff;
      hash = Math.imul(hash, 0x01000193) >>> 0;
      hash ^= text.charCodeAt(index) >>> 8;
      hash = Math.imul(hash, 0x01000193) >>> 0;
    }
    hash ^= 0x1f;
    hash = Math.imul(hash, 0x01000193) >>> 0;
  };
  for (const column of columns) bump(column);
  for (const row of rows) for (const cell of row) bump(cell);
  return (hash >>> 0).toString(16).padStart(8, "0");
}

const safeTableFile = (name: string): string =>
  [...name].map((ch) => (/[\p{L}\p{N}._\- ]/u.test(ch) ? ch : "_")).join("");

/** The committed cell-for-cell dump of one table, or null when not committed. */
export function oracleRows(
  fixture: string,
  table: string,
): { columns: string[]; rows: string[][] } | null {
  const file = path.join(
    FIXTURES,
    "oracle",
    fixture,
    `T_${safeTableFile(table)}.tsv`,
  );
  if (!existsSync(file)) return null;
  const lines = readFileSync(file, "utf8").split("\n");
  return {
    columns: lines[0]!.split("\t"),
    rows: lines
      .slice(1)
      .filter((line) => line.length > 0)
      .map((line) => line.split("\t")),
  };
}
