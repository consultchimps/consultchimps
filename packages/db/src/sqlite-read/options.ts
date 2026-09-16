import { ConsultChimpsError } from "@consultchimps/core";

/** Node callers may omit the source to use the installed SQLite WASM asset. */
export interface SqliteReadRuntimeConfig {
  readonly locateFile?: (fileName: string) => string;
  readonly wasmBinary?: Uint8Array;
}

export interface SqliteReadOptions {
  readonly runtime?: SqliteReadRuntimeConfig;
  readonly maxDatabaseBytes?: number;
  readonly maxSqliteBytes?: number;
  readonly maxSqlBytes?: number;
  readonly maxRows?: number;
  readonly maxResultBytes?: number;
  readonly maxSteps?: number;
}

export interface ReadLimits {
  maxDatabaseBytes: number;
  maxSqliteBytes: number;
  maxSqlBytes: number;
  maxRows: number;
  maxResultBytes: number;
  maxSteps: number;
}

const defaults: ReadLimits = {
  maxDatabaseBytes: 32 * 1024 * 1024,
  maxSqliteBytes: 64 * 1024 * 1024,
  maxSqlBytes: 64 * 1024,
  maxRows: 10_000,
  maxResultBytes: 4 * 1024 * 1024,
  maxSteps: 1_000_000,
};

export function readError(
  code: string,
  stage: string,
  message: string,
): ConsultChimpsError {
  // Raw SQLite errors can contain schema names, SQL, or source URLs.
  return new ConsultChimpsError(code, message, { details: { stage } });
}

export function limitError(
  option: keyof ReadLimits,
  limit: number,
): ConsultChimpsError {
  return new ConsultChimpsError(
    "DB_SQLITE_READ_LIMIT_EXCEEDED",
    "The SQLite read exceeded a configured limit. Read a smaller catalog or result, or explicitly increase the limit if your environment supports it.",
    { details: { option, limit } },
  );
}

export function validateReadOptions(options: SqliteReadOptions): ReadLimits {
  if (
    options === null ||
    typeof options !== "object" ||
    Array.isArray(options)
  ) {
    throw readError(
      "DB_SQLITE_READ_INVALID_OPTIONS",
      "options",
      "Supply a SQLite read options object, or omit it to use the defaults.",
    );
  }
  const invalidOptions: { path: string; requirement: string }[] = [];
  const result = { ...defaults };
  for (const path of Object.keys(defaults) as (keyof ReadLimits)[]) {
    const value = options[path];
    if (value === undefined) continue;
    if (!Number.isSafeInteger(value) || value <= 0 || value > 0x7fffffff) {
      invalidOptions.push({
        path,
        requirement: "a positive integer no greater than 2147483647",
      });
    } else result[path] = value;
  }
  if (options.runtime !== undefined) {
    const runtime = options.runtime;
    if (
      runtime === null ||
      typeof runtime !== "object" ||
      Array.isArray(runtime) ||
      !(
        (typeof runtime.locateFile === "function" &&
          runtime.wasmBinary === undefined) ||
        (runtime.locateFile === undefined &&
          runtime.wasmBinary instanceof Uint8Array &&
          runtime.wasmBinary.byteLength > 0)
      )
    )
      invalidOptions.push({
        path: "runtime",
        requirement: "exactly one locator function or nonempty Uint8Array",
      });
  }
  if (invalidOptions.length) {
    throw new ConsultChimpsError(
      "DB_SQLITE_READ_INVALID_OPTIONS",
      "One or more SQLite read options are invalid. Supply the required shapes or omit those options.",
      { details: { invalidOptions } },
    );
  }
  return result;
}

/** Count UTF-8 bytes without allocating a second copy of caller-owned text. */
export function utf8Bytes(value: string): number {
  let size = 0;
  for (const character of value) {
    const point = character.codePointAt(0)!;
    size += point < 0x80 ? 1 : point < 0x800 ? 2 : point < 0x10000 ? 3 : 4;
  }
  return size;
}
