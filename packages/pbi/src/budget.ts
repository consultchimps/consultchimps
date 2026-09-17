import { ConsultChimpsError } from "@consultchimps/core";
import type { SqliteReadRuntimeConfig } from "@consultchimps/db/sqlite-read";
import type { PbiStage } from "./errors.js";
import type { Xpress9RuntimeConfig } from "./xpress9/runtime.js";

/** Per-invocation runtime configuration. A cached binary is never reused across configs. */
export interface PbiRuntimeConfig {
  readonly sql?: SqliteReadRuntimeConfig;
  readonly xpress9?: Xpress9RuntimeConfig;
}

/** Limits for the container reader, not the whole export pipeline. */
export interface PbiContainerOptions {
  inputBytes?: number;
  decodedBytes?: number;
  peakBytes?: number;
}

/** Everything readPbiTables accepts. */
export interface PbiReadOptions {
  readonly inputBytes?: number;
  readonly decodedBytes?: number;
  readonly peakBytes?: number;
  readonly includeHiddenTables?: boolean;
  readonly runtime?: PbiRuntimeConfig;
}

/** Everything exportPbiTables accepts. */
export interface PbiExportOptions extends PbiReadOptions {
  readonly outputBytes?: number;
  readonly outputName?: string;
}

export interface ContainerLimits {
  inputBytes: number;
  decodedBytes: number;
  peakBytes: number;
}

export interface ExportLimits extends ContainerLimits {
  outputBytes: number;
  includeHiddenTables: boolean;
  outputName: string;
}

// Conservative reader ceilings. Browser export defaults need end-to-end browser
// measurements, including the runtimes and workbook writer, before shipping.
const defaults: ContainerLimits = {
  inputBytes: 64 * 1024 * 1024,
  decodedBytes: 256 * 1024 * 1024,
  peakBytes: 768 * 1024 * 1024,
};

// Measured, not guessed: a 1,048,575-row five-column worksheet, the most a
// single part can hold, wrote 31,935,638 bytes through the shipped jszip
// settings. 256 MiB leaves room for a model of many such parts plus its
// manifest while still refusing a runaway export.
const outputBytesDefault = 256 * 1024 * 1024;

const capacityRequirement = "a positive safe-integer byte count";
const runtimeRequirement =
  "exactly one locator function or nonempty Uint8Array";

interface InvalidOption {
  path: string;
  requirement: string;
}

function invalidOptionsError(
  invalidOptions: readonly InvalidOption[],
): ConsultChimpsError {
  return new ConsultChimpsError(
    "PBI_INVALID_OPTIONS",
    "One or more options are invalid. Omit each option or supply the required shape.",
    { details: { stage: "options", invalidOptions } },
  );
}

function capacityProblem(value: number | undefined): boolean {
  return value !== undefined && (!Number.isSafeInteger(value) || value <= 0);
}

export function validateLimits(options: PbiContainerOptions): ContainerLimits {
  const paths = ["inputBytes", "decodedBytes", "peakBytes"] as const;
  const invalidOptions = paths
    .filter((path) => capacityProblem(options[path]))
    .map((path) => ({ path, requirement: capacityRequirement }));
  if (invalidOptions.length > 0) {
    throw new ConsultChimpsError(
      "PBI_INVALID_OPTIONS",
      "One or more byte limits are invalid. Omit each option or supply a positive safe-integer byte count.",
      { details: { stage: "options", invalidOptions } },
    );
  }
  return {
    inputBytes: options.inputBytes ?? defaults.inputBytes,
    decodedBytes: options.decodedBytes ?? defaults.decodedBytes,
    peakBytes: options.peakBytes ?? defaults.peakBytes,
  };
}

/**
 * Exactly one source is required for an explicitly supplied runtime config: a
 * locator function, or nonempty bytes. Neither is valid only where a host
 * default exists, which this package resolves in Node.
 */
function runtimeSourceValid(source: unknown): boolean {
  if (source === null || typeof source !== "object" || Array.isArray(source))
    return false;
  const config = source as Xpress9RuntimeConfig;
  const hasLocator = config.locateFile !== undefined;
  const hasBinary = config.wasmBinary !== undefined;
  // Exactly one source. An empty object is a caller who meant to configure a
  // runtime and did not, which both runtimes refuse rather than silently
  // falling back to a host default the caller may not have.
  if (hasLocator === hasBinary) return false;
  if (hasLocator) return typeof config.locateFile === "function";
  return (
    config.wasmBinary instanceof Uint8Array && config.wasmBinary.byteLength > 0
  );
}

/**
 * The full nine-path validation of ADR Decision 4, run before any input read,
 * locator call, or allocation. At most one detail per path, in the fixed order,
 * and no supplied value is ever echoed back.
 */
export function validateExportOptions(
  options: PbiExportOptions,
  includeOutput: boolean,
): ExportLimits {
  if (options === null || typeof options !== "object" || Array.isArray(options))
    throw invalidOptionsError([
      { path: "options", requirement: "an options object, or omitted" },
    ]);
  const invalidOptions: InvalidOption[] = [];
  const capacityPaths = includeOutput
    ? (["inputBytes", "decodedBytes", "outputBytes", "peakBytes"] as const)
    : (["inputBytes", "decodedBytes", "peakBytes"] as const);
  for (const path of capacityPaths) {
    if (capacityProblem(options[path as keyof ExportLimits] as number))
      invalidOptions.push({ path, requirement: capacityRequirement });
  }
  if (
    options.includeHiddenTables !== undefined &&
    typeof options.includeHiddenTables !== "boolean"
  )
    invalidOptions.push({
      path: "includeHiddenTables",
      requirement: "true or false",
    });
  if (
    includeOutput &&
    options.outputName !== undefined &&
    typeof options.outputName !== "string"
  )
    invalidOptions.push({ path: "outputName", requirement: "a string" });
  const runtime = options.runtime;
  if (runtime !== undefined) {
    if (
      runtime === null ||
      typeof runtime !== "object" ||
      Array.isArray(runtime)
    )
      // A malformed parent reports only the parent; its children are not read.
      invalidOptions.push({
        path: "runtime",
        requirement: "an object with optional sql and xpress9 configurations",
      });
    else {
      if (runtime.sql !== undefined && !runtimeSourceValid(runtime.sql))
        invalidOptions.push({
          path: "runtime.sql",
          requirement: runtimeRequirement,
        });
      if (runtime.xpress9 !== undefined && !runtimeSourceValid(runtime.xpress9))
        invalidOptions.push({
          path: "runtime.xpress9",
          requirement: runtimeRequirement,
        });
    }
  }
  if (invalidOptions.length > 0) throw invalidOptionsError(invalidOptions);
  return {
    inputBytes: options.inputBytes ?? defaults.inputBytes,
    decodedBytes: options.decodedBytes ?? defaults.decodedBytes,
    peakBytes: options.peakBytes ?? defaults.peakBytes,
    outputBytes: options.outputBytes ?? outputBytesDefault,
    includeHiddenTables: options.includeHiddenTables ?? false,
    outputName: options.outputName ?? "power-bi-tables.xlsx",
  };
}

/** The one refusal both the workbook writer and the pipeline raise for outputBytes. */
export function outputLimitExceeded(
  limit: number,
  required: number,
): ConsultChimpsError {
  return new ConsultChimpsError(
    "PBI_EXPORT_LIMIT_EXCEEDED",
    "The workbook and its manifest would exceed the configured output byte limit. Export a smaller model, or explicitly increase the limit if your environment can support it.",
    {
      details: { stage: "workbook", option: "outputBytes", limit, required },
    },
  );
}

export function checkLimit(
  limits: ContainerLimits,
  option: keyof ContainerLimits | "outputBytes",
  required: number,
  stage: PbiStage,
): void {
  const limit = (limits as ExportLimits)[option];
  if (!Number.isSafeInteger(required) || required > limit) {
    throw new ConsultChimpsError(
      "PBI_EXPORT_LIMIT_EXCEEDED",
      "Reading the Power BI model would exceed a configured byte limit. Use a smaller model, or explicitly increase the limit if your environment can support it.",
      { details: { stage, option, limit, required } },
    );
  }
}

/**
 * A running bound over the whole pipeline, checked before each allocation and
 * never by catching an out-of-memory. Terms are added and released explicitly
 * so the outcome can report the accounting that produced the peak.
 */
export class PipelineBudget {
  readonly limits: ExportLimits;
  #live = 0;
  #peak = 0;
  #decoded = 0;
  readonly #terms = new Map<string, number>();

  constructor(limits: ExportLimits) {
    this.limits = limits;
  }

  /** Reserve bytes for a named term, refusing before the allocation happens. */
  reserve(term: string, bytes: number, stage: PbiStage): void {
    const next = this.#live + bytes;
    checkLimit(this.limits, "peakBytes", next, stage);
    this.#live = next;
    this.#peak = Math.max(this.#peak, next);
    // Summed, not maximised: a term the pipeline grows in steps, such as the
    // decoder's two buffers, is only honest when the accounting shows what it
    // reached in total rather than its largest single step.
    this.#terms.set(term, (this.#terms.get(term) ?? 0) + bytes);
  }

  release(bytes: number): void {
    this.#live = Math.max(0, this.#live - bytes);
  }

  /** Every buffer decompression produces counts here, released or not. */
  decode(bytes: number, stage: PbiStage): void {
    this.#decoded += bytes;
    checkLimit(this.limits, "decodedBytes", this.#decoded, stage);
  }

  get peakBytes(): number {
    return this.#peak;
  }

  get decodedBytes(): number {
    return this.#decoded;
  }

  /** The named accounting behind the peak, for the browser release to calibrate. */
  terms(): Record<string, number> {
    return Object.fromEntries([...this.#terms.entries()].sort());
  }
}

export function unboundedContainer(): ConsultChimpsError {
  // Refused as a limit because inflation or ZIP64 sizes cannot be bounded
  // before allocation, not because any configured byte limit was exceeded.
  return new ConsultChimpsError(
    "PBI_EXPORT_LIMIT_EXCEEDED",
    "This Power BI file uses a ZIP layout the reader cannot bound: the model part must be stored without ZIP compression in a single-volume, non-ZIP64 archive. Save the .pbix again in Power BI Desktop and try again.",
    { details: { stage: "container", reason: "unsupported-zip-layout" } },
  );
}
