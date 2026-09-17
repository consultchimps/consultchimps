import { ConsultChimpsError } from "@consultchimps/core";

/** Refusal codes reserved by ADR 0004 across every reading and export stage. */
export type PbiErrorCode =
  | "PBI_INVALID_OPTIONS"
  | "PBI_RUNTIME_UNAVAILABLE"
  | "PBI_INVALID_CONTAINER"
  | "PBI_MODEL_UNREADABLE"
  | "PBI_NO_MODEL"
  | "PBI_MODEL_ENCRYPTED"
  | "PBI_NO_EXPORTABLE_TABLES"
  | "PBI_EXPORT_LIMIT_EXCEEDED";

/** The fourteen manifest reason codes of ADR 0004 Decision 9. */
export type PbiReasonCode =
  | "PBI_BINARY_AS_BASE64"
  | "PBI_BINARY_CELL_TOO_LONG"
  | "PBI_COLUMN_ROW_ALIGNMENT_UNRECOVERABLE"
  | "PBI_COLUMN_UNREADABLE"
  | "PBI_COLUMN_UNSUPPORTED_ENCODING"
  | "PBI_DATE_AS_TEXT"
  | "PBI_DATE_ROUNDED"
  | "PBI_NONFINITE_AS_TEXT"
  | "PBI_NUMERIC_AS_TEXT"
  | "PBI_TABLE_HIDDEN"
  | "PBI_TABLE_NO_EXPORTABLE_COLUMNS"
  | "PBI_TABLE_SPLIT"
  | "PBI_TABLE_TOO_WIDE"
  | "PBI_TEXT_TRUNCATED";

/**
 * Reference paths the spike corpus never exercised. They attempt the decode and
 * record a manifest warning instead of refusing (contract decision 3).
 */
export type PbiUnverifiedCode =
  | "PBI_UNVERIFIED_BINARY_TYPE"
  | "PBI_UNVERIFIED_BOOLEAN_TYPE"
  | "PBI_UNVERIFIED_COMPRESSION_CLASS"
  | "PBI_UNVERIFIED_MULTIPLE_PARTITIONS"
  | "PBI_UNVERIFIED_MULTIPLE_SEGMENTS"
  | "PBI_UNVERIFIED_NON_LATIN_DICTIONARY"
  | "PBI_UNVERIFIED_XPRESS9_MULTITHREADED";

/** Stages named in refusals. Controlled labels only; never upstream text. */
export type PbiStage =
  | "container"
  | "model-part"
  | "options"
  | "xpress9"
  | "backup"
  | "catalog"
  | "decode"
  | "workbook";

export function invalidContainer(): ConsultChimpsError {
  return new ConsultChimpsError(
    "PBI_INVALID_CONTAINER" satisfies PbiErrorCode,
    "The file is not a supported Power BI ZIP container, or its required parts are damaged or missing. Save a new .pbix in Power BI Desktop and try again.",
    { details: { stage: "container" } },
  );
}

export function unreadableModel(
  stage: PbiStage = "model-part",
): ConsultChimpsError {
  return new ConsultChimpsError(
    "PBI_MODEL_UNREADABLE" satisfies PbiErrorCode,
    "The embedded model part is empty or damaged. Save a new .pbix with imported data in Power BI Desktop and try again.",
    { details: { stage } },
  );
}

/**
 * A cleanup failure never masks a successful export: the host is told to drop
 * the worker rather than reuse a reader that could not release its runtime.
 */
export function cleanupRequired(): ConsultChimpsError {
  return new ConsultChimpsError(
    "PBI_MODEL_UNREADABLE" satisfies PbiErrorCode,
    "The model catalog reader could not be released. Discard the worker or process that ran this export before reading another model.",
    { details: { stage: "catalog", recovery: "discard-worker" } },
  );
}

export function encryptedModel(): ConsultChimpsError {
  return new ConsultChimpsError(
    "PBI_MODEL_ENCRYPTED" satisfies PbiErrorCode,
    "The embedded model is encrypted or password-protected. Ask the file owner for an unencrypted .pbix saved with imported data.",
    { details: { stage: "model-part" } },
  );
}

export type PbiRuntimeLabel = "sqlite" | "xpress9";
export type PbiRuntimeStage = "load" | "compile" | "instantiate";

export function runtimeUnavailable(
  runtime: PbiRuntimeLabel,
  stage: PbiRuntimeStage,
): ConsultChimpsError {
  // No URL, no locator argument, and no upstream message: the caller learns
  // which runtime and which step, and how to supply the asset.
  return new ConsultChimpsError(
    "PBI_RUNTIME_UNAVAILABLE" satisfies PbiErrorCode,
    runtime === "sqlite"
      ? "SQLite could not load its WebAssembly runtime. Install the pinned @sqlite.org/sqlite-wasm peer dependency, or configure the runtime.sql option with a locator or the binary itself."
      : "The XPress9 decoder could not load its WebAssembly runtime. Serve the pinned xpress9.wasm asset, or configure the runtime.xpress9 option with a locator or the binary itself.",
    // The same shape the XPress9 loader emits, so a caller switching on
    // details.stage reads a load stage from both runtimes, not "runtime" from
    // one of them.
    { details: { runtime, stage } },
  );
}

export interface ExclusionCounts {
  readonly code: PbiReasonCode;
  readonly tables: number;
  readonly columns: number;
}

/**
 * No table survived. Counts are anonymous, aggregated by reason code in
 * ascending ASCII order; no name, identifier, or DAX appears anywhere.
 */
export function noExportableTables(
  counts: readonly ExclusionCounts[],
  hiddenExcluded: boolean,
): ConsultChimpsError {
  const summary =
    counts.length === 0
      ? "This model has no tables to export."
      : "Every table in this model was excluded from the export.";
  const cure = hiddenExcluded
    ? " Set includeHiddenTables to true to export the tables the model marks hidden."
    : "";
  return new ConsultChimpsError(
    "PBI_NO_EXPORTABLE_TABLES" satisfies PbiErrorCode,
    `${summary}${cure} No workbook was produced.`,
    { details: { stage: "decode", exclusions: counts } },
  );
}
