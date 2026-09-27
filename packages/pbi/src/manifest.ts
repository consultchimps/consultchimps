import type { PbiReasonCode, PbiUnverifiedCode } from "./errors.js";
import type { PbiColumnType } from "./model.js";

/**
 * The manifest is bounded by the model's structure, never by its row count: a
 * million high-precision identifiers add one count to their column, not a
 * million entries. Property order is fixed, there is no wall-clock timestamp,
 * no random id, and no copy of any source value.
 */

export interface PbiReasonEntry {
  readonly code: PbiReasonCode;
  readonly count: number;
}

export interface PbiWorksheetPart {
  readonly sheetName: string;
  readonly start: number;
  readonly end: number;
}

export interface PbiManifestColumn {
  readonly name: string;
  readonly id: number;
  readonly type: PbiColumnType;
  readonly dax?: string;
  readonly encoding?: "base64";
  readonly dateEpoch?: "1899-12-30";
  readonly reasons: readonly PbiReasonEntry[];
}

export interface PbiManifestTable {
  readonly name: string;
  readonly id: number;
  readonly dax?: string;
  readonly reasons: readonly PbiReasonEntry[];
  readonly parts: readonly PbiWorksheetPart[];
  readonly columns: readonly PbiManifestColumn[];
}

export interface PbiManifestExclusion {
  readonly name: string;
  readonly id: number;
  readonly reasons: readonly PbiReasonEntry[];
  readonly columns: readonly PbiManifestColumn[];
}

export interface PbiUnverifiedPath {
  readonly code: PbiUnverifiedCode;
  readonly count: number;
}

export interface PbiManifest {
  readonly schemaVersion: 1;
  readonly tables: readonly PbiManifestTable[];
  readonly excludedTables: readonly PbiManifestExclusion[];
  readonly unverifiedPaths: readonly PbiUnverifiedPath[];
}

/** Zero counts are omitted; entries sort by code in ascending ASCII order. */
export function reasonEntries(
  counts: ReadonlyMap<PbiReasonCode, number> | undefined,
): PbiReasonEntry[] {
  if (counts === undefined) return [];
  return [...counts.entries()]
    .filter(([, count]) => count > 0)
    .map(([code, count]) => ({ code, count }))
    .sort((left, right) => (left.code < right.code ? -1 : 1));
}

const REASON_TEXT: Record<PbiReasonCode, (count: number) => string> = {
  PBI_BINARY_AS_BASE64: (count) =>
    `${count} binary value${count === 1 ? " was" : "s were"} written as base64 text. See the manifest for the affected columns.`,
  PBI_BINARY_CELL_TOO_LONG: (count) =>
    `${count} column${count === 1 ? " was" : "s were"} left out because a base64 value is longer than a worksheet cell can hold. See the manifest for which.`,
  PBI_COLUMN_DECODER_ERROR: (count) =>
    `${count} column${count === 1 ? " was" : "s were"} left out because this reader failed while decoding ${count === 1 ? "it" : "them"}. The model itself may be fine. See the manifest for which, and report the problem.`,
  PBI_COLUMN_ROW_ALIGNMENT_UNRECOVERABLE: (count) =>
    `${count} column${count === 1 ? " was" : "s were"} left out because its decoded values could not be placed on the table's rows. No row was moved or dropped. See the manifest for which.`,
  PBI_COLUMN_UNREADABLE: (count) =>
    `${count} column${count === 1 ? " could" : "s could"} not be read as its declared type. Check the source data and save the model again. See the manifest for which.`,
  PBI_COLUMN_UNSUPPORTED_ENCODING: (count) =>
    `${count} column${count === 1 ? " uses" : "s use"} an encoding this reader does not support and ${count === 1 ? "was" : "were"} left out. See the manifest for which.`,
  PBI_DATE_AS_TEXT: (count) =>
    `${count} date value${count === 1 ? " was" : "s were"} written as text because the workbook cannot hold them as dates. See the manifest for the affected columns.`,
  PBI_DATE_ROUNDED: (count) =>
    `${count} date value${count === 1 ? " was" : "s were"} rounded to the nearest millisecond. See the manifest for the affected columns.`,
  PBI_NONFINITE_AS_TEXT: (count) =>
    `${count} value${count === 1 ? " was" : "s were"} not a finite number and ${count === 1 ? "was" : "were"} written as text. See the manifest for the affected columns.`,
  PBI_NUMERIC_AS_TEXT: (count) =>
    `${count} number${count === 1 ? " was" : "s were"} written as exact text because a worksheet number would lose precision. See the manifest for the affected columns.`,
  PBI_TABLE_HIDDEN: (count) =>
    `${count} table${count === 1 ? " that the model marks hidden was" : "s that the model marks hidden were"} skipped. Export hidden tables as well to include them.`,
  PBI_TABLE_NO_EXPORTABLE_COLUMNS: (count) =>
    `${count} table${count === 1 ? " was" : "s were"} left out because no column could be read. See the manifest for which.`,
  PBI_TABLE_SPLIT: (count) =>
    `${count} table${count === 1 ? " has" : "s have"} more rows than one worksheet holds and ${count === 1 ? "was" : "were"} split across numbered worksheets. See the manifest for the row range of each part.`,
  PBI_TABLE_TOO_WIDE: (count) =>
    `${count} table${count === 1 ? " has" : "s have"} more columns than a worksheet holds and ${count === 1 ? "was" : "were"} left out. See the manifest for which.`,
  PBI_TEXT_TRUNCATED: (count) =>
    `${count} text value${count === 1 ? " was" : "s were"} shortened to the worksheet cell limit. See the manifest for the affected columns.`,
};

const UNVERIFIED_TEXT: Record<PbiUnverifiedCode, string> = {
  PBI_UNVERIFIED_BINARY_TYPE:
    "This model has binary columns. No test model has ever exercised that path, so check those values against Power BI before relying on them.",
  PBI_UNVERIFIED_BOOLEAN_TYPE:
    "This model has true or false columns. No test model has ever exercised that path, so check those values against Power BI before relying on them.",
  PBI_UNVERIFIED_COMPRESSION_CLASS:
    "Part of this model uses a column compression the reader has never seen in a test model. The values were decoded anyway, so check them against Power BI before relying on them.",
  PBI_UNVERIFIED_MULTIPLE_PARTITIONS:
    "Some columns of this model are stored in several partitions. No test model has ever exercised that path, so check the row order and totals against Power BI before relying on them.",
  PBI_UNVERIFIED_MULTIPLE_SEGMENTS:
    "Some columns of this model are stored in several segments. No test model has ever exercised that path, so check the row order and totals against Power BI before relying on them.",
  PBI_UNVERIFIED_NON_LATIN_DICTIONARY:
    "Some text in this model is stored in a non-Latin script dictionary. No test model has ever exercised that path, so check those values against Power BI before relying on them.",
  PBI_UNVERIFIED_XPRESS9_MULTITHREADED:
    "This model was written with the multithreaded compression variant. No test model has ever exercised that path, so check the exported rows against Power BI before relying on them.",
};

/**
 * One plain-language warning per distinct code present, sorted by code in
 * ascending ASCII order, aggregating in the unit that code describes. No source
 * value and no per-cell string ever appears.
 */
export function buildWarnings(
  manifest: PbiManifest,
  splitTables: number,
): string[] {
  const totals = new Map<PbiReasonCode, number>();
  const add = (code: PbiReasonCode, count: number): void => {
    if (count > 0) totals.set(code, (totals.get(code) ?? 0) + count);
  };
  const countTable = (
    reasons: readonly PbiReasonEntry[],
    columns: readonly PbiManifestColumn[],
  ): void => {
    for (const entry of reasons)
      if (entry.code !== "PBI_TABLE_SPLIT") add(entry.code, 1);
    for (const column of columns)
      for (const entry of column.reasons)
        add(
          entry.code,
          entry.code.startsWith("PBI_COLUMN_") ||
            entry.code === "PBI_BINARY_CELL_TOO_LONG"
            ? 1
            : entry.count,
        );
  };
  for (const table of manifest.tables) countTable(table.reasons, table.columns);
  for (const table of manifest.excludedTables)
    countTable(table.reasons, table.columns);
  add("PBI_TABLE_SPLIT", splitTables);
  const warnings = [...totals.entries()].map(([code, count]) => ({
    code: code as string,
    text: REASON_TEXT[code](count),
  }));
  for (const path of manifest.unverifiedPaths)
    warnings.push({ code: path.code, text: UNVERIFIED_TEXT[path.code] });
  return warnings
    .sort((left, right) => (left.code < right.code ? -1 : 1))
    .map((entry) => entry.text);
}

/** Deterministic property order, UTF-8, no timestamp and no random id. */
export function serializeManifest(manifest: PbiManifest): Uint8Array {
  const reasons = (entries: readonly PbiReasonEntry[]): unknown =>
    entries.map((entry) => ({ code: entry.code, count: entry.count }));
  const column = (value: PbiManifestColumn): unknown => ({
    name: value.name,
    id: value.id,
    type: value.type,
    ...(value.dax === undefined ? {} : { dax: value.dax }),
    ...(value.encoding === undefined ? {} : { encoding: value.encoding }),
    ...(value.dateEpoch === undefined ? {} : { dateEpoch: value.dateEpoch }),
    reasons: reasons(value.reasons),
  });
  const document = {
    schemaVersion: manifest.schemaVersion,
    tables: manifest.tables.map((table) => ({
      name: table.name,
      id: table.id,
      ...(table.dax === undefined ? {} : { dax: table.dax }),
      reasons: reasons(table.reasons),
      parts: table.parts.map((part) => ({
        sheetName: part.sheetName,
        start: part.start,
        end: part.end,
      })),
      columns: table.columns.map(column),
    })),
    excludedTables: manifest.excludedTables.map((table) => ({
      name: table.name,
      id: table.id,
      reasons: reasons(table.reasons),
      columns: table.columns.map(column),
    })),
    unverifiedPaths: manifest.unverifiedPaths.map((path) => ({
      code: path.code,
      count: path.count,
    })),
  };
  return new TextEncoder().encode(JSON.stringify(document, null, 2));
}
