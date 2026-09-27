import { isConsultChimpsError } from "@consultchimps/core";
import { memberBytes } from "./abf.js";
import type { AbfImage } from "./abf.js";
import type { CatalogColumn, CatalogTable } from "./catalog.js";
import type { PbiReasonCode, PbiUnverifiedCode } from "./errors.js";
import { HuffmanError } from "./huffman.js";
import {
  decodeSegmentIds,
  parseDictionary,
  parseIdf,
  parseIdfmeta,
  VertipaqAlignmentError,
  VertipaqError,
  XM_DATA_ID_NULL,
  XM_FIRST_DATA_ID,
} from "./vertipaq.js";
import type {
  Dictionary,
  DictionaryValue,
  ReserveBytes,
  SegmentDescriptor,
} from "./vertipaq.js";

export type PbiColumnType =
  | "string"
  | "int64"
  | "double"
  | "dateTimeSerial"
  | "currency"
  | "boolean"
  | "binary";

/**
 * null            a null cell in any column
 * string          "string"
 * bigint          "int64" (the stored 64-bit integer) and
 *                 "currency" (the stored count of ten-thousandths, undivided)
 * number          "double" and "dateTimeSerial" (the day serial, epoch 1899-12-30)
 * boolean         "boolean"
 * Uint8Array      "binary"
 */
export type PbiValue = null | string | number | bigint | boolean | Uint8Array;

export interface PbiColumn {
  readonly name: string;
  readonly id: number;
  readonly storagePosition: number;
  readonly type: PbiColumnType;
  readonly hidden: boolean;
  readonly dax?: string;
  readonly values: readonly PbiValue[];
}

export interface PbiTable {
  readonly id: number;
  readonly name: string;
  readonly hidden: boolean;
  readonly calculated: boolean;
  readonly dax?: string;
  readonly rowCount: number;
  readonly columns: readonly PbiColumn[];
}

/** AMO data type to the reader's column type. Anything else is unsupported. */
const AMO_TYPES = new Map<number, PbiColumnType>([
  [2, "string"],
  [6, "int64"],
  [8, "double"],
  [9, "dateTimeSerial"],
  [10, "currency"],
  [11, "boolean"],
  [17, "binary"],
]);

export type UnverifiedTally = Map<PbiUnverifiedCode, number>;

export function tallyUnverified(
  tally: UnverifiedTally,
  code: PbiUnverifiedCode,
  count = 1,
): void {
  if (count > 0) tally.set(code, (tally.get(code) ?? 0) + count);
}

export interface ColumnOutcome {
  readonly column: CatalogColumn;
  readonly type: PbiColumnType | undefined;
  readonly values: PbiValue[] | undefined;
  readonly excluded: PbiReasonCode | undefined;
  /**
   * How many values the exclusion code describes, when the code counts values
   * rather than the column itself. Only `PBI_BINARY_CELL_TOO_LONG` uses it, and
   * it is the real number of oversized cells, not one per column.
   */
  readonly excludedCount?: number;
}

function coerce(raw: DictionaryValue | null, type: PbiColumnType): PbiValue {
  if (raw === null) return null;
  switch (type) {
    case "string":
      return typeof raw === "string" ? raw : String(raw);
    case "int64":
      if (typeof raw === "bigint") return raw;
      if (typeof raw === "number")
        return Number.isFinite(raw) ? BigInt(Math.trunc(raw)) : null;
      return null;
    case "currency":
      if (typeof raw === "bigint") return raw;
      if (typeof raw === "number")
        return Number.isFinite(raw) ? BigInt(Math.trunc(raw)) : null;
      return null;
    case "double":
    case "dateTimeSerial":
      if (typeof raw === "bigint") return Number(raw);
      if (typeof raw === "number") return raw;
      return null;
    case "boolean":
      if (typeof raw === "bigint") return raw !== 0n;
      if (typeof raw === "number") return raw !== 0;
      return Boolean(raw);
    case "binary":
      if (typeof raw === "string") {
        const bytes = new Uint8Array(raw.length);
        for (let index = 0; index < raw.length; index++)
          bytes[index] = raw.charCodeAt(index) & 0xff;
        return bytes;
      }
      return null;
  }
}

/**
 * Decode one catalog column. Every failure is a per-column exclusion, never a
 * thrown pipeline error: a table survives on one decodable column.
 */
export function decodeColumn(
  image: Uint8Array,
  backup: AbfImage,
  column: CatalogColumn,
  rowCount: number,
  unverified: UnverifiedTally,
  reserve?: ReserveBytes,
): ColumnOutcome {
  const type = AMO_TYPES.get(column.dataType);
  if (type === undefined)
    return {
      column,
      type: undefined,
      values: undefined,
      excluded: "PBI_COLUMN_UNSUPPORTED_ENCODING",
    };
  if (type === "boolean")
    tallyUnverified(unverified, "PBI_UNVERIFIED_BOOLEAN_TYPE");
  if (type === "binary")
    tallyUnverified(unverified, "PBI_UNVERIFIED_BINARY_TYPE");
  if (column.idfs.length > 1)
    tallyUnverified(unverified, "PBI_UNVERIFIED_MULTIPLE_PARTITIONS");

  const mode =
    column.dictionary !== null
      ? "dictionary"
      : column.hierarchyIndex !== null
        ? "value"
        : "none";
  let dictionary: Dictionary | null = null;
  const values: PbiValue[] = [];
  try {
    if (mode === "dictionary") {
      dictionary = parseDictionary(
        memberBytes(image, backup, column.dictionary!),
        XM_FIRST_DATA_ID,
        reserve,
      );
      if (dictionary === null)
        return {
          column,
          type,
          values: undefined,
          excluded: "PBI_COLUMN_UNSUPPORTED_ENCODING",
        };
      tallyUnverified(
        unverified,
        "PBI_UNVERIFIED_NON_LATIN_DICTIONARY",
        dictionary.nonLatinPages,
      );
    }
    const exactBase = column.magnitude === 1;
    for (const name of column.idfs) {
      const descriptors = parseIdfmeta(
        memberBytes(image, backup, `${name}meta`),
        rowCount,
      );
      if (descriptors.length > 1)
        tallyUnverified(unverified, "PBI_UNVERIFIED_MULTIPLE_SEGMENTS");
      if (mode === "none") {
        // parseIdfmeta has already bounded the descriptors against the row
        // count, so this loop cannot push more nulls than the table has rows.
        for (const descriptor of descriptors)
          for (let index = 0; index < (descriptor.records || 0); index++)
            values.push(null);
        continue;
      }
      const segments = parseIdf(memberBytes(image, backup, name));
      for (let index = 0; index < descriptors.length; index++) {
        const descriptor = descriptors[index]!;
        const segment = segments[index];
        if (segment === undefined)
          return {
            column,
            type,
            values: undefined,
            excluded: "PBI_COLUMN_UNREADABLE",
          };
        if (descriptor.bitWidth === 0 && descriptor.countBitPacked > 0)
          // The width is unknowable, so the ids cannot be read at all. This is
          // the one unverified path that cannot attempt the decode.
          return {
            column,
            type,
            values: undefined,
            excluded: "PBI_COLUMN_UNSUPPORTED_ENCODING",
          };
        if (descriptor.bitWidth === 0)
          tallyUnverified(unverified, "PBI_UNVERIFIED_COMPRESSION_CLASS");
        appendSegment(
          values,
          decodeSegmentIds(segment, descriptor),
          descriptor,
          dictionary,
          column,
          type,
          exactBase,
        );
      }
    }
  } catch (error) {
    return {
      column,
      type,
      values: undefined,
      excluded: columnFailureReason(error),
    };
  }
  return { column, type, values, excluded: undefined };
}

/**
 * Codes that say the file is at fault, and are the only refusals a column
 * decoder may report as damaged source data.
 *
 * The list is explicit rather than "any error this repository raises", because
 * `ConsultChimpsError` is this package's own class: a refusal added to the
 * decode path later, or raised through it by a shared package, would otherwise
 * be reported to the user as damage to their model. Each entry here is a refusal
 * about the bytes of this column's own members.
 */
const FILE_FAULT_CODES: ReadonlySet<string> = new Set([
  // The backup container could not produce this column's member.
  "PBI_MODEL_UNREADABLE",
  // The container the member lives in is not the shape a .pbix declares.
  "PBI_INVALID_CONTAINER",
  // The member is there and encrypted, which is the file's state, not ours.
  "PBI_MODEL_ENCRYPTED",
]);

/**
 * Which exclusion a failure inside a column decoder becomes.
 *
 * Classification is by code, not by class. A capacity refusal is the caller's,
 * not this column's, so it passes through: a budget that says stop must stop the
 * export, not quietly drop one column and carry on. A refusal whose code says
 * the file is at fault, and the reader's two structural error types, which exist
 * only to report a column store that contradicts itself, stay
 * `PBI_COLUMN_UNREADABLE`.
 *
 * Everything else is an unexpected fault in this reader, not evidence about the
 * file, so it excludes the column as `PBI_COLUMN_DECODER_ERROR` rather than
 * telling the user their data is damaged. That includes a `ConsultChimpsError`
 * carrying any other code, which is the case a class-based rule got wrong. The
 * export continues either way, and nothing of the original error's text
 * survives, so a file-declared length cannot ride out inside a RangeError
 * message.
 */
export function columnFailureReason(error: unknown): PbiReasonCode {
  if (error instanceof VertipaqAlignmentError)
    return "PBI_COLUMN_ROW_ALIGNMENT_UNRECOVERABLE";
  if (isConsultChimpsError(error)) {
    if (error.code === "PBI_EXPORT_LIMIT_EXCEEDED") throw error;
    return FILE_FAULT_CODES.has(error.code)
      ? "PBI_COLUMN_UNREADABLE"
      : "PBI_COLUMN_DECODER_ERROR";
  }
  // Both are declared by this package for one purpose: a member whose own
  // structure is inconsistent. Neither is ever raised for an internal fault.
  if (error instanceof VertipaqError || error instanceof HuffmanError)
    return "PBI_COLUMN_UNREADABLE";
  if (error instanceof Error) return "PBI_COLUMN_DECODER_ERROR";
  throw error;
}

function appendSegment(
  values: PbiValue[],
  ids: Float64Array,
  descriptor: SegmentDescriptor,
  dictionary: Dictionary | null,
  column: CatalogColumn,
  type: PbiColumnType,
  exactBase: boolean,
): void {
  if (dictionary !== null) {
    const entries = dictionary.values;
    const low = dictionary.minId;
    const high = low + entries.length;
    for (let index = 0; index < ids.length; index++) {
      const id = ids[index]!;
      // A data id outside the dictionary range is a null cell, not a failure.
      values.push(
        id >= low && id < high ? coerce(entries[id - low]!, type) : null,
      );
    }
    return;
  }
  const nullId = descriptor.hasNulls ? XM_DATA_ID_NULL : null;
  const base = column.baseId;
  const magnitude = column.magnitude;
  const numericBase = Number(base);
  for (let index = 0; index < ids.length; index++) {
    const id = ids[index]!;
    if (nullId !== null && id === nullId) {
      values.push(null);
      continue;
    }
    if (exactBase && (type === "int64" || type === "currency")) {
      // Magnitude 1 means the stored value is the id plus the base exactly, so
      // it stays a bigint. Passing it through Number would lose the low digits
      // of any identifier past 2^53.
      values.push(BigInt(id) + base);
      continue;
    }
    values.push(coerce((id + numericBase) / magnitude, type));
  }
}

export interface DecodedTable {
  readonly table: PbiTable;
  readonly reasons: ReadonlyMap<number, PbiReasonCode>;
}

/**
 * Assemble one table from its decoded columns. The row count is the catalog's
 * declared count for the table; a column whose decoded length differs cannot be
 * placed on it and is excluded, without shifting or dropping any row.
 */
export function assembleTable(
  source: CatalogTable,
  outcomes: readonly ColumnOutcome[],
): DecodedTable {
  const reasons = new Map<number, PbiReasonCode>();
  // The row count is the catalog's, not a vote among the decoded columns. A
  // majority that decodes to a common wrong length must not be able to redefine
  // the table and exclude the columns that decoded correctly.
  const rowCount = source.rowCount;
  const columns: PbiColumn[] = [];
  for (const outcome of outcomes) {
    if (outcome.excluded !== undefined) {
      reasons.set(outcome.column.id, outcome.excluded);
      continue;
    }
    if (outcome.values!.length !== rowCount) {
      reasons.set(outcome.column.id, "PBI_COLUMN_ROW_ALIGNMENT_UNRECOVERABLE");
      continue;
    }
    columns.push({
      name: outcome.column.name,
      id: outcome.column.id,
      storagePosition: outcome.column.storagePosition,
      type: outcome.type!,
      hidden: outcome.column.hidden,
      ...(outcome.column.dax === undefined ? {} : { dax: outcome.column.dax }),
      values: outcome.values!,
    });
  }
  return {
    table: {
      id: source.id,
      name: source.name,
      hidden: source.hidden,
      calculated: source.calculated,
      ...(source.dax === undefined ? {} : { dax: source.dax }),
      rowCount,
      columns,
    },
    reasons,
  };
}
