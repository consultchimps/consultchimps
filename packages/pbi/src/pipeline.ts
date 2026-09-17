import { safeNameFragment } from "@consultchimps/core";
import type { ByteArtifact, ByteOperationOutcome } from "@consultchimps/core";
import { hasMember, memberBytes, parseAbf } from "./abf.js";
import type { AbfImage } from "./abf.js";
import {
  PipelineBudget,
  outputLimitExceeded,
  validateExportOptions,
} from "./budget.js";
import type { PbiExportOptions, PbiReadOptions } from "./budget.js";
import { readCatalog } from "./catalog.js";
import type { CatalogColumn, CatalogTable } from "./catalog.js";
import { readPbiModelPart } from "./container.js";
import { noExportableTables, unreadableModel } from "./errors.js";
import type { ExclusionCounts, PbiReasonCode } from "./errors.js";
import { assembleTable, decodeColumn, tallyUnverified } from "./model.js";
import type {
  ColumnOutcome,
  PbiColumn,
  PbiTable,
  UnverifiedTally,
} from "./model.js";
import { buildWarnings, reasonEntries, serializeManifest } from "./manifest.js";
import type {
  PbiManifest,
  PbiManifestColumn,
  PbiManifestExclusion,
  PbiManifestTable,
  PbiReasonEntry,
  PbiWorksheetPart,
} from "./manifest.js";
import { SheetNameAllocator } from "./sheets.js";
import { decompressModelPart } from "./xpress9/stream.js";
import {
  MAX_DATA_ROWS_PER_PART,
  MAX_WORKSHEET_COLUMNS,
  dateColumnReadable,
  oversizedBinaryValues,
} from "./values.js";
import { writeWorkbook } from "./workbook.js";
import type { ConversionCounts, WorksheetPlan } from "./workbook.js";

/** The six export metrics of ADR 0004 Decision 9, in their fixed order. */
export type PowerBiExportMetric =
  | "inputFiles"
  | "outputFiles"
  | "exportedTables"
  | "exportedColumns"
  | "exportedRows"
  | "outputWorksheets";

export interface PbiModel {
  readonly tables: readonly PbiTable[];
  readonly manifest: PbiManifest;
  readonly warnings: readonly string[];
}

const WORKBOOK_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const MANIFEST_MEDIA_TYPE = "application/json";
const CATALOG_MEMBER = "metadata.sqlitedb";

/**
 * What one decoded cell of each AMO data type costs in the JavaScript heap.
 *
 * Measured, not assumed: `scripts/measure-cell-cost.ts` decodes the corpus one
 * column type at a time and divides the heap growth by the cell count. A
 * reference is eight bytes, which is all a shared dictionary string costs, but a
 * boxed bigint measured 32 and a boxed double 24, and charging eight for
 * everything is what made the first build's estimate a quarter of the measured
 * peak. The figures below are those measurements, with strings carrying
 * headroom over the dictionary term that is charged beside them.
 */
const CELL_BYTES: Record<number, number> = {
  2: 16, // string: an 8-byte slot, plus headroom over the shared dictionary
  6: 32, // int64: a boxed bigint
  8: 24, // double: a boxed double in a mixed-type array
  9: 24, // dateTimeSerial: the same boxed double
  10: 32, // currency: a boxed bigint
  11: 8, // boolean: the slot alone
  17: 64, // binary: a slot and a small typed array per cell
};

/** The fallback is the widest measured shape, not the narrowest. */
function cellBytes(dataType: number): number {
  return CELL_BYTES[dataType] ?? 32;
}

interface ExcludedTable {
  readonly source: CatalogTable;
  readonly reasons: PbiReasonCode[];
  readonly columns: PbiManifestColumn[];
}

interface DecodeResult {
  readonly tables: PbiTable[];
  readonly manifestTables: Map<number, PbiManifestTable>;
  readonly excluded: ExcludedTable[];
  readonly unverified: UnverifiedTally;
  readonly hiddenExcluded: boolean;
  readonly budget: PipelineBudget;
}

function manifestColumn(
  column: CatalogColumn,
  type: PbiManifestColumn["type"],
  reasons: readonly PbiReasonEntry[],
): PbiManifestColumn {
  return {
    name: column.name,
    id: column.id,
    type,
    ...(column.dax === undefined ? {} : { dax: column.dax }),
    reasons: [...reasons],
  };
}

/** What the catalog reservation needs from an open reader. */
export interface ClosableCatalog {
  readonly wasmMemoryBytes: number;
  close(): void;
}

/**
 * Charge the open reader's linear memory, inside the reader's own cleanup
 * scope. A refusal here describes memory the reader is already holding, so the
 * reader is released before the refusal propagates rather than leaving one
 * WebAssembly instance alive per failed attempt.
 */
export function reserveCatalogMemory(
  catalog: ClosableCatalog,
  budget: PipelineBudget,
): void {
  try {
    budget.reserve("sqliteLinearMemory", catalog.wasmMemoryBytes, "catalog");
  } catch (error) {
    try {
      catalog.close();
    } catch {
      // The capacity refusal is the actionable error.
    }
    throw error;
  }
}

/**
 * Decode one column with its dictionary reservation, inside the per-column
 * failure boundary. A column naming a dictionary member the backup does not
 * carry excludes that one column; it never aborts an export the rest of the
 * model could still produce.
 */
export function decodeColumnWithBudget(
  image: Uint8Array,
  backup: AbfImage,
  column: CatalogColumn,
  rowCount: number,
  budget: PipelineBudget,
  unverified: UnverifiedTally,
): ColumnOutcome {
  let dictionaryBytes = 0;
  try {
    if (column.dictionary !== null)
      dictionaryBytes =
        memberBytes(image, backup, column.dictionary).byteLength * 2;
  } catch {
    return {
      column,
      type: undefined,
      values: undefined,
      excluded: "PBI_COLUMN_UNREADABLE",
    };
  }
  if (dictionaryBytes > 0)
    budget.reserve("dictionaries", dictionaryBytes, "decode");
  return validateTypedValues(
    decodeColumn(image, backup, column, rowCount, unverified, (bytes) =>
      // Each Huffman page is charged from its validated bit count before it
      // decodes, so an expansion the member's own length does not bound refuses
      // with the capacity code instead of ending in an out-of-memory.
      budget.reserve("dictionaries", bytes, "decode"),
    ),
  );
}

/**
 * Why a table is not eligible at all, before any column is decoded. Exclusions
 * apply in ADR Decision 9's fixed order: the hidden policy, then the source
 * column limit. "No exportable columns" can only be known after the decode.
 */
export function tableExclusion(
  source: CatalogTable,
  includeHiddenTables: boolean,
): PbiReasonCode | undefined {
  if (source.hidden && !includeHiddenTables) return "PBI_TABLE_HIDDEN";
  if (source.columns.length > MAX_WORKSHEET_COLUMNS)
    return "PBI_TABLE_TOO_WIDE";
  return undefined;
}

/**
 * Decode every eligible table. Failures are exclusions, not refusals: a table
 * survives on one decodable column, and dropping a column never shifts a row.
 */
async function decodeModel(
  input: Uint8Array,
  options: PbiExportOptions,
  includeOutput: boolean,
): Promise<DecodeResult> {
  const limits = validateExportOptions(options, includeOutput);
  const budget = new PipelineBudget(limits);
  const modelPart = readPbiModelPart(input, {
    inputBytes: limits.inputBytes,
    decodedBytes: limits.decodedBytes,
    peakBytes: limits.peakBytes,
  });
  budget.reserve("retainedInput", input.byteLength, "container");
  budget.reserve("modelPart", modelPart.byteLength, "model-part");
  budget.decode(modelPart.byteLength, "model-part");

  const unverified: UnverifiedTally = new Map();
  const stream = await decompressModelPart(
    modelPart,
    budget,
    options.runtime?.xpress9,
  );
  if (stream.compression === "multithreaded")
    tallyUnverified(unverified, "PBI_UNVERIFIED_XPRESS9_MULTITHREADED");

  const backup = parseAbf(stream.bytes, budget);
  if (!hasMember(backup, CATALOG_MEMBER)) throw unreadableModel("backup");
  const catalogBytes = memberBytes(stream.bytes, backup, CATALOG_MEMBER);
  budget.reserve("catalogBuffer", catalogBytes.byteLength, "catalog");
  const catalog = await readCatalog(catalogBytes, options.runtime?.sql);
  reserveCatalogMemory(catalog, budget);

  const tables: PbiTable[] = [];
  const manifestTables = new Map<number, PbiManifestTable>();
  const excluded: ExcludedTable[] = [];
  let hiddenExcluded = false;
  try {
    for (const source of catalog.tables) {
      const ineligible = tableExclusion(source, limits.includeHiddenTables);
      if (ineligible !== undefined) {
        if (ineligible === "PBI_TABLE_HIDDEN") hiddenExcluded = true;
        excluded.push({ source, reasons: [ineligible], columns: [] });
        continue;
      }
      const outcomes: ColumnOutcome[] = [];
      for (const column of source.columns) {
        // Reserved BEFORE the decode, from the catalog's declared row count and
        // the column's declared type, so the refusal happens instead of the
        // allocation rather than after it.
        budget.reserve(
          "decoderWorkingSet",
          source.rowCount * cellBytes(column.dataType),
          "decode",
        );
        outcomes.push(
          decodeColumnWithBudget(
            stream.bytes,
            backup,
            column,
            source.rowCount,
            budget,
            unverified,
          ),
        );
      }
      const { table, reasons } = assembleTable(source, outcomes);
      const columnEntries: PbiManifestColumn[] = [];
      for (const outcome of outcomes) {
        const code = reasons.get(outcome.column.id);
        columnEntries.push(
          manifestColumn(
            outcome.column,
            outcome.type ?? "string",
            code === undefined
              ? []
              : [{ code, count: outcome.excludedCount ?? 1 }],
          ),
        );
      }
      if (table.columns.length === 0) {
        excluded.push({
          source,
          // The column exclusions are kept so the user can see why none remain.
          reasons: ["PBI_TABLE_NO_EXPORTABLE_COLUMNS"],
          columns: columnEntries,
        });
        continue;
      }
      tables.push(table);
      manifestTables.set(table.id, {
        name: table.name,
        id: table.id,
        ...(table.dax === undefined ? {} : { dax: table.dax }),
        reasons: [],
        parts: [],
        columns: columnEntries,
      });
    }
  } catch (error) {
    // A close failure must not mask the failure that is already propagating:
    // the decode error is the actionable one, and whatever the reader could not
    // release belongs to it alone and is unreachable afterwards.
    try {
      catalog.close();
    } catch {
      // Deliberately dropped. See above.
    }
    throw error;
  }
  catalog.close();
  return {
    tables,
    manifestTables,
    excluded,
    unverified,
    hiddenExcluded,
    budget,
  };
}

/**
 * Semantic validation that runs after a successful byte decode: a date column
 * holding a non-finite serial could not be read as its declared type, and an
 * oversized base64 value excludes its whole column rather than being truncated.
 */
export function validateTypedValues(outcome: ColumnOutcome): ColumnOutcome {
  if (outcome.values === undefined || outcome.excluded !== undefined)
    return outcome;
  if (outcome.type === "dateTimeSerial" && !dateColumnReadable(outcome.values))
    return { ...outcome, values: undefined, excluded: "PBI_COLUMN_UNREADABLE" };
  if (outcome.type === "binary") {
    const oversized = oversizedBinaryValues(outcome.values);
    if (oversized > 0)
      return {
        ...outcome,
        values: undefined,
        excluded: "PBI_BINARY_CELL_TOO_LONG",
        // The code counts oversized values, so the manifest reports how many
        // there were, not one per column.
        excludedCount: oversized,
      };
  }
  return outcome;
}

function unverifiedPaths(
  tally: UnverifiedTally,
): PbiManifest["unverifiedPaths"] {
  return [...tally.entries()]
    .filter(([, count]) => count > 0)
    .map(([code, count]) => ({ code, count }))
    .sort((left, right) => (left.code < right.code ? -1 : 1));
}

function exclusionCounts(
  excluded: readonly ExcludedTable[],
): ExclusionCounts[] {
  const totals = new Map<PbiReasonCode, { tables: number; columns: number }>();
  for (const entry of excluded) {
    for (const code of entry.reasons) {
      const bucket = totals.get(code) ?? { tables: 0, columns: 0 };
      bucket.tables++;
      totals.set(code, bucket);
    }
    for (const column of entry.columns)
      for (const reason of column.reasons) {
        const bucket = totals.get(reason.code) ?? { tables: 0, columns: 0 };
        bucket.columns += reason.count;
        totals.set(reason.code, bucket);
      }
  }
  return [...totals.entries()]
    .map(([code, value]) => ({
      code,
      tables: value.tables,
      columns: value.columns,
    }))
    .sort((left, right) => (left.code < right.code ? -1 : 1));
}

function excludedEntries(
  excluded: readonly ExcludedTable[],
): PbiManifestExclusion[] {
  return excluded.map((entry) => ({
    name: entry.source.name,
    id: entry.source.id,
    reasons: entry.reasons
      .map((code) => ({ code, count: 1 }))
      .sort((left, right) => (left.code < right.code ? -1 : 1)),
    columns: entry.columns,
  }));
}

/** Zero-based half-open parts of `[0, rowCount)`, `[0, 0)` for a header-only table. */
export function planParts(rowCount: number): { start: number; end: number }[] {
  if (rowCount === 0) return [{ start: 0, end: 0 }];
  const parts: { start: number; end: number }[] = [];
  for (let start = 0; start < rowCount; start += MAX_DATA_ROWS_PER_PART)
    parts.push({
      start,
      end: Math.min(start + MAX_DATA_ROWS_PER_PART, rowCount),
    });
  return parts;
}

export interface Allocation {
  readonly plans: WorksheetPlan[];
  readonly parts: Map<number, PbiWorksheetPart[]>;
  readonly splitTables: number;
}

/**
 * Deterministic allocation: tables in ascending numeric id, each claiming every
 * one of its names before the next table is considered.
 */
export function allocateWorksheets(tables: readonly PbiTable[]): Allocation {
  const allocator = new SheetNameAllocator();
  const plans: WorksheetPlan[] = [];
  const parts = new Map<number, PbiWorksheetPart[]>();
  let splitTables = 0;
  for (const table of tables) {
    const ranges = planParts(table.rowCount);
    if (ranges.length > 1) splitTables++;
    const entries: PbiWorksheetPart[] = [];
    for (let index = 0; index < ranges.length; index++) {
      const sheetName = allocator.claim(table.name, index + 1);
      entries.push({ sheetName, ...ranges[index]! });
      plans.push({
        sheetName,
        start: ranges[index]!.start,
        end: ranges[index]!.end,
        columns: table.columns.map((column: PbiColumn) => ({
          key: `${table.id}:${column.id}`,
          header: column.name,
          type: column.type,
          values: column.values,
        })),
      });
    }
    parts.set(table.id, entries);
  }
  return { plans, parts, splitTables };
}

function buildManifest(
  result: DecodeResult,
  parts: Map<number, PbiWorksheetPart[]> | undefined,
  counts: ConversionCounts | undefined,
): PbiManifest {
  const tables: PbiManifestTable[] = [];
  for (const table of result.tables) {
    const entry = result.manifestTables.get(table.id)!;
    const worksheetParts = parts?.get(table.id) ?? [];
    const tableReasons =
      worksheetParts.length > 1
        ? [{ code: "PBI_TABLE_SPLIT" as PbiReasonCode, count: 1 }]
        : [];
    tables.push({
      ...entry,
      reasons: tableReasons,
      parts: worksheetParts,
      columns: entry.columns.map((column) => {
        const emitted = counts?.get(`${table.id}:${column.id}`);
        const merged = new Map<PbiReasonCode, number>();
        for (const reason of column.reasons)
          merged.set(reason.code, reason.count);
        if (emitted !== undefined)
          for (const [code, count] of emitted)
            merged.set(code, (merged.get(code) ?? 0) + count);
        const dateFallback = merged.has("PBI_DATE_AS_TEXT");
        return {
          ...column,
          ...(merged.has("PBI_BINARY_AS_BASE64")
            ? { encoding: "base64" as const }
            : {}),
          ...(dateFallback ? { dateEpoch: "1899-12-30" as const } : {}),
          reasons: reasonEntries(merged),
        };
      }),
    });
  }
  return {
    schemaVersion: 1,
    tables,
    excludedTables: excludedEntries(result.excluded),
    unverifiedPaths: unverifiedPaths(result.unverified),
  };
}

/** Decode every exportable table. No workbook and no manifest serialization. */
export async function readPbiTables(
  input: Uint8Array,
  options: PbiReadOptions = {},
): Promise<PbiModel> {
  const result = await decodeModel(input, options, false);
  if (result.tables.length === 0)
    throw noExportableTables(
      exclusionCounts(result.excluded),
      result.hiddenExcluded,
    );
  const { parts, splitTables } = allocateWorksheets(result.tables);
  const manifest = buildManifest(result, parts, undefined);
  return {
    tables: result.tables,
    manifest,
    warnings: buildWarnings(manifest, splitTables),
  };
}

/** The workbook stem and its two output names, independent of input and locale. */
export function outputNames(outputName: string): {
  workbook: string;
  manifest: string;
} {
  const normalized = outputName
    .toWellFormed()
    .normalize("NFKC")
    .trim()
    .replace(/\.xlsx$/i, "");
  const stem = safeNameFragment(normalized, "power-bi-tables");
  return { workbook: `${stem}.xlsx`, manifest: `${stem}.manifest.json` };
}

/** The full operation: workbook bytes, manifest bytes, metrics and warnings. */
export async function exportPbiTables(
  input: Uint8Array,
  options: PbiExportOptions = {},
): Promise<ByteOperationOutcome<PowerBiExportMetric>> {
  const result = await decodeModel(input, options, true);
  if (result.tables.length === 0)
    throw noExportableTables(
      exclusionCounts(result.excluded),
      result.hiddenExcluded,
    );
  const { plans, parts, splitTables } = allocateWorksheets(result.tables);
  const budget = result.budget;
  // The materialized worksheet XML and the finished ZIP both stay live while
  // the package is produced, so both are reserved before either is allocated.
  let cells = 0;
  for (const plan of plans)
    cells += (plan.end - plan.start) * plan.columns.length;
  budget.reserve("workbook", cells * 48 + 1024 * 1024, "workbook");
  // The limit is enforced while the package is produced: the worksheet XML is
  // refused as it is built, and the compressed bytes are counted as they arrive,
  // so an over-limit export costs neither the whole archive nor the memory the
  // limit existed to guard.
  const { bytes: workbookBytes, counts } = await writeWorkbook(
    plans,
    budget.limits.outputBytes,
  );

  const manifest = buildManifest(result, parts, counts);
  const manifestBytes = serializeManifest(manifest);
  // Both artifacts are reserved together before either is emitted.
  const required = workbookBytes.byteLength + manifestBytes.byteLength;
  if (required > budget.limits.outputBytes)
    throw outputLimitExceeded(budget.limits.outputBytes, required);

  const names = outputNames(budget.limits.outputName);
  const outputs: ByteArtifact[] = [
    {
      name: names.workbook,
      bytes: workbookBytes,
      mediaType: WORKBOOK_MEDIA_TYPE,
    },
    {
      name: names.manifest,
      bytes: manifestBytes,
      mediaType: MANIFEST_MEDIA_TYPE,
    },
  ];
  let exportedColumns = 0;
  let exportedRows = 0;
  for (const table of result.tables) {
    exportedColumns += table.columns.length;
    exportedRows += table.rowCount;
  }
  return {
    result: {
      operation: "pbi.export-tables",
      artifacts: outputs.map((output) => ({
        kind: "file" as const,
        path: output.name,
        mediaType: output.mediaType!,
      })),
      warnings: buildWarnings(manifest, splitTables),
      metrics: {
        inputFiles: 1,
        outputFiles: 2,
        exportedTables: result.tables.length,
        exportedColumns,
        exportedRows,
        outputWorksheets: plans.length,
      },
    },
    outputs,
  };
}
