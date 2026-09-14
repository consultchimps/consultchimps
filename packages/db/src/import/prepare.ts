import { sha256 } from "@noble/hashes/sha2.js";
import { bytesToHex } from "@noble/hashes/utils.js";

import { isConsultChimpsError, throwIfAborted } from "@consultchimps/core";

import { engineOf, valueAsString } from "../database.js";
import { databaseError } from "../errors.js";
import type { EngineValue } from "../internal/engine.js";
import { PREPARED_STORAGE } from "../internal/storage-layouts.js";
import { insertInternalRow } from "../internal/storage-schema.js";
import { CAPTURE_ROW_TABLE } from "../metadata.js";
import { validateTableSchema } from "../schema.js";
import {
  PREPARED_BINDING_TABLE,
  PREPARED_CAPTURE_TABLE,
  PREPARED_ROW_TABLE,
  assertImportBatchWritable,
  preparedEngineOf,
  readPreparedReviewSnapshot,
  updatePreparedCaptureMetadata,
  updatePreparedPlan,
} from "../prepared.js";
import {
  addToProfile,
  emptyProfile,
  inferColumns,
  type ColumnProfile,
} from "./inference.js";
import {
  evaluateConflicts,
  findReusableCapture,
  preparedCapturesFromEngine,
} from "./planning.js";
import type {
  ImportRegionReader,
  PrepareImportOptions,
  PrepareImportOutcome,
  PrepareImportReviewOutcome,
  ImportBatchPage,
} from "./types.js";
import {
  validateColumnMappings,
  validateImportProfile,
} from "../validators.js";
import { assertValidImportDateCell } from "./date-cell.js";
import { assertValidImportNumberCell } from "./number-cell.js";
import { inspectUpdatedImport, assertImportReviewPage } from "./inspection.js";
import { createCaptureRowChecksum } from "./row-checksum.js";
import { assertCaptureRowCount, readSourceRowPage } from "./source-rows.js";

const HASH_CHUNK_BYTES = 1024 * 1024;
const CAPTURE_BATCH_ROWS = 2_000;

function validateSourceSelections(
  sources: PrepareImportOptions["sources"],
): void {
  const sourceKeys = new Set<string>();
  for (const source of sources) {
    if (sourceKeys.has(source.key)) {
      throw databaseError(
        "DB_DUPLICATE_IMPORT_SOURCE",
        `Source key "${source.key}" appears more than once. Give each import source a unique key.`,
        { source: source.key },
      );
    }
    sourceKeys.add(source.key);
    const selectionKeys = new Set<string>();
    for (const selection of source.selections) {
      if (selectionKeys.has(selection.key)) {
        throw databaseError(
          "DB_DUPLICATE_IMPORT_SELECTION",
          `Source "${source.key}" declares selection key "${selection.key}" more than once. Give each selection in a source a unique key.`,
          { source: source.key, selection: selection.key },
        );
      }
      selectionKeys.add(selection.key);
    }
  }
}

async function checksumDatabaseCapture(options: {
  readonly database: PrepareImportOptions["database"];
  readonly captureId: string;
  readonly rowCount: bigint;
  readonly signal?: AbortSignal | undefined;
}): Promise<string> {
  return engineOf(options.database).readTransaction(async (transaction) => {
    const checksum = createCaptureRowChecksum();
    let cursor: bigint | undefined;
    let rowsRead = 0n;
    while (true) {
      throwIfAborted(options.signal, "db.import.prepare");
      const page = await readSourceRowPage({
        engine: transaction,
        table: CAPTURE_ROW_TABLE,
        captureId: options.captureId,
        cursor,
        limit: CAPTURE_BATCH_ROWS,
        owner: "database",
      });
      throwIfAborted(options.signal, "db.import.prepare");
      if (page.rows.length === 0) break;
      for (const { row, sourceRow } of page.rows) {
        checksum.update(
          sourceRow,
          valueAsString(row["values_json"], "captured values"),
        );
      }
      rowsRead += BigInt(page.rows.length);
      cursor = page.cursor;
    }
    assertCaptureRowCount({
      owner: "database",
      captureId: options.captureId,
      expected: options.rowCount,
      actual: rowsRead,
    });
    return checksum.digest();
  });
}

function validateReaderColumns(options: {
  readonly source: string;
  readonly selection: string;
  readonly columns: readonly string[];
  readonly profile: PrepareImportOptions["profile"];
}): void {
  const exactNames = new Set<string>();
  for (const column of options.columns) {
    if (exactNames.has(column)) {
      throw databaseError(
        "DB_INVALID_RECIPE",
        `Source "${options.source}" selection "${options.selection}" declares column "${column}" more than once. Source columns must be unique.`,
        {
          source: options.source,
          selection: options.selection,
          column,
        },
      );
    }
    exactNames.add(column);
  }
  const route = options.profile.routes.find(
    (candidate) =>
      candidate.source === options.source &&
      candidate.selection === options.selection,
  );
  if (route === undefined) return;
  const effectiveColumns =
    route.columns.length === 0
      ? options.columns.map((column) => ({
          source: column,
          target: column,
          type: "text" as const,
        }))
      : route.columns;
  if (route.columns.length === 0) {
    validateColumnMappings(options.source, options.selection, effectiveColumns);
  }
  if (route.destination.kind === "new-table-infer") {
    validateTableSchema({
      name: route.destination.name,
      columns: effectiveColumns.map((column) => ({
        name: column.target,
        type: "text",
      })),
      recordId: route.destination.recordId,
      foreignKeys: [],
    });
    return;
  }
}

async function readAndClose<T>(
  reader: ImportRegionReader,
  read: () => Promise<T>,
): Promise<T> {
  let result:
    | { readonly ok: true; readonly value: T }
    | { readonly ok: false; readonly error: unknown };
  try {
    result = { ok: true, value: await read() };
  } catch (error) {
    result = { ok: false, error };
  }
  try {
    await reader.close();
  } catch (error) {
    if (!result.ok) {
      throw new AggregateError(
        [result.error, error],
        "The import reader failed and could not be closed.",
        { cause: error },
      );
    }
    throw error;
  }
  if (!result.ok) throw result.error;
  return result.value;
}

async function hashSource(
  source: PrepareImportOptions["sources"][number]["bytes"],
  signal?: AbortSignal,
): Promise<string> {
  const hash = sha256.create();
  for (let offset = 0; offset < source.size; offset += HASH_CHUNK_BYTES) {
    throwIfAborted(signal, "db.import.prepare");
    const expected = Math.min(HASH_CHUNK_BYTES, source.size - offset);
    const bytes = await source.readAt(offset, expected, signal);
    if (bytes.length !== expected) {
      throw databaseError(
        "DB_SOURCE_SHORT_READ",
        `The source "${source.name}" changed or ended while it was being hashed.`,
        { source: source.name, offset, expected, actual: bytes.length },
      );
    }
    hash.update(bytes);
  }
  return bytesToHex(hash.digest());
}

export function prepareImport(
  options: PrepareImportOptions & { readonly reviewPage: ImportBatchPage },
): Promise<PrepareImportReviewOutcome>;
export function prepareImport(
  options: PrepareImportOptions,
): Promise<PrepareImportOutcome>;
export async function prepareImport(
  options: PrepareImportOptions,
): Promise<PrepareImportOutcome | PrepareImportReviewOutcome> {
  if (options.reviewPage !== undefined)
    assertImportReviewPage(options.reviewPage);
  throwIfAborted(options.signal, "db.import.prepare");
  validateImportProfile(options.profile);
  validateSourceSelections(options.sources);
  if (options.prepared.databaseId !== options.database.id) {
    throw databaseError(
      "DB_PREPARED_WRONG_DATABASE",
      "The import batch belongs to a different database.",
    );
  }
  const preparedEngine = preparedEngineOf(options.prepared);
  let sourcesRead = 0;
  let sourcesReused = 0;
  let rowsCaptured = 0;
  for (const source of options.sources) {
    let capturedSource = false;
    let reusedSource = false;
    const contentHash = await hashSource(source.bytes, options.signal);
    await source.verifyUnchanged?.();
    for (const selection of source.selections) {
      throwIfAborted(options.signal, "db.import.prepare");
      const existingBinding = await preparedEngine.query(
        `SELECT c.content_hash, c.reader_version FROM ${PREPARED_BINDING_TABLE} b JOIN ${PREPARED_CAPTURE_TABLE} c ON c.capture_id = b.capture_id WHERE b.source_key = ? AND b.selection_key = ?`,
        [source.key, selection.key],
      );
      if (existingBinding[0] !== undefined) {
        if (
          valueAsString(existingBinding[0]["content_hash"], "content hash") !==
            contentHash ||
          valueAsString(
            existingBinding[0]["reader_version"],
            "reader version",
          ) !== source.readerVersion
        ) {
          throw databaseError(
            "DB_PREPARED_SOURCE_CHANGED",
            `Source "${source.key}" selection "${selection.key}" differs from the content already stored in this import batch. Create a new batch for the changed source.`,
            { source: source.key, selection: selection.key },
          );
        }
        continue;
      }
      const duplicate = await preparedEngine.query(
        `SELECT capture_id FROM ${PREPARED_CAPTURE_TABLE} WHERE content_hash = ? AND selection_key = ? AND reader_version = ? LIMIT 1`,
        [contentHash, selection.key, source.readerVersion],
      );
      if (duplicate[0] !== undefined) {
        const duplicateCaptureId = valueAsString(
          duplicate[0]["capture_id"],
          "capture ID",
        );
        await updatePreparedCaptureMetadata(
          options.prepared,
          async (transaction) => {
            await insertInternalRow(
              transaction,
              PREPARED_STORAGE.tables.bindings,
              {
                source_key: source.key,
                selection_key: selection.key,
                capture_id: duplicateCaptureId,
                display_name: source.bytes.name,
              },
            );
          },
        );
        reusedSource = true;
        continue;
      }
      const reusable = await findReusableCapture({
        database: options.database,
        contentHash,
        selectionKey: selection.key,
        readerVersion: source.readerVersion,
      });
      if (reusable !== null) {
        const rowChecksum = await checksumDatabaseCapture({
          database: options.database,
          captureId: reusable.captureId,
          rowCount: reusable.rowCount,
          signal: options.signal,
        });
        await updatePreparedCaptureMetadata(
          options.prepared,
          async (transaction) => {
            await insertInternalRow(
              transaction,
              PREPARED_STORAGE.tables.captures,
              {
                capture_id: reusable.captureId,
                source_file_id: reusable.sourceFileId,
                source_key: source.key,
                display_name: source.bytes.name,
                selection_key: selection.key,
                selection_label: selection.label,
                reader_version: source.readerVersion,
                content_hash: contentHash,
                byte_count: BigInt(source.bytes.size),
                reused: 1n,
                row_count: reusable.rowCount,
                columns_json: reusable.columns,
                row_checksum: rowChecksum,
              },
            );
            await insertInternalRow(
              transaction,
              PREPARED_STORAGE.tables.bindings,
              {
                source_key: source.key,
                selection_key: selection.key,
                capture_id: reusable.captureId,
                display_name: source.bytes.name,
              },
            );
          },
        );
        reusedSource = true;
        continue;
      }
      const captureId = `CAPTURE-${globalThis.crypto.randomUUID()}`;
      const reader = await selection.open({
        signal: options.signal,
        onProgress: options.onProgress,
      });
      let rowCount = 0;
      let lastSourceRow = 0;
      const rowChecksum = createCaptureRowChecksum();
      try {
        const columns = await readAndClose(reader, async () => {
          const columnNames = [...reader.columns];
          validateReaderColumns({
            source: source.key,
            selection: selection.key,
            columns: columnNames,
            profile: options.profile,
          });
          const declaredColumns = new Set(columnNames);
          const profiles = new Map<string, ColumnProfile>(
            columnNames.map((column) => [column, emptyProfile()]),
          );
          for await (const batch of reader.batches({
            batchSize: CAPTURE_BATCH_ROWS,
            signal: options.signal,
            onProgress: options.onProgress,
          })) {
            throwIfAborted(options.signal, "db.import.prepare");
            const rows: Array<readonly EngineValue[]> = [];
            for (const row of batch) {
              if (
                !Number.isSafeInteger(row.sourceRow) ||
                row.sourceRow < 1 ||
                row.sourceRow <= lastSourceRow
              ) {
                throw databaseError(
                  "DB_INVALID_SOURCE_ROW",
                  "Imported source rows must have unique positive row numbers in ascending order.",
                  {
                    source: source.key,
                    selection: selection.key,
                    sourceRow: row.sourceRow,
                    previousSourceRow: lastSourceRow,
                  },
                );
              }
              lastSourceRow = row.sourceRow;
              const cellColumns = Object.keys(row.cells);
              const unexpectedColumn = cellColumns.find(
                (column) => !declaredColumns.has(column),
              );
              if (unexpectedColumn !== undefined) {
                throw databaseError(
                  "DB_INVALID_SOURCE_COLUMN",
                  `Source "${source.key}" selection "${selection.key}" returned an undeclared column. Make the reader's column list match its row cells before importing.`,
                  {
                    source: source.key,
                    selection: selection.key,
                    sourceRow: row.sourceRow,
                    column: unexpectedColumn,
                  },
                );
              }
              for (const column of cellColumns) {
                const cell = row.cells[column]!;
                try {
                  assertValidImportDateCell(cell, "source");
                  assertValidImportNumberCell(cell, "source");
                } catch (error) {
                  if (
                    !isConsultChimpsError(error) ||
                    (error.code !== "DB_INVALID_SOURCE_DATE" &&
                      error.code !== "DB_INVALID_SOURCE_NUMBER")
                  ) {
                    throw error;
                  }
                  throw databaseError(
                    error.code,
                    error.message,
                    {
                      source: source.key,
                      selection: selection.key,
                      sourceRow: row.sourceRow,
                      column,
                    },
                    error,
                  );
                }
              }
              for (const column of columnNames) {
                const cell = row.cells[column] ?? { kind: "blank" };
                const profile = profiles.get(column);
                if (profile !== undefined) addToProfile(profile, cell);
              }
              const sourceRow = BigInt(row.sourceRow);
              const valuesJson = JSON.stringify(
                Object.fromEntries(
                  cellColumns.map((column) => [column, row.cells[column]]),
                ),
              );
              rowChecksum.update(sourceRow, valuesJson);
              rows.push([captureId, sourceRow, valuesJson]);
            }
            await preparedEngine.transaction(async (transaction) => {
              await assertImportBatchWritable(transaction);
              await transaction.bulkInsert({
                table: PREPARED_ROW_TABLE,
                columns: ["capture_id", "source_row", "values_json"],
                rows,
                signal: options.signal,
              });
            });
            rowCount += batch.length;
            rowsCaptured += batch.length;
          }
          return inferColumns(columnNames, profiles);
        });
        await source.verifyUnchanged?.();
        await updatePreparedCaptureMetadata(
          options.prepared,
          async (transaction) => {
            await insertInternalRow(
              transaction,
              PREPARED_STORAGE.tables.captures,
              {
                capture_id: captureId,
                source_file_id: null,
                source_key: source.key,
                display_name: source.bytes.name,
                selection_key: selection.key,
                selection_label: selection.label,
                reader_version: source.readerVersion,
                content_hash: contentHash,
                byte_count: BigInt(source.bytes.size),
                reused: 0n,
                row_count: BigInt(rowCount),
                columns_json: JSON.stringify(columns),
                row_checksum: rowChecksum.digest(),
              },
            );
            await insertInternalRow(
              transaction,
              PREPARED_STORAGE.tables.bindings,
              {
                source_key: source.key,
                selection_key: selection.key,
                capture_id: captureId,
                display_name: source.bytes.name,
              },
            );
          },
        );
      } catch (error) {
        if (rowCount === 0) throw error;
        try {
          await preparedEngine.transaction(async (transaction) => {
            await assertImportBatchWritable(transaction);
            await transaction.execute(
              `DELETE FROM ${PREPARED_ROW_TABLE} WHERE capture_id = ?`,
              [captureId],
            );
          });
        } catch (cleanupError) {
          throw new AggregateError(
            [error, cleanupError],
            "Import preparation failed and its staged rows could not be removed.",
            { cause: cleanupError },
          );
        }
        throw error;
      }
      capturedSource = true;
    }
    if (capturedSource) sourcesRead += 1;
    if (reusedSource) sourcesReused += 1;
  }
  const reviewSnapshot = await readPreparedReviewSnapshot(
    options.prepared,
    async (review, transaction) => ({
      review,
      captures: await preparedCapturesFromEngine(transaction),
    }),
  );
  const captures = reviewSnapshot.captures;
  const conflicts = await evaluateConflicts(
    options.database,
    options.prepared,
    captures,
    options.profile,
  );
  const prepared = await updatePreparedPlan({
    prepared: options.prepared,
    profile: options.profile,
    conflicts,
    ready: conflicts.length === 0,
    expectedReviewFingerprint: reviewSnapshot.review.prepared.reviewFingerprint,
  });
  await preparedEngine.checkpoint();
  const result: PrepareImportOutcome["result"] = {
    operation: "db.import.prepare",
    artifacts: [],
    warnings: [],
    metrics: {
      sourcesRead,
      sourcesReused,
      rowsCaptured,
      conflicts: conflicts.length,
    },
  };
  if (options.reviewPage === undefined) return { prepared, result };
  return {
    ...(await inspectUpdatedImport({
      database: options.database,
      prepared: options.prepared,
      expected: prepared,
      page: options.reviewPage,
    })),
    result,
  };
}
