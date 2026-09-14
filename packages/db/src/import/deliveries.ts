import {
  engineOf,
  valueAsBigInt,
  valueAsPositiveBigInt,
  valueAsString,
  type Database,
} from "../database.js";
import { databaseError } from "../errors.js";
import { assertMetadataAllocationCounters } from "../internal/allocation-counters.js";
import { assertNoManagedDatabaseTriggers } from "../internal/database-layout.js";
import type { EngineTransaction } from "../internal/engine.js";
import { canonicalJson } from "../internal/json.js";
import { DATABASE_STORAGE } from "../internal/storage-layouts.js";
import { insertInternalRow } from "../internal/storage-schema.js";
import { parseStoredBatchContext } from "../internal/stored-delivery.js";
import {
  CAPTURE_TABLE,
  COUNTERS_TABLE,
  DATABASE_METADATA_TABLE,
  DELIVERY_MEMBERSHIP_TABLE,
  DELIVERY_TABLE,
  IMPORT_REQUEST_TABLE,
} from "../metadata.js";
import type { BatchContext, BatchHistoryPage, BatchRecord } from "./types.js";
import { parseBatchContext } from "../validators.js";
import type { BatchRecordResult } from "./types.js";

const DELIVERY_PRECEDES =
  "(LENGTH(prior.delivery_id) < LENGTH(membership.delivery_id) OR (LENGTH(prior.delivery_id) = LENGTH(membership.delivery_id) AND prior.delivery_id < membership.delivery_id))";

async function deliveryMemberships(
  reader: Pick<EngineTransaction, "query">,
  batchId: string,
): Promise<{
  readonly captureIds: readonly string[];
  readonly reusedCaptureIds: readonly string[];
}> {
  const rows = await reader.query(
    `SELECT membership.capture_id, CASE WHEN EXISTS (SELECT 1 FROM ${DELIVERY_MEMBERSHIP_TABLE} AS prior WHERE prior.capture_id = membership.capture_id AND ${DELIVERY_PRECEDES}) THEN 1 ELSE 0 END AS reused_capture FROM ${DELIVERY_MEMBERSHIP_TABLE} AS membership WHERE membership.delivery_id = ? ORDER BY membership.capture_id`,
    [batchId],
  );
  const captureIds: string[] = [];
  const reusedCaptureIds: string[] = [];
  for (const row of rows) {
    const captureId = valueAsString(row["capture_id"], "capture ID");
    captureIds.push(captureId);
    if (valueAsBigInt(row["reused_capture"], "capture reuse flag") !== 0n) {
      reusedCaptureIds.push(captureId);
    }
  }
  return { captureIds, reusedCaptureIds };
}

async function allocateDelivery(
  transaction: EngineTransaction,
): Promise<string> {
  await assertMetadataAllocationCounters(transaction, ["delivery"]);
  const rows = await transaction.query(
    `SELECT next_value FROM ${COUNTERS_TABLE} WHERE counter_name = ?`,
    ["delivery"],
  );
  const next = valueAsPositiveBigInt(rows[0]?.["next_value"], "batch counter");
  await transaction.execute(
    `UPDATE ${COUNTERS_TABLE} SET next_value = ? WHERE counter_name = ?`,
    [next + 1n, "delivery"],
  );
  return `DEL-${next.toString().padStart(6, "0")}`;
}

export async function recordBatch(options: {
  readonly database: Database;
  readonly captureIds: readonly string[];
  readonly context: BatchContext;
  readonly requestId: string;
}): Promise<BatchRecordResult> {
  const context = parseBatchContext(options.context);
  if (options.requestId.trim().length === 0) {
    throw databaseError(
      "DB_DELIVERY_REQUEST_ID_REQUIRED",
      "Give the batch a request ID so retrying it cannot create a duplicate event.",
    );
  }
  if (options.captureIds.length === 0) {
    throw databaseError(
      "DB_DELIVERY_CAPTURE_REQUIRED",
      "Choose at least one completed capture for this batch.",
    );
  }
  return engineOf(options.database).transaction(async (transaction) => {
    const uniqueCaptures = [...new Set(options.captureIds)].sort();
    const existing = await transaction.query(
      `SELECT delivery_id, context_json FROM ${DELIVERY_TABLE} WHERE request_id = ?`,
      [options.requestId],
    );
    const existingDelivery =
      existing[0] === undefined
        ? undefined
        : {
            row: existing[0],
            context: parseStoredBatchContext(existing[0]["context_json"]),
          };
    if (existingDelivery !== undefined) {
      const id = valueAsString(existingDelivery.row["delivery_id"], "batch ID");
      const memberships = await deliveryMemberships(transaction, id);
      const delivery: BatchRecord = {
        id,
        requestId: options.requestId,
        context: existingDelivery.context,
        captureIds: memberships.captureIds,
        reusedCaptureIds: memberships.reusedCaptureIds,
      };
      if (
        canonicalJson(delivery.context) !== canonicalJson(context) ||
        JSON.stringify(delivery.captureIds) !== JSON.stringify(uniqueCaptures)
      ) {
        throw databaseError(
          "DB_REQUEST_ID_CONFLICT",
          "This request ID was already used with different batch details. Choose a new request ID.",
          { requestId: options.requestId },
        );
      }
      return {
        databaseWrite: "unchanged",
        operation: "db.import.record",
        artifacts: [],
        warnings: [],
        metrics: {
          batchesRecorded: 0,
          capturesLinked: delivery.captureIds.length,
        },
        batch: delivery,
      };
    }
    const importReceipts = await transaction.query(
      `SELECT request_id FROM ${IMPORT_REQUEST_TABLE} WHERE request_id = ? LIMIT 1`,
      [options.requestId],
    );
    if (importReceipts.length > 0) {
      throw databaseError(
        "DB_REQUEST_ID_CONFLICT",
        "This request ID was already used for an import. Choose a new request ID for the batch.",
        { requestId: options.requestId },
      );
    }
    await assertNoManagedDatabaseTriggers(transaction, options.database.format);
    for (const captureId of uniqueCaptures) {
      const captures = await transaction.query(
        `SELECT capture_id FROM ${CAPTURE_TABLE} WHERE capture_id = ? AND state = 'completed'`,
        [captureId],
      );
      if (captures.length === 0) {
        throw databaseError(
          "DB_CAPTURE_NOT_FOUND",
          `The completed capture "${captureId}" does not exist in this database.`,
          { captureId },
        );
      }
    }
    const id = await allocateDelivery(transaction);
    await insertInternalRow(transaction, DATABASE_STORAGE.tables.deliveries, {
      delivery_id: id,
      request_id: options.requestId,
      context_json: canonicalJson(context),
    });
    for (const captureId of uniqueCaptures) {
      await insertInternalRow(
        transaction,
        DATABASE_STORAGE.tables.deliveryMemberships,
        { delivery_id: id, capture_id: captureId },
      );
    }
    const memberships = await deliveryMemberships(transaction, id);
    await transaction.execute(
      `UPDATE ${DATABASE_METADATA_TABLE} SET revision = revision + 1`,
    );
    const delivery: BatchRecord = {
      id,
      requestId: options.requestId,
      context,
      captureIds: memberships.captureIds,
      reusedCaptureIds: memberships.reusedCaptureIds,
    };
    return {
      databaseWrite: "committed",
      operation: "db.import.record",
      artifacts: [],
      warnings: [],
      metrics: {
        batchesRecorded: 1,
        capturesLinked: uniqueCaptures.length,
      },
      batch: delivery,
    };
  });
}

export async function listBatches(options: {
  readonly database: Database;
  readonly limit: number;
  readonly cursor?: string | undefined;
}): Promise<BatchHistoryPage> {
  if (
    !Number.isInteger(options.limit) ||
    options.limit < 1 ||
    options.limit > 100
  ) {
    throw databaseError(
      "DB_INVALID_PAGE_SIZE",
      "Choose a batch page size from 1 to 100.",
    );
  }
  if (options.cursor !== undefined && !/^DEL-\d+$/u.test(options.cursor)) {
    throw databaseError("DB_INVALID_CURSOR", "The batch cursor is invalid.");
  }
  const engine = engineOf(options.database);
  const cursorClause =
    options.cursor === undefined
      ? ""
      : " WHERE LENGTH(delivery_id) > LENGTH(?) OR (LENGTH(delivery_id) = LENGTH(?) AND delivery_id > ?)";
  const rows = await engine.query(
    `SELECT delivery_id, request_id, context_json FROM ${DELIVERY_TABLE}${cursorClause} ORDER BY LENGTH(delivery_id), delivery_id LIMIT ?`,
    options.cursor === undefined
      ? [BigInt(options.limit + 1)]
      : [
          options.cursor,
          options.cursor,
          options.cursor,
          BigInt(options.limit + 1),
        ],
  );
  const deliveries: BatchRecord[] = [];
  for (const row of rows.slice(0, options.limit)) {
    const id = valueAsString(row["delivery_id"], "batch ID");
    const memberships = await deliveryMemberships(engine, id);
    deliveries.push({
      id,
      requestId: valueAsString(row["request_id"], "batch request ID"),
      context: parseStoredBatchContext(row["context_json"]),
      captureIds: memberships.captureIds,
      reusedCaptureIds: memberships.reusedCaptureIds,
    });
  }
  return {
    batches: deliveries,
    ...(rows.length > options.limit
      ? { nextCursor: deliveries[deliveries.length - 1]?.id }
      : {}),
  };
}
