import {
  engineOf,
  valueAsBigInt,
  valueAsString,
  type Database,
} from "../database.js";
import { databaseError } from "../errors.js";
import type { EngineTransaction } from "../internal/engine.js";
import { canonicalJson } from "../internal/json.js";
import {
  CAPTURE_TABLE,
  COUNTERS_TABLE,
  DATABASE_METADATA_TABLE,
  DELIVERY_MEMBERSHIP_TABLE,
  DELIVERY_TABLE,
} from "../metadata.js";
import type { DeliveryContext, DeliveryPage, DeliveryRecord } from "./types.js";
import { parseDeliveryContext } from "../validators.js";
import type { DeliveryResult } from "./types.js";

async function allocateDelivery(
  transaction: EngineTransaction,
): Promise<string> {
  const rows = await transaction.query(
    `SELECT next_value FROM ${COUNTERS_TABLE} WHERE counter_name = ?`,
    ["delivery"],
  );
  const next = valueAsBigInt(rows[0]?.["next_value"], "delivery counter");
  await transaction.execute(
    `UPDATE ${COUNTERS_TABLE} SET next_value = ? WHERE counter_name = ?`,
    [next + 1n, "delivery"],
  );
  return `DEL-${next.toString().padStart(6, "0")}`;
}

export async function recordDelivery(options: {
  readonly database: Database;
  readonly captureIds: readonly string[];
  readonly context: DeliveryContext;
  readonly requestId: string;
}): Promise<DeliveryResult> {
  if (options.requestId.trim().length === 0) {
    throw databaseError(
      "DB_DELIVERY_REQUEST_ID_REQUIRED",
      "Give the delivery a request ID so retrying it cannot create a duplicate event.",
    );
  }
  if (options.captureIds.length === 0) {
    throw databaseError(
      "DB_DELIVERY_CAPTURE_REQUIRED",
      "Choose at least one completed capture for this delivery.",
    );
  }
  return engineOf(options.database).transaction(async (transaction) => {
    const uniqueCaptures = [...new Set(options.captureIds)].sort();
    const existing = await transaction.query(
      `SELECT delivery_id, context_json FROM ${DELIVERY_TABLE} WHERE request_id = ?`,
      [options.requestId],
    );
    if (existing[0] !== undefined) {
      const id = valueAsString(existing[0]["delivery_id"], "delivery ID");
      const memberships = await transaction.query(
        `SELECT capture_id FROM ${DELIVERY_MEMBERSHIP_TABLE} WHERE delivery_id = ? ORDER BY capture_id`,
        [id],
      );
      const delivery: DeliveryRecord = {
        id,
        requestId: options.requestId,
        context: parseDeliveryContext(
          JSON.parse(
            valueAsString(existing[0]["context_json"], "delivery context"),
          ),
        ),
        captureIds: memberships.map((row) =>
          valueAsString(row["capture_id"], "capture ID"),
        ),
      };
      if (
        canonicalJson(delivery.context) !== canonicalJson(options.context) ||
        JSON.stringify(delivery.captureIds) !== JSON.stringify(uniqueCaptures)
      ) {
        throw databaseError(
          "DB_REQUEST_ID_CONFLICT",
          "This request ID was already used with different delivery details. Choose a new request ID.",
          { requestId: options.requestId },
        );
      }
      return {
        operation: "db.delivery.record",
        artifacts: [],
        warnings: [],
        metrics: {
          deliveriesRecorded: 0,
          capturesLinked: delivery.captureIds.length,
        },
        delivery,
      };
    }
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
    await transaction.execute(
      `INSERT INTO ${DELIVERY_TABLE} VALUES (?, ?, ?)`,
      [id, options.requestId, canonicalJson(options.context)],
    );
    for (const captureId of uniqueCaptures) {
      await transaction.execute(
        `INSERT INTO ${DELIVERY_MEMBERSHIP_TABLE} VALUES (?, ?)`,
        [id, captureId],
      );
    }
    await transaction.execute(
      `UPDATE ${DATABASE_METADATA_TABLE} SET revision = revision + 1`,
    );
    const delivery: DeliveryRecord = {
      id,
      requestId: options.requestId,
      context: options.context,
      captureIds: uniqueCaptures,
    };
    return {
      operation: "db.delivery.record",
      artifacts: [],
      warnings: [],
      metrics: {
        deliveriesRecorded: 1,
        capturesLinked: uniqueCaptures.length,
      },
      delivery,
    };
  });
}

export async function listDeliveries(options: {
  readonly database: Database;
  readonly limit: number;
  readonly cursor?: string | undefined;
}): Promise<DeliveryPage> {
  if (
    !Number.isInteger(options.limit) ||
    options.limit < 1 ||
    options.limit > 100
  ) {
    throw databaseError(
      "DB_INVALID_PAGE_SIZE",
      "Choose a delivery page size from 1 to 100.",
    );
  }
  if (options.cursor !== undefined && !/^DEL-\d+$/u.test(options.cursor)) {
    throw databaseError("DB_INVALID_CURSOR", "The delivery cursor is invalid.");
  }
  const engine = engineOf(options.database);
  const rows = await engine.query(
    `SELECT delivery_id, request_id, context_json FROM ${DELIVERY_TABLE}${options.cursor === undefined ? "" : " WHERE delivery_id > ?"} ORDER BY delivery_id LIMIT ?`,
    options.cursor === undefined
      ? [BigInt(options.limit + 1)]
      : [options.cursor, BigInt(options.limit + 1)],
  );
  const deliveries: DeliveryRecord[] = [];
  for (const row of rows.slice(0, options.limit)) {
    const id = valueAsString(row["delivery_id"], "delivery ID");
    const memberships = await engine.query(
      `SELECT capture_id FROM ${DELIVERY_MEMBERSHIP_TABLE} WHERE delivery_id = ? ORDER BY capture_id`,
      [id],
    );
    deliveries.push({
      id,
      requestId: valueAsString(row["request_id"], "delivery request ID"),
      context: parseDeliveryContext(
        JSON.parse(valueAsString(row["context_json"], "delivery context")),
      ),
      captureIds: memberships.map((membership) =>
        valueAsString(membership["capture_id"], "capture ID"),
      ),
    });
  }
  return {
    deliveries,
    ...(rows.length > options.limit
      ? { nextCursor: deliveries[deliveries.length - 1]?.id }
      : {}),
  };
}
