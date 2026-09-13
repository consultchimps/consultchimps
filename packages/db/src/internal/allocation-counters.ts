import { valueAsPositiveBigInt, valueAsString } from "../database.js";
import { databaseError } from "../errors.js";
import {
  APPLICATION_TABLE,
  CAPTURE_TABLE,
  COUNTERS_TABLE,
  DELIVERY_TABLE,
  SOURCE_FILE_TABLE,
  TABLE_REGISTRY_TABLE,
} from "../metadata.js";
import {
  formatRecordId,
  quoteIdentifier,
  type TableSchema,
} from "../schema.js";
import { parseStoredTableSchema } from "./database-layout.js";
import type { EngineTransaction } from "./engine.js";

type MetadataCounter = "source_file" | "capture" | "import" | "delivery";

const METADATA_ALLOCATIONS: readonly {
  readonly counter: MetadataCounter;
  readonly table: string;
  readonly column: string;
  readonly prefix: string;
}[] = [
  {
    counter: "source_file",
    table: SOURCE_FILE_TABLE,
    column: "source_file_id",
    prefix: "SRC",
  },
  {
    counter: "capture",
    table: CAPTURE_TABLE,
    column: "capture_id",
    prefix: "CAP",
  },
  {
    counter: "import",
    table: APPLICATION_TABLE,
    column: "import_id",
    prefix: "IMP",
  },
  {
    counter: "delivery",
    table: DELIVERY_TABLE,
    column: "delivery_id",
    prefix: "DEL",
  },
] as const;

function corruptCounter(counter: string, table: string): never {
  throw databaseError(
    "DB_CORRUPT_DATABASE",
    "An allocation counter conflicts with stored identifiers. Restore a verified database copy before importing or recording another delivery.",
    { counter, table },
  );
}

async function readCounter(
  transaction: EngineTransaction,
  counter: string,
): Promise<bigint> {
  const rows = await transaction.query(
    `SELECT next_value FROM ${COUNTERS_TABLE} WHERE counter_name = ?`,
    [counter],
  );
  if (rows.length !== 1) return corruptCounter(counter, COUNTERS_TABLE);
  return valueAsPositiveBigInt(rows[0]?.["next_value"], `${counter} counter`);
}

export async function assertMetadataAllocationCounters(
  transaction: EngineTransaction,
  counters: readonly MetadataCounter[] = METADATA_ALLOCATIONS.map(
    ({ counter }) => counter,
  ),
): Promise<void> {
  for (const allocation of METADATA_ALLOCATIONS) {
    if (!counters.includes(allocation.counter)) continue;
    const next = await readCounter(transaction, allocation.counter);
    let cursor: string | undefined;
    while (true) {
      const column = quoteIdentifier(allocation.column);
      const rows = await transaction.query(
        `SELECT ${column} AS allocated_id FROM ${allocation.table}${cursor === undefined ? "" : ` WHERE ${column} > ?`} ORDER BY ${column} LIMIT 1000`,
        cursor === undefined ? [] : [cursor],
      );
      for (const row of rows) {
        const id = valueAsString(row["allocated_id"], "allocated identifier");
        const prefix = `${allocation.prefix}-`;
        const suffix = id.slice(prefix.length);
        if (!id.startsWith(prefix) || !/^\d+$/u.test(suffix)) {
          return corruptCounter(allocation.counter, allocation.table);
        }
        const value = BigInt(suffix);
        if (
          value < 1n ||
          value >= next ||
          id !== `${prefix}${value.toString().padStart(6, "0")}`
        ) {
          return corruptCounter(allocation.counter, allocation.table);
        }
        cursor = id;
      }
      if (rows.length < 1000) break;
    }
  }
}

export async function assertRowAllocationCounters(
  transaction: EngineTransaction,
): Promise<void> {
  const nextImportedRow = await readCounter(transaction, "imported_row");
  const registered = await transaction.query(
    `SELECT table_name, schema_json, next_record_id FROM ${TABLE_REGISTRY_TABLE} ORDER BY table_name`,
  );
  for (const entry of registered) {
    const table = valueAsString(entry["table_name"], "table name");
    const nextRecord = valueAsPositiveBigInt(
      entry["next_record_id"],
      "Record ID counter",
    );
    const rows = await transaction.query(
      `SELECT _imported_row_id, record_id FROM ${quoteIdentifier(table)} WHERE _imported_row_id IS NOT NULL ORDER BY _imported_row_id DESC LIMIT 1`,
    );
    const latest = rows[0];
    if (latest === undefined) continue;
    const importedRow = valueAsPositiveBigInt(
      latest["_imported_row_id"],
      "imported row ID",
    );
    if (nextImportedRow <= importedRow)
      return corruptCounter("imported_row", table);
    const schema = parseStoredTableSchema(
      valueAsString(entry["schema_json"], "table schema"),
      table,
    );
    const id = valueAsString(latest["record_id"], "Record ID");
    const prefix = `${schema.recordId.prefix}${schema.recordId.separator ?? "-"}`;
    const suffix = id.slice(prefix.length);
    if (!id.startsWith(prefix) || !/^\d+$/u.test(suffix))
      return corruptCounter("record_id", table);
    const record = BigInt(suffix);
    if (
      record < 1n ||
      nextRecord <= record ||
      formatRecordId(schema.recordId, record) !== id
    ) {
      return corruptCounter("record_id", table);
    }
  }
}

export async function assertRecordAllocationAvailable(options: {
  readonly transaction: EngineTransaction;
  readonly schema: TableSchema;
  readonly nextRecord: bigint;
  readonly count: bigint;
}): Promise<void> {
  const { transaction, schema, nextRecord, count } = options;
  const last = nextRecord + count - 1n;
  let first = nextRecord;
  while (first <= last) {
    const width = Math.max(schema.recordId.padding, first.toString().length);
    const widthEnd = 10n ** BigInt(width) - 1n;
    const end = last < widthEnd ? last : widthEnd;
    const lower = formatRecordId(schema.recordId, first);
    const upper = formatRecordId(schema.recordId, end);
    let cursor: string | undefined;
    while (true) {
      const rows = await transaction.query(
        `SELECT record_id FROM ${quoteIdentifier(schema.name)} WHERE record_id >= ? AND record_id <= ?${cursor === undefined ? "" : " AND record_id > ?"} ORDER BY record_id LIMIT 1000`,
        cursor === undefined ? [lower, upper] : [lower, upper, cursor],
      );
      for (const row of rows) {
        const id = valueAsString(row["record_id"], "Record ID");
        const prefix = `${schema.recordId.prefix}${schema.recordId.separator ?? "-"}`;
        const suffix = id.slice(prefix.length);
        if (id.startsWith(prefix) && /^\d+$/u.test(suffix)) {
          const value = BigInt(suffix);
          if (
            value >= first &&
            value <= end &&
            formatRecordId(schema.recordId, value) === id
          ) {
            return corruptCounter("record_id", schema.name);
          }
        }
        cursor = id;
      }
      if (rows.length < 1000) break;
    }
    first = end + 1n;
  }
}
