import { valueAsBigInt } from "../database.js";
import { databaseError } from "../errors.js";
import type { DatabaseEngine, EngineRow } from "../internal/engine.js";
import { quoteIdentifier } from "../schema.js";

const MAX_SOURCE_ROW = BigInt(Number.MAX_SAFE_INTEGER);

export type SourceRowOwner = "prepared-import" | "database";

export interface ValidatedSourceRow {
  readonly row: EngineRow;
  readonly sourceRow: bigint;
}

function invalidSourceRow(
  owner: SourceRowOwner,
  captureId: string,
  sourceRow?: bigint,
  cause?: unknown,
) {
  return databaseError(
    owner === "prepared-import"
      ? "DB_INVALID_PREPARED_IMPORT"
      : "DB_CORRUPT_DATABASE",
    owner === "prepared-import"
      ? "The import plan has invalid source row numbers. Prepare the source again or restore a verified plan copy."
      : "A saved capture has invalid source row numbers. Restore a verified database copy before importing it again.",
    {
      captureId,
      ...(sourceRow === undefined ? {} : { sourceRow: sourceRow.toString() }),
    },
    cause,
  );
}

function validatedSourceRow(
  value: unknown,
  previous: bigint | undefined,
  captureId: string,
  owner: SourceRowOwner,
): bigint {
  let sourceRow: bigint;
  try {
    sourceRow = valueAsBigInt(value, "source row");
  } catch (cause) {
    throw invalidSourceRow(owner, captureId, undefined, cause);
  }
  if (
    sourceRow < 1n ||
    sourceRow > MAX_SOURCE_ROW ||
    (previous !== undefined && sourceRow <= previous)
  ) {
    throw invalidSourceRow(owner, captureId, sourceRow);
  }
  return sourceRow;
}

export async function readSourceRowPage(options: {
  readonly engine: Pick<DatabaseEngine, "query">;
  readonly table: string;
  readonly captureId: string;
  readonly cursor: bigint | undefined;
  readonly limit: number;
  readonly owner: SourceRowOwner;
}): Promise<{
  readonly rows: readonly ValidatedSourceRow[];
  readonly cursor: bigint | undefined;
}> {
  const table = quoteIdentifier(options.table);
  const rows = await options.engine.query(
    options.cursor === undefined
      ? `SELECT source_row, values_json FROM ${table} WHERE capture_id = ? ORDER BY source_row LIMIT ?`
      : `SELECT source_row, values_json FROM ${table} WHERE capture_id = ? AND source_row > ? ORDER BY source_row LIMIT ?`,
    options.cursor === undefined
      ? [options.captureId, BigInt(options.limit)]
      : [options.captureId, options.cursor, BigInt(options.limit)],
  );
  let cursor = options.cursor;
  const validated = rows.map((row) => {
    const sourceRow = validatedSourceRow(
      row["source_row"],
      cursor,
      options.captureId,
      options.owner,
    );
    cursor = sourceRow;
    return { row, sourceRow };
  });
  return { rows: validated, cursor };
}

export function assertCaptureRowCount(options: {
  readonly owner: SourceRowOwner;
  readonly captureId: string;
  readonly expected: bigint;
  readonly actual: bigint;
}): void {
  if (options.actual === options.expected) return;
  throw databaseError(
    options.owner === "prepared-import"
      ? "DB_INVALID_PREPARED_IMPORT"
      : "DB_CORRUPT_DATABASE",
    options.owner === "prepared-import"
      ? "The import plan's captured row count does not match its recorded total. Prepare the source again or restore a verified plan copy."
      : "A saved capture's row count does not match its recorded total. Restore a verified database copy before importing it again.",
    {
      captureId: options.captureId,
      expectedRows: options.expected.toString(),
      actualRows: options.actual.toString(),
    },
  );
}
