import { sha256 } from "@noble/hashes/sha2.js";
import { bytesToHex } from "@noble/hashes/utils.js";

import { canonicalJson } from "../internal/json.js";
import { databaseError } from "../errors.js";
import { identifierKey, type TableSchema } from "../schema.js";
import type { ColumnRoute } from "./types.js";

const APPLICATION_KEY_PREFIX = "mapping-v1:";

export function effectiveMappingKey(options: {
  readonly captureId: string;
  readonly tableName: string;
  readonly schema: TableSchema;
  readonly columns: readonly ColumnRoute[];
}): string {
  const targets = new Map(
    options.schema.columns.map((column) => [
      identifierKey(column.name),
      column.name,
    ]),
  );
  const mappings = options.columns
    .map((column) => {
      const target = targets.get(identifierKey(column.target));
      if (target === undefined) {
        throw databaseError(
          "DB_STALE_IMPORT_PLAN",
          `The destination column "${column.target}" no longer exists in table "${options.tableName}".`,
          { table: options.tableName, column: column.target },
        );
      }
      return { source: column.source, target: identifierKey(target) };
    })
    .sort(
      (left, right) =>
        (left.target < right.target
          ? -1
          : left.target > right.target
            ? 1
            : 0) ||
        (left.source < right.source ? -1 : left.source > right.source ? 1 : 0),
    );
  return `${APPLICATION_KEY_PREFIX}${bytesToHex(
    sha256(
      new TextEncoder().encode(
        canonicalJson([
          options.captureId,
          identifierKey(options.tableName),
          mappings,
        ]),
      ),
    ),
  )}`;
}
