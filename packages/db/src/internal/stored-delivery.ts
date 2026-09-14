import { databaseError } from "../errors.js";
import type { BatchContext } from "../import/types.js";
import { parseBatchContext } from "../validators.js";

export function parseStoredBatchContext(value: unknown): BatchContext {
  try {
    if (typeof value !== "string") throw new Error("Expected stored JSON text");
    const parsed: unknown = JSON.parse(value);
    return parseBatchContext(parsed);
  } catch (cause) {
    throw databaseError(
      "DB_CORRUPT_DATABASE",
      "The recorded batch details are damaged. Restore a verified database copy before retrying.",
      undefined,
      cause,
    );
  }
}
