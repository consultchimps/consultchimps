import type { ConsultChimpsError, OperationResult } from "@consultchimps/core";

import type { Database } from "./database.js";
import { databaseError } from "./errors.js";

export interface DatabaseWriteResult {
  readonly databaseWrite: "unchanged" | "committed";
}

export type DatabaseCheckpoint =
  | { readonly state: "checkpoint-completed" }
  | {
      readonly state: "checkpoint-required";
      readonly code: "DB_CHECKPOINT_REQUIRED";
      readonly message: string;
      readonly error: ConsultChimpsError;
    };

export interface CheckpointedDatabaseWrite<T extends DatabaseWriteResult> {
  readonly result: T;
  readonly checkpoint: DatabaseCheckpoint;
}

export async function checkpointDatabaseWrite<
  T extends DatabaseWriteResult & Pick<OperationResult, "operation">,
>(options: {
  readonly database: Database;
  readonly result: T;
}): Promise<CheckpointedDatabaseWrite<T>> {
  try {
    await options.database.checkpoint();
    return {
      result: options.result,
      checkpoint: { state: "checkpoint-completed" },
    };
  } catch (cause) {
    const error = databaseError(
      "DB_CHECKPOINT_REQUIRED",
      "The database operation completed, but saving its latest state could not be confirmed. Keep the database open and retry the checkpoint.",
      {
        databaseId: options.database.id,
        operation: options.result.operation,
        databaseWrite: options.result.databaseWrite,
      },
      cause,
    );
    return {
      result: options.result,
      checkpoint: {
        state: "checkpoint-required",
        code: "DB_CHECKPOINT_REQUIRED",
        message: error.message,
        error,
      },
    };
  }
}
