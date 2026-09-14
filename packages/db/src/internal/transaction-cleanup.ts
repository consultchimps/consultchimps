import type { ConsultChimpsError } from "@consultchimps/core";

import { databaseError } from "../errors.js";
import type { DatabaseFormat } from "../schema.js";

type RollbackOutcome =
  | { readonly state: "rolled-back"; readonly error: unknown }
  | { readonly state: "unresolved"; readonly error: ConsultChimpsError };

export async function rollbackAfterFailure(options: {
  readonly cause: unknown;
  readonly rollback: () => void | Promise<void>;
  readonly format: DatabaseFormat;
}): Promise<RollbackOutcome> {
  try {
    await options.rollback();
    return { state: "rolled-back", error: options.cause };
  } catch (rollbackCause) {
    return {
      state: "unresolved",
      error: databaseError(
        "DB_TRANSACTION_ROLLBACK_FAILED",
        "The database operation failed, and its transaction could not be rolled back. Close and reopen the database, then inspect its state before retrying the operation.",
        {
          format: options.format,
          rollbackFailed: true,
          transactionState: "unknown",
        },
        new AggregateError(
          [options.cause, rollbackCause],
          "The database operation and transaction rollback failed.",
        ),
      ),
    };
  }
}
