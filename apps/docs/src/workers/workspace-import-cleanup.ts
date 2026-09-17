import {
  ConsultChimpsError,
  isConsultChimpsError,
  OwnedResources,
} from "@consultchimps/core";

export interface RetryableCleanupOwner {
  close(): Promise<void>;
}

export interface CleanupFailure {
  readonly error: unknown;
}

export class RetryableCleanupOwners {
  readonly #owners = new OwnedResources<RetryableCleanupOwner>();
  #closing: Promise<void> | undefined;

  get size(): number {
    return this.#owners.size;
  }

  retain(owner: RetryableCleanupOwner): void {
    this.#owners.add(owner);
  }

  close(): Promise<void> {
    if (this.#closing !== undefined) return this.#closing;
    const attempt = this.#owners.close().then((failures) => {
      if (failures.length === 1) throw failures[0]!.error;
      if (failures.length > 1) {
        throw new AggregateError(
          failures.map((failure) => failure.error),
          "Closing retained import resources failed",
        );
      }
    });
    this.#closing = attempt;
    void attempt.then(
      () => {
        if (this.#closing === attempt) this.#closing = undefined;
      },
      () => {
        if (this.#closing === attempt) this.#closing = undefined;
      },
    );
    return attempt;
  }
}

export async function closeAndRetainFailures(
  owners: readonly RetryableCleanupOwner[],
  retained: RetryableCleanupOwners,
): Promise<void> {
  const pending = new OwnedResources(owners);
  const failed = await pending.close();
  for (const failure of failed) {
    retained.retain(failure.resource);
  }
  const failures = failed.map((failure) => failure.error);
  if (failures.length === 1) throw failures[0];
  if (failures.length > 1)
    throw new AggregateError(failures, "Workbook source cleanup failed");
}

export function stagedPrivatePlanCleanup(options: {
  close(): Promise<void>;
  discard(): Promise<void>;
}): RetryableCleanupOwner {
  let closed = false;
  return {
    async close() {
      if (!closed) {
        await options.close();
        closed = true;
      }
      await options.discard();
    },
  };
}

export function importCleanupError(options: {
  readonly operationFailure?: CleanupFailure;
  readonly sourceCleanupFailure?: CleanupFailure;
  readonly planCleanupFailure?: CleanupFailure;
  readonly preparationCompleted: boolean;
}): ConsultChimpsError {
  const causes = [
    ...(options.operationFailure === undefined
      ? []
      : [options.operationFailure.error]),
    ...(options.sourceCleanupFailure === undefined
      ? []
      : [options.sourceCleanupFailure.error]),
    ...(options.planCleanupFailure === undefined
      ? []
      : [options.planCleanupFailure.error]),
  ];
  return new ConsultChimpsError(
    "DB_BROWSER_IMPORT_CLEANUP_REQUIRED",
    `${options.preparationCompleted ? "The batch was prepared, but success was not reported because" : "Preparing the batch failed, and"} its private workbook or batch resources could not finish cleanup. Retry preparing the batch to finish the retained cleanup before opening the sources again.`,
    {
      details: {
        preparationCompleted: options.preparationCompleted,
        sourceCleanupFailed: options.sourceCleanupFailure !== undefined,
        planCleanupFailed: options.planCleanupFailure !== undefined,
        ...(options.operationFailure !== undefined &&
        isConsultChimpsError(options.operationFailure.error)
          ? { primaryCode: options.operationFailure.error.code }
          : {}),
      },
      cause: new AggregateError(causes, "Import preparation cleanup failed"),
    },
  );
}

export function savedPlanCleanupError(options: {
  readonly operationFailures: readonly unknown[];
  readonly cleanupFailures: readonly unknown[];
}): ConsultChimpsError {
  return new ConsultChimpsError(
    "DB_BROWSER_IMPORT_CLEANUP_REQUIRED",
    "Saved reviews could not finish releasing their private resources. Choose Retry saved reviews to finish cleanup before reopening them.",
    {
      details: {
        savedPlanCleanupFailed: true,
        affectedPlans: options.cleanupFailures.length,
      },
      cause: new AggregateError(
        [...options.operationFailures, ...options.cleanupFailures],
        "Saved batch inspection cleanup failed",
      ),
    },
  );
}
