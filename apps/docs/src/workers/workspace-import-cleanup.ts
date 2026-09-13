import { ConsultChimpsError, isConsultChimpsError } from "@consultchimps/core";

import { closeTrackedResources } from "../lib/workspace-replacement";

export interface RetryableCleanupOwner {
  close(): Promise<void>;
}

export interface CleanupFailure {
  readonly error: unknown;
}

export class RetryableCleanupOwners {
  readonly #owners = new Map<symbol, RetryableCleanupOwner>();
  #closing: Promise<void> | undefined;

  get size(): number {
    return this.#owners.size;
  }

  retain(owner: RetryableCleanupOwner): void {
    this.#owners.set(Symbol(), owner);
  }

  close(): Promise<void> {
    if (this.#closing !== undefined) return this.#closing;
    const attempt = closeTrackedResources({
      resources: this.#owners,
      close: async (owner) => owner.close(),
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
  const results = await Promise.allSettled(
    owners.map(async (owner) => owner.close()),
  );
  const failures: unknown[] = [];
  for (const [index, result] of results.entries()) {
    if (result.status === "fulfilled") continue;
    retained.retain(owners[index]!);
    failures.push(result.reason);
  }
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
    `${options.preparationCompleted ? "The import was prepared, but success was not reported because" : "Preparing the import failed, and"} its private workbook or plan resources could not finish cleanup. Retry preparing the import to finish the retained cleanup before opening the sources again.`,
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
