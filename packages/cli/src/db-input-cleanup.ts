import { ConsultChimpsError, OwnedResources } from "@consultchimps/core";

interface DbInputCleanupResources {
  readonly workbooks: readonly (() => Promise<void>)[];
  readonly files: readonly (() => Promise<void>)[];
  readonly scratch: () => Promise<void>;
}

interface CleanupFailure {
  readonly stage: "workbook" | "file" | "scratch";
  readonly error: unknown;
}

async function closeStage(
  stage: CleanupFailure["stage"],
  resources: OwnedResources<CleanupOwner>,
): Promise<readonly CleanupFailure[]> {
  const failures = await resources.close();
  return failures.map((failure) => ({ stage, error: failure.error }));
}

interface CleanupOwner {
  close(): Promise<void>;
}

function cleanupOwners(
  closes: readonly (() => Promise<void>)[],
): OwnedResources<CleanupOwner> {
  return new OwnedResources(closes.map((close) => ({ close })));
}

export function createDbInputCloser(
  resources: DbInputCleanupResources,
): () => Promise<void> {
  const workbooks = cleanupOwners(resources.workbooks);
  const files = cleanupOwners(resources.files);
  const scratch = cleanupOwners([resources.scratch]);
  let closed = false;
  let closing: Promise<void> | undefined;
  return async () => {
    if (closing !== undefined) return closing;
    if (closed) return;
    closing = (async () => {
      const failures: CleanupFailure[] = [];
      failures.push(
        ...(await closeStage("workbook", workbooks)),
        ...(await closeStage("file", files)),
        ...(await closeStage("scratch", scratch)),
      );
      if (failures.length > 0) {
        throw new ConsultChimpsError(
          "CLI_DB_INPUT_CLEANUP_REQUIRED",
          "The workbook inputs could not finish releasing their local files and scratch storage. Restart the process before inspecting or removing retained temporary data.",
          {
            details: {
              workbookCloseFailures: failures.filter(
                (failure) => failure.stage === "workbook",
              ).length,
              fileCloseFailures: failures.filter(
                (failure) => failure.stage === "file",
              ).length,
              scratchCloseFailures: failures.filter(
                (failure) => failure.stage === "scratch",
              ).length,
            },
            cause: new AggregateError(
              failures.map((failure) => failure.error),
              "Workbook input cleanup failed",
            ),
          },
        );
      }
      closed = true;
    })();
    try {
      await closing;
    } finally {
      closing = undefined;
    }
  };
}

export function dbInputInitializationCleanupError(
  primary: unknown,
  cleanup: unknown,
): ConsultChimpsError {
  const cleanupCauses =
    cleanup instanceof AggregateError
      ? cleanup.errors
      : cleanup instanceof Error && cleanup.cause instanceof AggregateError
        ? cleanup.cause.errors
        : [cleanup];
  return new ConsultChimpsError(
    "CLI_DB_INPUT_CLEANUP_REQUIRED",
    "Opening the workbook inputs failed, and their local files or scratch storage could not be released. Restart the process before inspecting or removing retained temporary data.",
    {
      details: {
        ...(cleanup instanceof ConsultChimpsError
          ? { cleanupCode: cleanup.code, ...cleanup.details }
          : {}),
        initializationFailed: true,
      },
      cause: new AggregateError(
        [primary, ...cleanupCauses],
        "Workbook input initialization and cleanup failed",
      ),
    },
  );
}
