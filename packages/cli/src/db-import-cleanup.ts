import { ConsultChimpsError } from "@consultchimps/core";

export type CliImportOutcome =
  | { readonly status: "completed" }
  | { readonly status: "failed"; readonly error: unknown };

interface CliImportCleanupOptions {
  readonly outcome: CliImportOutcome;
  readonly temporaryPath?: string | undefined;
  readonly closePrepared?: (() => Promise<void>) | undefined;
  readonly closeInputs?: (() => Promise<void>) | undefined;
  readonly closeDatabase: () => Promise<void>;
  readonly removeTemporary: () => Promise<void>;
}

export async function finishCliImport(
  options: CliImportCleanupOptions,
): Promise<void> {
  const [prepared, inputs, database] = await Promise.allSettled([
    (async () => options.closePrepared?.())(),
    (async () => options.closeInputs?.())(),
    (async () => options.closeDatabase())(),
  ]);
  const preparedCloseFailed = prepared.status === "rejected";
  const inputCloseFailed = inputs.status === "rejected";
  const databaseCloseFailed = database.status === "rejected";
  let removal:
    | { readonly status: "completed" }
    | { readonly status: "skipped" }
    | { readonly status: "failed"; readonly error: unknown } = {
    status: "completed",
  };
  if (options.temporaryPath !== undefined) {
    if (preparedCloseFailed) {
      removal = { status: "skipped" };
    } else {
      try {
        await options.removeTemporary();
      } catch (error) {
        removal = { status: "failed", error };
      }
    }
  }

  const cleanupFailures = [prepared, inputs, database].flatMap((result) =>
    result.status === "rejected" ? [result.reason] : [],
  );
  if (removal.status === "failed") cleanupFailures.push(removal.error);
  if (cleanupFailures.length === 0 && removal.status !== "skipped") {
    if (options.outcome.status === "failed") throw options.outcome.error;
    return;
  }

  const temporaryGuidance =
    options.temporaryPath === undefined
      ? ""
      : removal.status === "completed"
        ? ` The private import-plan directory at "${options.temporaryPath}" was removed.`
        : ` The private import-plan directory may remain at "${options.temporaryPath}". Restart the process before inspecting or removing it if its plan handle could not be closed.`;
  const operationCompleted = options.outcome.status === "completed";
  const outcomeGuidance = operationCompleted
    ? "The database operation completed, but its local resources could not finish closing. Inspect the committed result before deciding whether to retry."
    : "The database import failed, and its local resources could not finish closing.";
  throw new ConsultChimpsError(
    "CLI_DB_IMPORT_CLEANUP_REQUIRED",
    `${outcomeGuidance}${temporaryGuidance} Resolve the reported cleanup issue before continuing.`,
    {
      details: {
        ...(options.temporaryPath === undefined
          ? {}
          : { temporaryPath: options.temporaryPath }),
        preparedCloseFailed,
        inputCloseFailed,
        databaseCloseFailed,
        removalFailed: removal.status === "failed",
        removalSkipped: removal.status === "skipped",
        operationCompleted,
      },
      cause: new AggregateError(
        [
          ...(options.outcome.status === "failed"
            ? [options.outcome.error]
            : []),
          ...cleanupFailures,
        ],
        "Database import operation and cleanup failed",
      ),
    },
  );
}
