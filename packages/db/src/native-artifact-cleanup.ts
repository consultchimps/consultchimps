import { isConsultChimpsError } from "@consultchimps/core";

import { databaseError } from "./errors.js";
import type { DatabaseFormat } from "./schema.js";

export interface NativeTemporaryArtifactCleanup {
  readonly operation: "create" | "export" | "plan" | "prepare";
  readonly stage: "preparation" | "validation" | "publication";
  readonly kind: "database" | "prepared";
  readonly format: DatabaseFormat;
  readonly temporaryPath: string;
  readonly storagePaths: readonly string[];
  readonly cause: unknown;
  readonly close?: readonly (() => Promise<void>)[] | undefined;
  remove(): Promise<void>;
}

export async function failAfterNativeArtifactCleanup(
  cleanup: NativeTemporaryArtifactCleanup,
): Promise<never> {
  const closeResults = await Promise.allSettled(
    (cleanup.close ?? []).map(async (close) => close()),
  );
  const closeFailures = closeResults.flatMap((result) =>
    result.status === "rejected" ? [result.reason] : [],
  );
  const inaccessibleHandle =
    isConsultChimpsError(cleanup.cause) &&
    (cleanup.cause.code === "DB_NATIVE_SQLITE_CLEANUP_REQUIRED" ||
      cleanup.cause.code === "DB_DUCKDB_EXPORT_CLEANUP_FAILED" ||
      cleanup.cause.code === "DB_DUCKDB_OPEN_CLEANUP_FAILED");
  let removalOutcome:
    | { readonly status: "completed" }
    | { readonly status: "skipped" }
    | { readonly status: "failed"; readonly error: unknown };
  if (inaccessibleHandle || closeFailures.length > 0) {
    removalOutcome = { status: "skipped" };
  } else {
    try {
      await cleanup.remove();
      removalOutcome = { status: "completed" };
    } catch (error) {
      removalOutcome = { status: "failed", error };
    }
  }
  const publishedPath =
    isConsultChimpsError(cleanup.cause) &&
    cleanup.cause.details?.["published"] === true
      ? typeof cleanup.cause.details["output"] === "string"
        ? cleanup.cause.details["output"]
        : typeof cleanup.cause.details["publishedPath"] === "string"
          ? cleanup.cause.details["publishedPath"]
          : undefined
      : undefined;
  if (closeFailures.length === 0 && removalOutcome.status === "completed") {
    throw cleanup.cause;
  }

  const closeFailed = inaccessibleHandle || closeFailures.length > 0;
  const removalFailed = removalOutcome.status === "failed";
  const removalSkipped = removalOutcome.status === "skipped";
  const retainedPaths = cleanup.storagePaths.join(" and ");
  const releaseGuidance = closeFailed
    ? " One or more temporary handles could not be confirmed closed. Restart this process before inspecting or removing the files."
    : "";
  const publicationOutcome =
    publishedPath === undefined
      ? ""
      : ` The operation published "${publishedPath}" before the later failure.`;
  throw databaseError(
    "DB_NATIVE_TEMPORARY_CLEANUP_REQUIRED",
    `The database ${cleanup.operation} operation failed during ${cleanup.stage}, and cleanup of its private ${cleanup.kind} artifact did not finish.${publicationOutcome} Temporary files may remain at ${retainedPaths}.${releaseGuidance} Inspect or preserve recoverable data, then remove the private files after resolving the storage issue and before retrying.`,
    {
      operation: cleanup.operation,
      stage: cleanup.stage,
      kind: cleanup.kind,
      format: cleanup.format,
      temporaryPath: cleanup.temporaryPath,
      storagePaths: cleanup.storagePaths,
      closeFailed,
      removalFailed,
      removalSkipped,
      ...(publishedPath === undefined
        ? {}
        : { published: true, publishedPath }),
      ...(isConsultChimpsError(cleanup.cause)
        ? { primaryCode: cleanup.cause.code }
        : {}),
    },
    new AggregateError(
      [
        cleanup.cause,
        ...closeFailures,
        ...(removalOutcome.status !== "failed" ? [] : [removalOutcome.error]),
      ],
      "Native temporary artifact operation and cleanup failed",
    ),
  );
}
