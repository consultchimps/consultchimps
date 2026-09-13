import { databaseError } from "./errors.js";
import type { DatabaseFormat } from "./schema.js";

export interface BrowserConversionArtifact {
  readonly name: string;
  readonly format: DatabaseFormat;
  readonly directory: string;
  readonly storageNames?: readonly string[] | undefined;
  close(): Promise<void>;
  remove(): Promise<void>;
}

export interface BrowserConversionResult<T> {
  readonly value: T;
  readonly warnings: readonly string[];
}

function retainedArtifactGuidance(artifact: BrowserConversionArtifact): string {
  if (artifact.format === "sqlite") {
    return `The logical working copy "${artifact.name}" may remain in the SQLite SAH pool at "${artifact.directory}". Configure BrowserDatabaseRuntime with that same pool and use openDatabase({ name: "${artifact.name}" }) to inspect or export it after resolving the storage issue.`;
  }
  return `Files may remain in the ${artifact.directory} origin-private directory as ${artifact.storageNames?.join(" and ") ?? artifact.name}. Use BrowserDatabaseRuntime.openDatabase({ name: "${artifact.name}" }) to inspect or export recoverable data after resolving the storage issue.`;
}

export async function useBrowserConversionArtifact<T>(
  artifact: BrowserConversionArtifact,
  work: () => Promise<T>,
): Promise<BrowserConversionResult<T>> {
  let outcome:
    | { readonly status: "completed"; readonly value: T }
    | { readonly status: "failed"; readonly error: unknown };
  try {
    outcome = { status: "completed", value: await work() };
  } catch (error) {
    outcome = { status: "failed", error };
  }

  let cleanupOutcome:
    | { readonly status: "completed" }
    | { readonly status: "failed"; readonly error: unknown };
  try {
    await artifact.close();
    cleanupOutcome = { status: "completed" };
  } catch (error) {
    cleanupOutcome = { status: "failed", error };
  }
  if (cleanupOutcome.status === "completed") {
    try {
      await artifact.remove();
      cleanupOutcome = { status: "completed" };
    } catch (error) {
      cleanupOutcome = { status: "failed", error };
    }
  }

  if (outcome.status === "failed") {
    if (cleanupOutcome.status === "completed") throw outcome.error;
    throw databaseError(
      "DB_BROWSER_CONVERSION_CLEANUP_REQUIRED",
      `The browser export failed, and its temporary ${artifact.format} conversion working copy could not be removed. ${retainedArtifactGuidance(artifact)} Retry the export after resolving the storage issue.`,
      {
        directory: artifact.directory,
        conversionName: artifact.name,
        format: artifact.format,
        ...(artifact.storageNames === undefined
          ? {}
          : { storageNames: artifact.storageNames }),
      },
      new AggregateError(
        [outcome.error, cleanupOutcome.error],
        "Browser conversion and cleanup failed",
      ),
    );
  }

  return {
    value: outcome.value,
    warnings:
      cleanupOutcome.status === "completed"
        ? []
        : [
            `The export completed, but browser storage could not remove its temporary ${artifact.format} conversion working copy. ${retainedArtifactGuidance(artifact)}`,
          ],
  };
}
