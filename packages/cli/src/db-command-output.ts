import {
  ConsultChimpsError,
  isConsultChimpsError,
  type OperationResult,
} from "@consultchimps/core";

import { withoutTerminalControlsInProse } from "./text.js";

export interface DbCommandOutput {
  json(): boolean;
  result(value: OperationResult): void;
  data(value: unknown, humanText: string): void;
}

export interface DeferredDbCommandOutput extends DbCommandOutput {
  completedData(value: unknown, humanText: string): void;
  prose(value: string): void;
}

export async function withDeferredDbCommandOutput(
  output: DbCommandOutput,
  work: (deferred: DeferredDbCommandOutput) => Promise<void>,
  recoveryPaths: readonly string[] = [],
): Promise<void> {
  const pending: Array<() => void> = [];
  let completedOperation: string | undefined;
  const deferred: DeferredDbCommandOutput = {
    json: () => output.json(),
    result: (value) => {
      completedOperation = value.operation;
      pending.push(() => output.result(value));
    },
    data: (value, humanText) =>
      pending.push(() => output.data(value, humanText)),
    completedData: (value, humanText) => {
      completedOperation = "db.resolve";
      pending.push(() => output.data(value, humanText));
    },
    prose: (value) => {
      if (!output.json())
        pending.push(() =>
          process.stdout.write(withoutTerminalControlsInProse(value)),
        );
    },
  };

  try {
    await work(deferred);
  } catch (error) {
    if (
      completedOperation === undefined ||
      (isConsultChimpsError(error) &&
        error.details?.["operationCompleted"] === true)
    )
      throw error;
    const locations = recoveryPaths.map((value) => `"${value}"`).join(" and ");
    throw new ConsultChimpsError(
      "CLI_DB_COMMAND_CLEANUP_REQUIRED",
      `The database operation completed, but its local resources could not finish closing.${locations.length === 0 ? "" : ` Inspect the committed result at ${locations} before deciding whether to retry.`} Resolve the reported cleanup issue before continuing.`,
      {
        details: {
          operation: completedOperation,
          operationCompleted: true,
          recoveryPaths,
          ...(isConsultChimpsError(error) ? { primaryCode: error.code } : {}),
        },
        cause: error,
      },
    );
  }
  for (const emit of pending) emit();
}
