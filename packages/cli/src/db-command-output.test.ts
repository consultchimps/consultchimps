import { describe, expect, test, vi } from "vitest";

import {
  type DbCommandOutput,
  withDeferredDbCommandOutput,
} from "./db-command-output.js";

const operation = {
  artifacts: [],
  metrics: {},
  operation: "db.test",
  warnings: [],
};

function output(json = true): {
  readonly events: string[];
  readonly output: DbCommandOutput;
} {
  const events: string[] = [];
  return {
    events,
    output: {
      json: () => json,
      result: () => events.push("result"),
      data: () => events.push("data"),
    },
  };
}

describe("deferred database command output", () => {
  test("emits structured output after required cleanup succeeds", async () => {
    const target = output();

    await withDeferredDbCommandOutput(target.output, async (deferred) => {
      deferred.result(operation);
      target.events.push("operation");
      await Promise.resolve();
      target.events.push("cleanup");
    });

    expect(target.events).toEqual(["operation", "cleanup", "result"]);
  });

  test("discards success output when cleanup fails", async () => {
    const target = output();
    const cleanupFailure = new Error("Injected cleanup failure");

    await expect(
      withDeferredDbCommandOutput(
        target.output,
        async (deferred) => {
          deferred.result(operation);
          throw cleanupFailure;
        },
        ["review.sqlite"],
      ),
    ).rejects.toMatchObject({
      cause: cleanupFailure,
      code: "CLI_DB_COMMAND_CLEANUP_REQUIRED",
      details: {
        operation: "db.test",
        operationCompleted: true,
        recoveryPaths: ["review.sqlite"],
      },
    });

    expect(target.events).toEqual([]);
  });

  test("does not describe a failed read-only response as committed", async () => {
    const target = output();
    const closeFailure = new Error("Injected read-only close failure");

    await expect(
      withDeferredDbCommandOutput(target.output, async (deferred) => {
        deferred.data({ tables: [] }, "No tables\n");
        throw closeFailure;
      }),
    ).rejects.toBe(closeFailure);

    expect(target.events).toEqual([]);
  });

  test("keeps command prose behind cleanup and omits it in JSON mode", async () => {
    const write = vi.spyOn(process.stdout, "write").mockReturnValue(true);
    const human = output(false);
    const json = output(true);

    await withDeferredDbCommandOutput(human.output, async (deferred) => {
      deferred.prose("Finished[31m safely\n");
      expect(write).not.toHaveBeenCalled();
    });
    await withDeferredDbCommandOutput(json.output, async (deferred) => {
      deferred.prose("Hidden in JSON mode\n");
    });

    expect(write).toHaveBeenCalledOnce();
    expect(write).toHaveBeenCalledWith("Finished\\u001B[31m safely\n");
    write.mockRestore();
  });
});
