import { describe, expect, it } from "vitest";

import {
  isConsultChimpsError,
  OPERATION_ABORTED,
  throwIfAborted,
  type ConsultChimpsError,
} from "../src/index.js";

describe("throwIfAborted", () => {
  it("does nothing without a signal", () => {
    expect(() => throwIfAborted(undefined, "test.operation")).not.toThrow();
  });

  it("does nothing while the signal is not aborted", () => {
    const controller = new AbortController();
    expect(() =>
      throwIfAborted(controller.signal, "test.operation"),
    ).not.toThrow();
  });

  it("throws a stable structured error once aborted", () => {
    const controller = new AbortController();
    controller.abort(new Error("user cancelled"));

    let thrown: unknown;
    try {
      throwIfAborted(controller.signal, "test.operation");
    } catch (error) {
      thrown = error;
    }

    expect(isConsultChimpsError(thrown)).toBe(true);
    const error = thrown as ConsultChimpsError;
    expect(error.code).toBe(OPERATION_ABORTED);
    expect(error.details).toEqual({ operation: "test.operation" });
    expect(error.cause).toBeInstanceOf(Error);
    expect(error.message).toContain("cancelled");
  });
});
