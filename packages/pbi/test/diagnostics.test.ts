import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { ConsultChimpsError } from "@consultchimps/core";
import { noExportableTables, runtimeUnavailable } from "../src/errors.js";
import { exportPbiTables, readPbiTables } from "../src/pipeline.js";
import { FIXTURES, FETCHED } from "./oracle.js";
import { existsSync } from "node:fs";

/**
 * ADR 0004 Decision 4's prohibition on document contents in diagnostics.
 * Messages, structured details and exposed causes carry no model-provided name,
 * identifier, DAX, cell value or parser text. Named exclusions belong only in a
 * successful export's manifest.
 */

const marker = "CONFIDENTIAL_SYNTHETIC_MARKER";
const container = new Uint8Array(
  readFileSync(path.join(FIXTURES, "a-2018-fuzzy.pbix")),
);

/** Everything a host could serialize from one of our errors. */
function serialized(error: ConsultChimpsError): string {
  return JSON.stringify({
    code: error.code,
    message: error.message,
    details: error.details,
    cause: error.cause === undefined ? null : String(error.cause),
    stack: "",
  });
}

async function refusal(
  run: () => Promise<unknown>,
): Promise<ConsultChimpsError> {
  try {
    await run();
  } catch (error) {
    if (error instanceof ConsultChimpsError) return error;
    throw error;
  }
  throw new Error("expected a refusal");
}

describe("refusals carry no document contents", () => {
  it("names counts and codes only when no table survives", () => {
    const error = noExportableTables(
      [
        { code: "PBI_TABLE_HIDDEN", tables: 13, columns: 0 },
        { code: "PBI_COLUMN_UNREADABLE", tables: 0, columns: 4 },
      ],
      true,
    );
    const text = serialized(error);
    expect(error.code).toBe("PBI_NO_EXPORTABLE_TABLES");
    expect(text).toContain("PBI_TABLE_HIDDEN");
    expect(text).not.toMatch(/[Tt]able\d|Sales|Orders/);
    expect(error.message).toContain("Export hidden tables as well");
    expect(error.message).not.toContain("includeHiddenTables");
    expect(error.cause).toBeUndefined();
  });

  it("says no tables were found for a model with none", () => {
    const error = noExportableTables([], false);
    expect(error.message).toContain("no tables");
    expect(error.message).not.toContain("Export hidden tables");
  });

  it("gives a runtime label and stage, never a URL", () => {
    for (const runtime of ["sqlite", "xpress9"] as const)
      for (const stage of ["load", "compile", "instantiate"] as const) {
        const error = runtimeUnavailable(runtime, stage);
        expect(error.code).toBe("PBI_RUNTIME_UNAVAILABLE");
        // The same detail shape the XPress9 loader emits.
        expect(error.details).toEqual({ runtime, stage });
        expect(serialized(error)).not.toMatch(/https?:|file:|[A-Za-z]:\\/);
        expect(error.cause).toBeUndefined();
      }
  });

  it("leaks no table or column name through a capacity refusal on a real model", async () => {
    const error = await refusal(() =>
      exportPbiTables(container, { decodedBytes: 1000 }),
    );
    expect(error.code).toBe("PBI_EXPORT_LIMIT_EXCEEDED");
    const text = serialized(error);
    for (const name of ["People", "Sales", "DateTableTemplate", "Quantity"])
      expect(text).not.toContain(name);
    expect(error.cause).toBeUndefined();
  }, 60_000);

  it("leaks nothing through a damaged model part", async () => {
    const damaged = container.slice();
    // Corrupt inside the stored DataModel body, past every directory record.
    damaged.fill(0x41, 200, 400);
    const error = await refusal(() => readPbiTables(damaged));
    expect(["PBI_MODEL_UNREADABLE", "PBI_INVALID_CONTAINER"]).toContain(
      error.code,
    );
    expect(serialized(error)).not.toContain(marker);
    expect(error.cause).toBeUndefined();
  }, 60_000);
});

describe("the template refusal", () => {
  const template = path.join(FETCHED, "e-covid.pbit");
  it.skipIf(!existsSync(template))(
    "refuses a file with no model part",
    async () => {
      const error = await refusal(() =>
        readPbiTables(new Uint8Array(readFileSync(template))),
      );
      expect(error.code).toBe("PBI_NO_MODEL");
      expect(error.message).toContain("imported data");
      expect(error.cause).toBeUndefined();
    },
    60_000,
  );
});
