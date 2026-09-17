import { describe, expect, it } from "vitest";
import { ConsultChimpsError } from "@consultchimps/core";
import { exportPbiTables, readPbiTables } from "../src/pipeline.js";
import { outputNames } from "../src/pipeline.js";

/**
 * Section C, "Options and container": every invalid option aggregates into one
 * PBI_INVALID_OPTIONS with at most one detail per path, in the fixed nine-path
 * order, validated before any input read or locator call.
 */

const marker = "CONFIDENTIAL_SYNTHETIC_MARKER";

/** A view whose reads would throw, proving nothing touched the input. */
function unreadableInput(): Uint8Array {
  return new Proxy(new Uint8Array(0), {
    get(target, property) {
      if (property === "byteLength" || property === "length")
        throw new Error(`input was read: ${marker}`);
      return Reflect.get(target, property) as unknown;
    },
  }) as Uint8Array;
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

const paths = (error: ConsultChimpsError): string[] =>
  (error.details as { invalidOptions: { path: string }[] }).invalidOptions.map(
    (entry) => entry.path,
  );

describe("option validation", () => {
  it("aggregates every invalid path in the fixed order, before reading input", async () => {
    const error = await refusal(() =>
      exportPbiTables(unreadableInput(), {
        inputBytes: 0,
        decodedBytes: -1,
        outputBytes: 1.5,
        peakBytes: Number.NaN,
        includeHiddenTables: "yes" as unknown as boolean,
        outputName: 7 as unknown as string,
        runtime: {
          sql: { locateFile: () => "x", wasmBinary: new Uint8Array(1) },
          xpress9: {},
        },
      }),
    );
    expect(error.code).toBe("PBI_INVALID_OPTIONS");
    expect(paths(error)).toEqual([
      "inputBytes",
      "decodedBytes",
      "outputBytes",
      "peakBytes",
      "includeHiddenTables",
      "outputName",
      "runtime.sql",
      "runtime.xpress9",
    ]);
  });

  it("reports only the parent when runtime itself is malformed", async () => {
    const error = await refusal(() =>
      exportPbiTables(unreadableInput(), {
        runtime: [] as unknown as Record<string, never>,
      }),
    );
    expect(paths(error)).toEqual(["runtime"]);
  });

  it("never echoes the supplied value", async () => {
    const error = await refusal(() =>
      exportPbiTables(unreadableInput(), {
        outputName: marker as unknown as string,
        includeHiddenTables: marker as unknown as boolean,
      }),
    );
    expect(JSON.stringify(error.details)).not.toContain(marker);
    expect(error.message).not.toContain(marker);
    expect(error.cause).toBeUndefined();
  });

  it("omits outputBytes and outputName from the decode-only entry point", async () => {
    const error = await refusal(() =>
      readPbiTables(unreadableInput(), {
        peakBytes: 0,
        includeHiddenTables: 1 as unknown as boolean,
      }),
    );
    expect(paths(error)).toEqual(["peakBytes", "includeHiddenTables"]);
  });

  it("accepts a runtime child with exactly one source", async () => {
    // Reaching the input read means validation passed; the empty view then
    // fails as an unreadable container, not as an invalid option.
    const error = await refusal(() =>
      readPbiTables(new Uint8Array(0), {
        runtime: { xpress9: { wasmBinary: new Uint8Array([1]) } },
      }),
    );
    expect(error.code).toBe("PBI_INVALID_CONTAINER");
  });
});

describe("output names", () => {
  it("uses the fixed default", () => {
    expect(outputNames("power-bi-tables.xlsx")).toEqual({
      workbook: "power-bi-tables.xlsx",
      manifest: "power-bi-tables.manifest.json",
    });
  });

  it("removes one trailing .xlsx case-insensitively and keeps other extensions", () => {
    expect(outputNames("Quarter Report.XLSX").workbook).toBe(
      "Quarter Report.xlsx",
    );
    expect(outputNames("report.csv").workbook).toBe("report.csv.xlsx");
    expect(outputNames("a.xlsx.xlsx").workbook).toBe("a.xlsx.xlsx");
  });

  it("falls back for an empty or fully sanitized stem", () => {
    expect(outputNames("   ").workbook).toBe("power-bi-tables.xlsx");
    expect(outputNames(".xlsx").workbook).toBe("power-bi-tables.xlsx");
  });

  it("strips path separators and handles a reserved filename", () => {
    expect(outputNames("../../etc/passwd").workbook).not.toContain("/");
    expect(outputNames("CON").workbook).toBe("_CON.xlsx");
  });
});
