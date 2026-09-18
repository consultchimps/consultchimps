import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it, vi } from "vitest";

import type { PbiManifest } from "../src/manifest.js";
import type * as Vertipaq from "../src/vertipaq.js";
import { FIXTURES } from "./oracle.js";

/**
 * Surface contract decision 3: an unexpected engine error inside a column
 * decoder excludes that one column with its own reason code, so the manifest
 * never reports an internal fault as damaged source data, and the export still
 * succeeds on every other column.
 *
 * The fault is injected where a genuine one would happen, inside the segment
 * reader, and only on the first column it is asked for, so the run proves both
 * halves at once: the affected column carries the new code, and the rest of the
 * model is exported beside it.
 */
let faults = 0;
vi.mock("../src/vertipaq.js", async (importOriginal) => {
  const actual = await importOriginal<typeof Vertipaq>();
  return {
    ...actual,
    decodeSegmentIds: (
      ...parameters: Parameters<typeof actual.decodeSegmentIds>
    ) => {
      if (faults === 0) {
        faults++;
        // The shape an engine fault takes: not one of this reader's controlled
        // error types, and carrying text that must never reach the user.
        throw new TypeError("segment.records is not iterable");
      }
      return actual.decodeSegmentIds(...parameters);
    },
  };
});

const { exportPbiTables } = await import("../src/pipeline.js");

const fixture = new Uint8Array(
  readFileSync(path.join(FIXTURES, "a-2018-fuzzy.pbix")),
);

describe("an unexpected fault inside a column decoder", () => {
  it("excludes the column as a decoder error and exports the rest", async () => {
    const outcome = await exportPbiTables(fixture);
    expect(faults).toBe(1);
    expect(outcome.result.metrics.exportedTables).toBeGreaterThanOrEqual(1);

    const manifest = JSON.parse(
      new TextDecoder().decode(outcome.outputs[1]!.bytes),
    ) as PbiManifest;
    const affected = [
      ...manifest.tables.flatMap((table) => table.columns),
      ...manifest.excludedTables.flatMap((table) => table.columns),
    ].filter((column) =>
      column.reasons.some(
        (reason) => reason.code === "PBI_COLUMN_DECODER_ERROR",
      ),
    );
    expect(affected).toHaveLength(1);
    expect(affected[0]!.reasons).toEqual([
      { code: "PBI_COLUMN_DECODER_ERROR", count: 1 },
    ]);

    const warning = outcome.result.warnings.find((text) =>
      text.includes("failed while decoding"),
    );
    expect(warning).toBeDefined();
    // The injected message is the engine's, so none of it may survive into
    // anything the user reads.
    for (const text of outcome.result.warnings)
      expect(text).not.toContain("not iterable");
  });
});

describe("the classification of a column failure", () => {
  it("treats the parser's own structural errors as the file's fault, not the reader's", async () => {
    const { columnFailureReason } = await import("../src/model.js");
    const { VertipaqError } = await import("../src/vertipaq.js");
    const { HuffmanError } = await import("../src/huffman.js");
    expect(
      columnFailureReason(new VertipaqError("member contradicts itself")),
    ).toBe("PBI_COLUMN_UNREADABLE");
    expect(
      columnFailureReason(new HuffmanError("page contradicts itself")),
    ).toBe("PBI_COLUMN_UNREADABLE");
    expect(columnFailureReason(new TypeError("reader bug"))).toBe(
      "PBI_COLUMN_DECODER_ERROR",
    );
  });
});
