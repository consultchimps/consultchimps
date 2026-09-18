import { describe, expect, it } from "vitest";

import type { PbiColumnType, PbiValue } from "../src/model.js";
import { workbookExtent, workbookReservationBytes } from "../src/pipeline.js";
import { MAX_WORKSHEET_COLUMNS } from "../src/values.js";
import { writeWorkbook } from "../src/workbook.js";
import type { WorksheetPlan } from "../src/workbook.js";

/**
 * The workbook reservation is a bound, not an estimate.
 *
 * The pipeline reserves this term before the writer allocates anything, so the
 * assertion has to be about the heap the writer actually reaches, not about a
 * proxy for it. Each shape below runs the real writer while the heap is
 * sampled, exactly the way `scripts/measure-cell-cost.ts --workbook` measures a
 * real model, and asserts that the reservation covers the peak that run
 * reached.
 *
 * The shapes are chosen for the terms they attack: a header row wide enough
 * that headers dominate the cells, a table split into parts so the header is
 * emitted again per part, many small sheets against the single fixed term, and
 * the hostile cell shapes where one code unit becomes seven.
 *
 * A sampled peak can only be under-measured, never over-measured: the sampler
 * misses allocations that live and die between two reads, and a collection
 * during the run lowers the reading. So this test can pass a formula that is
 * too small only by luck, and fails one that is too small by structure, which
 * is what the invariant needs.
 */

function columnPlan(
  sheetName: string,
  type: PbiColumnType,
  values: readonly PbiValue[],
  header = "value",
): WorksheetPlan {
  return {
    sheetName,
    start: 0,
    end: values.length,
    columns: [{ key: "1:1", header, type, values }],
  };
}

/** Heap and the buffers hanging off it, which is where the package bytes live. */
function memoryBytes(): number {
  const usage = process.memoryUsage();
  return usage.heapUsed + usage.external;
}

/**
 * The peak this run reaches, measured against a baseline taken immediately
 * before it. The writer's compression is asynchronous, so the sampler observes
 * the stage while the worksheet XML of every part is still live.
 */
async function peakDuring(run: () => Promise<unknown>): Promise<number> {
  const before = memoryBytes();
  let peak = before;
  const sampler = setInterval(() => {
    peak = Math.max(peak, memoryBytes());
  }, 1);
  try {
    await run();
    peak = Math.max(peak, memoryBytes());
  } finally {
    clearInterval(sampler);
  }
  return peak - before;
}

/** A header that escapes to seven code units per character. */
const escapingHeader = "\u001b".repeat(255);

function wideHeaderRow(): WorksheetPlan {
  // Every column a worksheet can hold, each with a header that escapes
  // sevenfold, above a single data row: the case where the header row is the
  // whole workbook and the cell term alone cannot cover it.
  return {
    sheetName: "Wide",
    start: 0,
    end: 1,
    columns: Array.from({ length: MAX_WORKSHEET_COLUMNS }, (_, index) => ({
      key: `1:${index}`,
      header: escapingHeader,
      type: "string" as PbiColumnType,
      values: ["x"],
    })),
  };
}

function splitTable(parts: number, rowsPerPart: number): WorksheetPlan[] {
  // One table's rows shared across numbered parts, each repeating the header
  // row, which is what PBI_TABLE_SPLIT produces.
  const values = Array.from(
    { length: parts * rowsPerPart },
    (_, index) => `row ${index}`,
  );
  return Array.from({ length: parts }, (_, part) => ({
    sheetName: `Sales_${part + 1}`,
    start: part * rowsPerPart,
    end: (part + 1) * rowsPerPart,
    columns: Array.from({ length: 40 }, (_, index) => ({
      key: `1:${index}`,
      header: `${escapingHeader}${index}`,
      type: "string" as PbiColumnType,
      values,
    })),
  }));
}

function manySmallSheets(count: number): WorksheetPlan[] {
  // Enough sheets that the writer's per-part structures, not the cells, are the
  // cost: the single fixed term has to cover all of them together.
  return Array.from({ length: count }, (_, index) =>
    columnPlan(
      `Sheet${index + 1}`,
      "string",
      ["a", "b", "c"],
      `header ${index}`,
    ),
  );
}

const shapes: Array<[string, WorksheetPlan[]]> = [
  [
    "numbers",
    [
      columnPlan(
        "Numbers",
        "double",
        Array.from({ length: 50_000 }, (_, index) => index + 0.5),
      ),
    ],
  ],
  [
    "integers",
    [
      columnPlan(
        "Integers",
        "int64",
        Array.from({ length: 50_000 }, (_, index) => BigInt(index)),
      ),
    ],
  ],
  [
    "nulls",
    [
      columnPlan(
        "Nulls",
        "string",
        Array.from({ length: 50_000 }, () => null),
      ),
    ],
  ],
  [
    "short text",
    [
      columnPlan(
        "Text",
        "string",
        Array.from({ length: 50_000 }, (_, index) => `row ${index}`),
      ),
    ],
  ],
  [
    "long text",
    [
      columnPlan(
        "LongText",
        "string",
        Array.from({ length: 2_000 }, () => "x".repeat(32_000)),
      ),
    ],
  ],
  [
    "text that escapes to seven units per character",
    [
      columnPlan(
        "Escaped",
        "string",
        Array.from({ length: 2_000 }, () => "\u001b".repeat(4_000)),
      ),
    ],
  ],
  [
    "binary as base64",
    [
      columnPlan(
        "Binary",
        "binary",
        Array.from({ length: 5_000 }, () => new Uint8Array(3_000)),
      ),
    ],
  ],
  ["every column a worksheet holds, with escaping headers", [wideHeaderRow()]],
  ["one table split across eight parts", splitTable(8, 5_000)],
  ["two hundred small sheets", manySmallSheets(200)],
];

describe("the workbook reservation", () => {
  for (const [label, plans] of shapes)
    it(`covers the heap the writer reaches for ${label}`, async () => {
      const reserved = workbookReservationBytes(plans);
      const measured = await peakDuring(() => writeWorkbook(plans));
      expect(
        reserved,
        `${label}: reserved ${reserved} bytes, measured peak ${measured} bytes`,
      ).toBeGreaterThanOrEqual(measured);
    }, 120_000);

  it("counts data cells, header cells, rows, parts, header text and sheet names", () => {
    // One column, two rows: two data cells and one header cell, three rows with
    // the header row, one part, and the text is the two values, the header and
    // the sheet name.
    expect(
      workbookExtent([columnPlan("Sheet1", "string", ["abc", "de"], "head")]),
    ).toEqual({ cells: 3, rows: 3, parts: 1, textUnits: 3 + 2 + 4 + 6 });
    // Base64 is four characters per three bytes, rounded up, and a binary
    // column's header and sheet name still count.
    expect(
      workbookExtent([
        columnPlan(
          "S",
          "binary",
          [new Uint8Array(3), new Uint8Array(4)],
          "bytes",
        ),
      ]),
    ).toEqual({ cells: 3, rows: 3, parts: 1, textUnits: 4 + 8 + 5 + 1 });
    // A number column carries no cell text, only its header and sheet name.
    expect(
      workbookExtent([columnPlan("Numbers", "double", [1.5, 2.5], "n")]),
    ).toEqual({ cells: 3, rows: 3, parts: 1, textUnits: 1 + 7 });
  });

  it("charges a code unit the writer escapes at its escaped width", () => {
    const plain = workbookExtent([columnPlan("S", "string", ["aaaa"], "h")]);
    const escaped = workbookExtent([
      // Four control characters, each of which the writer writes as _xHHHH_.
      columnPlan("S", "string", [""], "h"),
    ]);
    expect(escaped.textUnits - plain.textUnits).toBe(4 * 6);
  });

  it("charges a split table's header row once per part", () => {
    const [first] = splitTable(1, 10);
    const parts = splitTable(4, 10);
    const single = workbookExtent([first!]);
    const split = workbookExtent(parts);
    // Four parts of ten rows each carry four header rows, not one.
    expect(split.cells).toBe(single.cells + 3 * 40 + 3 * 10 * 40);
    expect(split.textUnits).toBeGreaterThan(single.textUnits * 3);
  });

  it("grows with the text a model holds, not with its cell count alone", () => {
    expect(
      workbookReservationBytes([
        columnPlan("S", "string", ["a".repeat(10_000), "b"]),
      ]),
    ).toBeGreaterThan(
      workbookReservationBytes([columnPlan("S", "string", ["a", "b"])]),
    );
  });
});
