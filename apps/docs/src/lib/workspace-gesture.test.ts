import { describe, expect, it } from "vitest";

import {
  ONE_RECTANGLE_ONLY,
  planFill,
  planPaste,
  type GestureGrid,
  type GesturePlan,
  type GestureRect,
} from "./workspace-gesture";
import { WORKSPACE_MAX_GESTURE_CELLS } from "./workspace-protocol";

// The geometry half of a gesture: which cells a paste or a fill would write,
// what it writes into them, and which gestures are refused whole before
// anything is sent. Nothing here touches a DOM or a database.

const RECORD_ID = "record_id";

/**
 * A grid of `rows` records and the columns named, with the Record ID first as
 * the real one has it. A cell's text is its address, so a test can see exactly
 * which cell a value came from.
 */
function grid(
  rows: number,
  fields: readonly string[] = ["name", "region", "headcount"],
  options: {
    readonly references?: readonly string[];
    readonly refused?: readonly string[];
    readonly text?: (recordId: string, field: string) => string | undefined;
  } = {},
): GestureGrid {
  const references = new Set(options.references ?? []);
  const refused = new Set(options.refused ?? []);
  return {
    columns: [
      { field: RECORD_ID, writable: false, series: false },
      ...fields.map((field) => ({
        field,
        writable: true,
        series: !references.has(field),
      })),
    ],
    recordIds: Array.from(
      { length: rows },
      (_unused, index) => `REC-${String(index + 1).padStart(4, "0")}`,
    ),
    text: (recordId, field) =>
      options.text?.(recordId, field) ?? `${recordId}/${field}`,
    refused: (recordId, field) => refused.has(`${recordId}/${field}`),
  };
}

function rect(
  top: number,
  left: number,
  bottom = top,
  right = left,
): GestureRect {
  return { top, left, bottom, right };
}

/** The writes a plan makes, as "record/column=value", for readable assertions. */
function written(plan: GesturePlan): string[] {
  if (plan.kind !== "writes") {
    throw new Error(`Expected writes, got a refusal: ${plan.reason}`);
  }
  return plan.writes.map(
    (write) => `${write.recordId}/${write.column}=${write.value}`,
  );
}

/** What the gesture covered, which is what the selection follows. */
function covered(plan: GesturePlan): GestureRect | null {
  if (plan.kind !== "writes") {
    throw new Error(`Expected writes, got a refusal: ${plan.reason}`);
  }
  return plan.covered;
}

function refusal(plan: GesturePlan): string {
  if (plan.kind !== "refused") {
    throw new Error("Expected a refusal, got writes");
  }
  return plan.reason;
}

describe("planPaste", () => {
  it("writes a block from the anchor when the selection is one cell", () => {
    const plan = planPaste({
      grid: grid(4),
      ranges: [rect(0, 1)],
      block: [
        ["a", "b"],
        ["c", "d"],
      ],
    });

    expect(written(plan)).toEqual([
      "REC-0001/name=a",
      "REC-0001/region=b",
      "REC-0002/name=c",
      "REC-0002/region=d",
    ]);
  });

  it("fills the whole selection from a single copied value", () => {
    const plan = planPaste({
      grid: grid(3),
      ranges: [rect(0, 1, 2, 2)],
      block: [["x"]],
    });

    expect(written(plan)).toEqual([
      "REC-0001/name=x",
      "REC-0001/region=x",
      "REC-0002/name=x",
      "REC-0002/region=x",
      "REC-0003/name=x",
      "REC-0003/region=x",
    ]);
  });

  it("tiles a block whose size divides the selection", () => {
    const plan = planPaste({
      grid: grid(4),
      ranges: [rect(0, 1, 3, 1)],
      block: [["a"], ["b"]],
    });

    expect(written(plan)).toEqual([
      "REC-0001/name=a",
      "REC-0002/name=b",
      "REC-0003/name=a",
      "REC-0004/name=b",
    ]);
  });

  it("overflows the selection from the anchor when the sizes do not divide", () => {
    const plan = planPaste({
      grid: grid(4),
      ranges: [rect(0, 1, 1, 1)],
      block: [["a"], ["b"], ["c"]],
    });

    expect(written(plan)).toEqual([
      "REC-0001/name=a",
      "REC-0002/name=b",
      "REC-0003/name=c",
    ]);
  });

  it("treats a short row of the block as empty cells", () => {
    const plan = planPaste({
      grid: grid(2),
      ranges: [rect(0, 1)],
      block: [["a", "b"], ["c"]],
    });

    expect(written(plan)).toEqual([
      "REC-0001/name=a",
      "REC-0001/region=b",
      "REC-0002/name=c",
      "REC-0002/region=",
    ]);
  });

  it("refuses a paste that needs rows past the last record", () => {
    const plan = planPaste({
      grid: grid(2),
      ranges: [rect(1, 1)],
      block: [["a"], ["b"], ["c"]],
    });

    expect(refusal(plan)).toContain("2 more rows");
    expect(refusal(plan)).toContain("nothing was pasted");
  });

  it("refuses a paste that needs columns past the last one", () => {
    const plan = planPaste({
      grid: grid(2, ["name"]),
      ranges: [rect(0, 1)],
      block: [["a", "b", "c"]],
    });

    expect(refusal(plan)).toContain("2 more columns");
  });

  it("refuses a paste that would write to the Record ID", () => {
    const plan = planPaste({
      grid: grid(2),
      ranges: [rect(0, 0)],
      block: [["a"]],
    });

    expect(refusal(plan)).toContain("Record ID");
  });

  it("refuses a paste into more than one selected rectangle", () => {
    const plan = planPaste({
      grid: grid(3),
      ranges: [rect(0, 1), rect(2, 2)],
      block: [["a"]],
    });

    expect(refusal(plan)).toBe(ONE_RECTANGLE_ONLY);
  });

  it("refuses a paste with nothing selected", () => {
    const plan = planPaste({ grid: grid(3), ranges: [], block: [["a"]] });

    expect(refusal(plan)).toBe(ONE_RECTANGLE_ONLY);
  });

  it("covers the block it wrote, for the selection to follow", () => {
    const plan = planPaste({
      grid: grid(4),
      ranges: [rect(1, 1)],
      block: [
        ["a", "b"],
        ["c", "d"],
      ],
    });

    expect(covered(plan)).toEqual(rect(1, 1, 2, 2));
  });

  it("writes nothing for an empty clipboard", () => {
    expect(
      written(planPaste({ grid: grid(3), ranges: [rect(0, 1)], block: [] })),
    ).toEqual([]);
  });

  it("refuses a paste bigger than one step applies", () => {
    const columns = Array.from({ length: 10 }, (_unused, index) => `c${index}`);
    const rows = Math.ceil(WORKSPACE_MAX_GESTURE_CELLS / 10) + 1;
    const plan = planPaste({
      grid: grid(rows, columns),
      ranges: [rect(0, 1)],
      block: Array.from({ length: rows }, () => columns.map(() => "x")),
    });

    expect(refusal(plan)).toContain(String(WORKSPACE_MAX_GESTURE_CELLS));
    expect(refusal(plan)).toContain("smaller");
  });

  it("allows a paste of exactly the largest step", () => {
    const columns = Array.from({ length: 10 }, (_unused, index) => `c${index}`);
    const rows = WORKSPACE_MAX_GESTURE_CELLS / 10;
    const plan = planPaste({
      grid: grid(rows, columns),
      ranges: [rect(0, 1)],
      block: Array.from({ length: rows }, () => columns.map(() => "x")),
    });

    expect(written(plan)).toHaveLength(WORKSPACE_MAX_GESTURE_CELLS);
  });
});

describe("planFill", () => {
  const numbers = (recordId: string, field: string): string | undefined =>
    field === "headcount"
      ? String(Number(recordId.slice(-4)) * 10)
      : `${recordId}/${field}`;

  it("extends a series down from the source", () => {
    const plan = planFill({
      grid: grid(5, ["name", "region", "headcount"], { text: numbers }),
      ranges: [rect(0, 3, 1, 3)],
      pointer: { row: 3, column: 3 },
    });

    expect(written(plan)).toEqual([
      "REC-0003/headcount=30",
      "REC-0004/headcount=40",
    ]);
  });

  it("extends backwards when the drag goes up", () => {
    const plan = planFill({
      grid: grid(5, ["name", "region", "headcount"], { text: numbers }),
      ranges: [rect(2, 3, 3, 3)],
      pointer: { row: 0, column: 3 },
    });

    expect(written(plan)).toEqual([
      "REC-0001/headcount=10",
      "REC-0002/headcount=20",
    ]);
  });

  it("fills sideways when the drag moves further across than down", () => {
    const plan = planFill({
      grid: grid(2),
      ranges: [rect(0, 1)],
      pointer: { row: 0, column: 3 },
    });

    expect(written(plan)).toEqual([
      "REC-0001/region=REC-0001/name",
      "REC-0001/headcount=REC-0001/name",
    ]);
  });

  it("prefers the vertical axis when the drag is equally far both ways", () => {
    const plan = planFill({
      grid: grid(3),
      ranges: [rect(0, 1)],
      pointer: { row: 1, column: 2 },
    });

    expect(written(plan)).toEqual(["REC-0002/name=REC-0001/name"]);
  });

  it("writes nothing when the drag stays inside the source", () => {
    expect(
      written(
        planFill({
          grid: grid(4),
          ranges: [rect(0, 1, 2, 2)],
          pointer: { row: 1, column: 2 },
        }),
      ),
    ).toEqual([]);
  });

  it("copies a foreign-key column rather than reading a series from it", () => {
    const plan = planFill({
      grid: grid(4, ["name", "region"], {
        references: ["region"],
        text: (recordId, field) =>
          field === "region" ? `REG-000${recordId.slice(-1)}` : undefined,
      }),
      ranges: [rect(0, 2, 1, 2)],
      pointer: { row: 3, column: 2 },
    });

    expect(written(plan)).toEqual([
      "REC-0003/region=REG-0001",
      "REC-0004/region=REG-0002",
    ]);
  });

  it("refuses a fill whose target reaches the Record ID", () => {
    const plan = planFill({
      grid: grid(3),
      ranges: [rect(0, 1, 2, 1)],
      pointer: { row: 1, column: 0 },
    });

    expect(refusal(plan)).toContain("Record ID");
  });

  it("refuses a fill from a source a cell of which was refused", () => {
    const plan = planFill({
      grid: grid(4, ["name"], { refused: ["REC-0002/name"] }),
      ranges: [rect(0, 1, 1, 1)],
      pointer: { row: 3, column: 1 },
    });

    expect(refusal(plan)).toContain("refused");
  });

  it("refuses a fill into more than one selected rectangle", () => {
    const plan = planFill({
      grid: grid(3),
      ranges: [rect(0, 1), rect(2, 1)],
      pointer: { row: 2, column: 1 },
    });

    expect(refusal(plan)).toBe(ONE_RECTANGLE_ONLY);
  });

  it("clamps a drag past the last row to the last row", () => {
    const plan = planFill({
      grid: grid(3),
      ranges: [rect(0, 1)],
      pointer: { row: 99, column: 1 },
    });

    expect(written(plan)).toEqual([
      "REC-0002/name=REC-0001/name",
      "REC-0003/name=REC-0001/name",
    ]);
  });

  it("refuses a fill bigger than one step applies", () => {
    const rows = WORKSPACE_MAX_GESTURE_CELLS + 2;
    const plan = planFill({
      grid: grid(rows, ["name"]),
      ranges: [rect(0, 1)],
      pointer: { row: rows - 1, column: 1 },
    });

    expect(refusal(plan)).toContain(String(WORKSPACE_MAX_GESTURE_CELLS));
  });

  it("covers the source and what it filled, for the selection to follow", () => {
    const plan = planFill({
      grid: grid(5),
      ranges: [rect(1, 1, 1, 2)],
      pointer: { row: 3, column: 2 },
    });

    expect(covered(plan)).toEqual(rect(1, 1, 3, 2));
  });

  it("covers nothing when the drag stays inside the source", () => {
    expect(
      covered(
        planFill({
          grid: grid(4),
          ranges: [rect(0, 1, 2, 2)],
          pointer: { row: 1, column: 2 },
        }),
      ),
    ).toBeNull();
  });

  it("fills one line per column of a multi-column source", () => {
    const plan = planFill({
      grid: grid(3, ["a", "b"], {
        text: (recordId, field) => `${field}-${recordId.slice(-1)}`,
      }),
      ranges: [rect(0, 1, 1, 2)],
      pointer: { row: 2, column: 2 },
    });

    expect(written(plan)).toEqual(["REC-0003/a=a-3", "REC-0003/b=b-3"]);
  });
});
