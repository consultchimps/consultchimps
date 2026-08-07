/**
 * Guardrail 3 of ARCHITECTURE.md: contract completeness.
 *
 * The L4 table in src/contract.ts declares what each operation does to each
 * tracked structure. This test walks the table and reports every structure
 * that has no declared cell. Phase 1 asserts the exact missing set rather than
 * failing on it: the debt is visible and enumerated, and shrinking it requires
 * editing this file on purpose.
 */
import { describe, expect, it } from "vitest";

import {
  CONTRACT,
  OPERATIONS,
  TRACKED_STRUCTURES,
  UNDECIDED_MERGE_STRUCTURES,
  UNDECIDED_SPLIT_STRUCTURES,
  type Operation,
  type Structure,
} from "../src/contract.js";

/** Tracked structures with no declared behavior for `operation`. */
function undeclaredStructures(operation: Operation): readonly Structure[] {
  const declared = CONTRACT[operation];
  return TRACKED_STRUCTURES.filter(
    (structure) => declared[structure] === undefined,
  );
}

/**
 * The split cells still owed a decision. Each one is explained in
 * UNDECIDED_SPLIT_STRUCTURES; removing an entry here means the cell was
 * declared, which is the only way this list is allowed to change.
 */
const EXPECTED_MISSING_FOR_SPLIT: readonly Structure[] = [
  "drawings-charts",
  "defined-names",
  "excel-table-totals-row",
  "external-links",
  "formulas-array",
];

/**
 * The merge cells still owed a decision. Phase 1b declared everything the
 * conformance corpus can exercise; each entry here is explained in
 * UNDECIDED_MERGE_STRUCTURES.
 */
const EXPECTED_MISSING_FOR_MERGE: readonly Structure[] = ["external-links"];

/** Operations with no column yet; each is wholly undeclared. */
const UNDECLARED_OPERATIONS: readonly Operation[] = [
  "consolidate",
  "values",
  "describe",
];

describe("contract: table shape", () => {
  it("declares a column for every operation", () => {
    expect(Object.keys(CONTRACT).sort()).toEqual([...OPERATIONS].sort());
  });

  it("declares no behavior for a structure outside the tracked set", () => {
    const tracked = new Set<string>(TRACKED_STRUCTURES);
    for (const operation of OPERATIONS) {
      for (const structure of Object.keys(CONTRACT[operation])) {
        expect(tracked, `${operation}.${structure} is not tracked`).toContain(
          structure,
        );
      }
    }
  });

  it("tracks each structure exactly once", () => {
    expect(new Set(TRACKED_STRUCTURES).size).toBe(TRACKED_STRUCTURES.length);
  });
});

describe("contract: completeness", () => {
  it("owes split exactly the structures recorded as undecided", () => {
    // Sorted so the assertion reports a set difference, not an ordering diff.
    expect([...undeclaredStructures("split")].sort()).toEqual(
      [...EXPECTED_MISSING_FOR_SPLIT].sort(),
    );
  });

  it("explains every missing split cell", () => {
    expect(Object.keys(UNDECIDED_SPLIT_STRUCTURES).sort()).toEqual(
      [...EXPECTED_MISSING_FOR_SPLIT].sort(),
    );
    for (const reason of Object.values(UNDECIDED_SPLIT_STRUCTURES)) {
      expect(reason.length).toBeGreaterThan(0);
    }
  });

  it.each(UNDECLARED_OPERATIONS)(
    "leaves %s wholly undeclared until its phase lands",
    (operation) => {
      expect(CONTRACT[operation]).toEqual({});
      expect(undeclaredStructures(operation)).toEqual([...TRACKED_STRUCTURES]);
    },
  );

  it("owes merge exactly the structures recorded as undecided", () => {
    expect([...undeclaredStructures("merge")].sort()).toEqual(
      [...EXPECTED_MISSING_FOR_MERGE].sort(),
    );
  });

  it("explains every missing merge cell", () => {
    expect(Object.keys(UNDECIDED_MERGE_STRUCTURES).sort()).toEqual(
      [...EXPECTED_MISSING_FOR_MERGE].sort(),
    );
    for (const reason of Object.values(UNDECIDED_MERGE_STRUCTURES)) {
      expect(reason.length).toBeGreaterThan(0);
    }
  });

  it("reports the declared cells for merge", () => {
    const declared = CONTRACT.merge;
    expect(Object.keys(declared)).toHaveLength(
      TRACKED_STRUCTURES.length - EXPECTED_MISSING_FOR_MERGE.length,
    );
    // The three shapes of merge behavior, one example each: a structure the
    // transplant never touches, one it repairs, and one it refuses to carry.
    expect(declared["merged-cells"]).toBe("preserve");
    expect(declared["excel-tables"]).toBe("fix");
    expect(declared["pivot-tables"]).toBe("strip-warn");
  });

  it("reports the declared cells for split", () => {
    const declared = CONTRACT.split;
    // The count is asserted rather than described so that adding a cell
    // without updating EXPECTED_MISSING_FOR_SPLIT cannot pass silently.
    expect(Object.keys(declared)).toHaveLength(
      TRACKED_STRUCTURES.length - EXPECTED_MISSING_FOR_SPLIT.length,
    );
    expect(declared["merged-cells"]).toBe("fix");
    expect(declared["formulas-structured-ref"]).toBe("preserve");
    expect(declared["vba-project"]).toBe("preserve");
  });
});
