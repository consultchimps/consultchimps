import { describe, expect, it } from "vitest";

import {
  cellKey,
  failureReport,
  NO_FAILURES,
  tableKey,
  withFailure,
  withoutFailure,
} from "./workspace-cell-errors";

const headcount = cellKey("Customer", "CUST-0001", "headcount");
const name = cellKey("Customer", "CUST-0001", "name");
const otherRecord = cellKey("Customer", "CUST-0002", "headcount");

describe("cell keys", () => {
  it("tells apart every cell an edit can be addressed to", () => {
    const keys = new Set([
      headcount,
      name,
      otherRecord,
      cellKey("Region", "CUST-0001", "headcount"),
      tableKey("Customer"),
      tableKey("Region"),
    ]);
    expect(keys.size).toBe(6);
  });

  it("cannot be confused by a name that looks like a key", () => {
    // Names are encoded rather than joined, so a column called what a joiner
    // would produce still keys to itself alone.
    expect(cellKey("Customer", "CUST-0001", 'a","b')).not.toBe(
      cellKey("Customer", "CUST-0001", "a"),
    );
    expect(cellKey("a", "b", "c")).not.toBe(cellKey("a b", "", "c"));
  });
});

describe("standing explanations", () => {
  it("shows nothing until something is refused", () => {
    expect(failureReport(NO_FAILURES)).toBeNull();
  });

  it("keeps a refusal when a different cell succeeds", () => {
    // The case this exists for: two edits in flight, the first refused and the
    // second accepted, and the reply order erasing the only sentence saying why
    // a cell snapped back.
    const refused = withFailure(NO_FAILURES, headcount, "Not a whole number.");
    const after = withoutFailure(refused, name);

    expect(failureReport(after)).toBe("Not a whole number.");
  });

  it("clears a refusal when that same cell succeeds", () => {
    const refused = withFailure(NO_FAILURES, headcount, "Not a whole number.");

    expect(failureReport(withoutFailure(refused, headcount))).toBeNull();
  });

  it("replaces a refusal when that same cell is refused again", () => {
    const first = withFailure(NO_FAILURES, headcount, "Not a whole number.");
    const second = withFailure(first, headcount, "Outside the range.");

    expect(failureReport(second)).toBe("Outside the range.");
  });

  it("keeps one explanation per cell, and says how many are standing", () => {
    const both = withFailure(
      withFailure(NO_FAILURES, headcount, "Not a whole number."),
      name,
      "A required column was left empty.",
    );

    // The newest in full, because it is the one just earned; the rest counted,
    // because these sentences are long and stacking them buries it.
    expect(failureReport(both)).toBe(
      "A required column was left empty.\n\n1 other edit was refused as well, and that cell still holds the value the workspace has",
    );

    const three = withFailure(both, otherRecord, "Not a whole number.");
    expect(failureReport(three)).toContain(
      "2 other edits were refused as well, and those cells still hold the values the workspace has",
    );
  });

  it("brings the one underneath back as the cells ahead of it are dealt with", () => {
    const both = withFailure(
      withFailure(NO_FAILURES, headcount, "Not a whole number."),
      name,
      "A required column was left empty.",
    );

    expect(failureReport(withoutFailure(both, name))).toBe(
      "Not a whole number.",
    );
  });

  it("moves a cell refused again to the front of the queue", () => {
    const both = withFailure(
      withFailure(NO_FAILURES, headcount, "Not a whole number."),
      name,
      "A required column was left empty.",
    );
    const again = withFailure(both, headcount, "Still not a whole number.");

    expect(failureReport(again)).toContain("Still not a whole number.");
    expect(failureReport(again)).toContain("1 other edit was refused as well");
  });

  it("holds a refused edit and an unreadable table apart", () => {
    const both = withFailure(
      withFailure(
        NO_FAILURES,
        tableKey("Region"),
        "That table could not be read.",
      ),
      headcount,
      "Not a whole number.",
    );

    // Neither answers the other, so neither clears it.
    expect(failureReport(withoutFailure(both, headcount))).toBe(
      "That table could not be read.",
    );
  });

  it("leaves what it was given alone", () => {
    const refused = withFailure(NO_FAILURES, headcount, "Not a whole number.");

    withFailure(refused, name, "A required column was left empty.");
    withoutFailure(refused, headcount);

    expect(failureReport(refused)).toBe("Not a whole number.");
    expect(failureReport(NO_FAILURES)).toBeNull();
  });

  it("does nothing when told to clear what is not standing", () => {
    const refused = withFailure(NO_FAILURES, headcount, "Not a whole number.");

    expect(withoutFailure(refused, name)).toBe(refused);
  });
});
