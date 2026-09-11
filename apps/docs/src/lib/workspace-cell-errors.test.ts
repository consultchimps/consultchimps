import { describe, expect, it } from "vitest";

import {
  answerFailure,
  answerFailures,
  cellKey,
  dismissFailures,
  failureReport,
  failuresAt,
  NO_FAILURES,
  NOTHING_REFUSED,
  recordFailure,
  recordFailures,
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

describe("explanations belong to one workspace", () => {
  const FIRST = 4;
  const NEXT = 5;

  const refusedInFirst = recordFailure(
    NOTHING_REFUSED,
    FIRST,
    headcount,
    "Not a whole number.",
  );

  it("shows nothing before a workspace is open", () => {
    expect(failureReport(failuresAt(NOTHING_REFUSED, FIRST))).toBeNull();
  });

  it("shows what was refused in the workspace now held", () => {
    expect(failureReport(failuresAt(refusedInFirst, FIRST))).toBe(
      "Not a whole number.",
    );
  });

  it("shows nothing from a workspace that has been replaced", () => {
    // The case this exists for: New or Open replaces the database while the
    // grid stays mounted. A key names a table, a Record ID and a column, and
    // every one of those means something else in the next file, where
    // "CUST-0001" very likely exists as a different record.
    expect(failureReport(failuresAt(refusedInFirst, NEXT))).toBeNull();
  });

  it("starts the next workspace clean rather than merging into it", () => {
    const refusedInNext = recordFailure(
      refusedInFirst,
      NEXT,
      name,
      "A required column was left empty.",
    );

    expect(failureReport(failuresAt(refusedInNext, NEXT))).toBe(
      "A required column was left empty.",
    );
    // Not "1 other edit was refused as well": the first workspace's refusal did
    // not come across.
    expect(failuresAt(refusedInNext, NEXT).size).toBe(1);
  });

  it("does not let a reply from the old workspace answer the new one", () => {
    const refusedInNext = recordFailure(
      NOTHING_REFUSED,
      NEXT,
      headcount,
      "Not a whole number.",
    );

    // A command sent before the replacement can still settle after it. Its
    // answer is about a database nobody holds any more.
    expect(
      failureReport(
        failuresAt(answerFailure(refusedInNext, FIRST, headcount), NEXT),
      ),
    ).toBe("Not a whole number.");
  });

  it("clears what its own workspace answered", () => {
    expect(
      failureReport(
        failuresAt(answerFailure(refusedInFirst, FIRST, headcount), FIRST),
      ),
    ).toBeNull();
  });

  it("puts the whole set away when the visitor asks", () => {
    expect(failureReport(failuresAt(dismissFailures(FIRST), FIRST))).toBeNull();
  });

  it("leaves what it was given alone", () => {
    recordFailure(refusedInFirst, NEXT, name, "Something else.");
    answerFailure(refusedInFirst, FIRST, headcount);

    expect(failureReport(failuresAt(refusedInFirst, FIRST))).toBe(
      "Not a whole number.",
    );
  });
});

describe("a gesture's worth of explanations", () => {
  it("records every refused cell of one gesture in one move", () => {
    const recorded = recordFailures(NOTHING_REFUSED, 4, [
      { key: headcount, message: "not a whole number" },
      { key: otherRecord, message: "not a whole number either" },
    ]);

    expect(failuresAt(recorded, 4).size).toBe(2);
    // The newest is the last one recorded, shown in full above the count.
    expect(failureReport(failuresAt(recorded, 4))).toContain(
      "not a whole number either",
    );
    expect(failureReport(failuresAt(recorded, 4))).toContain("1 other edit");
  });

  it("answers the accepted cells of a gesture and leaves the refused ones", () => {
    const before = recordFailures(NOTHING_REFUSED, 4, [
      { key: headcount, message: "refused here" },
      { key: name, message: "refused there" },
    ]);

    const after = answerFailures(before, 4, [name, otherRecord]);

    expect([...failuresAt(after, 4).keys()]).toEqual([headcount]);
  });

  it("records nothing for a gesture that refused nothing", () => {
    expect(recordFailures(NOTHING_REFUSED, 4, [])).toBe(NOTHING_REFUSED);
    expect(answerFailures(NOTHING_REFUSED, 4, [])).toBe(NOTHING_REFUSED);
  });

  it("ignores a gesture recorded against a workspace that has been replaced", () => {
    const recorded = recordFailures(NOTHING_REFUSED, 4, [
      { key: headcount, message: "refused" },
    ]);

    expect(failuresAt(recorded, 5)).toBe(NO_FAILURES);
    expect(answerFailures(recorded, 5, [headcount])).toBe(recorded);
  });
});
