import { describe, expect, it } from "vitest";

import { fillLine } from "./workspace-series";

// The series table from the contract, one case per rule, in both directions.
// Everything here is text: a fill sends the same text the clipboard carries, so
// the library's one conversion point decides what a column will hold and this
// module never has to know a column's type.

/** The three cells after the source, which is the usual drag. */
function down(source: readonly string[], count = 3): string[] {
  const indices = Array.from(
    { length: count },
    (_unused, offset) => source.length + offset,
  );
  return fillLine(source, indices);
}

/** The cells before the source, nearest first. */
function up(source: readonly string[], count = 3): string[] {
  const indices = Array.from(
    { length: count },
    (_unused, offset) => -1 - offset,
  );
  return fillLine(source, indices);
}

describe("fillLine: numbers", () => {
  it("copies a single number", () => {
    expect(down(["5"])).toEqual(["5", "5", "5"]);
  });

  it("extends two numbers with a constant difference", () => {
    expect(down(["1", "2"])).toEqual(["3", "4", "5"]);
    expect(down(["10", "20", "30"])).toEqual(["40", "50", "60"]);
  });

  it("extends a decimal step without floating point noise", () => {
    expect(down(["0.1", "0.2"])).toEqual(["0.3", "0.4", "0.5"]);
  });

  it("keeps the decimal shape the source was written in", () => {
    expect(down(["1.50", "1.60"])).toEqual(["1.70", "1.80", "1.90"]);
  });

  it("extends a descending series past zero", () => {
    expect(down(["5", "3"])).toEqual(["1", "-1", "-3"]);
  });

  it("extends backwards with the same step", () => {
    expect(up(["10", "20"])).toEqual(["0", "-10", "-20"]);
  });

  it("copies a source whose differences are not constant", () => {
    expect(down(["1", "2", "4"])).toEqual(["1", "2", "4"]);
  });

  it("extends numbers too long for exact floating point", () => {
    expect(down(["9007199254740993", "9007199254740995"], 2)).toEqual([
      "9007199254740997",
      "9007199254740999",
    ]);
  });

  it("extends a source with more values than an argument list takes", () => {
    // A fill's source is as tall as the selection and nothing caps what a
    // gesture reads, so measuring the decimals by spreading the line into
    // Math.max threw a raw RangeError somewhere above 125,000 values and the
    // drag died instead of filling its one cell.
    const source = Array.from({ length: 130_000 }, (_unused, index) =>
      String(index),
    );

    expect(fillLine(source, [source.length])).toEqual(["130000"]);
  });

  it("copies a number written in exponent notation", () => {
    expect(down(["1e+21", "2e+21"], 2)).toEqual(["1e+21", "2e+21"]);
  });

  it("never steps the digits of an exponent, which are its magnitude", () => {
    // The trailing-integer rule would read "1e+21" as a counter and answer
    // "1e+22", which is ten times the number the cell holds.
    expect(down(["1e+21"], 2)).toEqual(["1e+21", "1e+21"]);
    expect(down(["1e-7"], 2)).toEqual(["1e-7", "1e-7"]);
    expect(down(["1e+21", "1e+22"], 1)).toEqual(["1e+21"]);
    expect(down(["2.5e+10"], 1)).toEqual(["2.5e+10"]);
  });
});

describe("fillLine: dates", () => {
  it("fills a one-day series from a single date", () => {
    expect(down(["2026-01-31"])).toEqual([
      "2026-02-01",
      "2026-02-02",
      "2026-02-03",
    ]);
  });

  it("extends a constant day step", () => {
    expect(down(["2026-01-01", "2026-01-08"])).toEqual([
      "2026-01-15",
      "2026-01-22",
      "2026-01-29",
    ]);
  });

  it("extends a constant month step when every date shares its day", () => {
    expect(down(["2026-01-15", "2026-02-15"])).toEqual([
      "2026-03-15",
      "2026-04-15",
      "2026-05-15",
    ]);
  });

  it("steps by days when the day of the month is above 28", () => {
    // The 31st has no counterpart in February, so a month step would have to
    // invent a rule for where it lands. The day step is what the dates state.
    expect(down(["2026-01-31", "2026-03-31"], 1)).toEqual(["2026-05-29"]);
  });

  it("crosses a leap day", () => {
    expect(down(["2024-02-28"], 2)).toEqual(["2024-02-29", "2024-03-01"]);
  });

  it("extends a timestamp source, keeping its time part exactly", () => {
    expect(
      down(["2024-01-01T00:00:00.000Z", "2024-01-02T00:00:00.000Z"]),
    ).toEqual([
      "2024-01-03T00:00:00.000Z",
      "2024-01-04T00:00:00.000Z",
      "2024-01-05T00:00:00.000Z",
    ]);
  });

  it("fills a one-day series from a single timestamp", () => {
    expect(down(["2024-01-01T00:00:00.000Z"], 2)).toEqual([
      "2024-01-02T00:00:00.000Z",
      "2024-01-03T00:00:00.000Z",
    ]);
  });

  it("steps a timestamp by months under the same day rule", () => {
    expect(
      down(["2024-01-15T08:00:00+02:00", "2024-02-15T08:00:00+02:00"], 2),
    ).toEqual(["2024-03-15T08:00:00+02:00", "2024-04-15T08:00:00+02:00"]);
  });

  it("copies a source whose values carry different time parts", () => {
    const source = ["2024-01-01T00:00:00.000Z", "2024-01-02T09:30:00.000Z"];
    expect(down(source)).toEqual([source[0], source[1], source[0]]);
  });

  it("copies rather than producing a value the column would refuse", () => {
    // Past the four digit year the grammar the date column keeps has nothing to
    // spell, so the line copies rather than writing values that would be
    // refused one by one.
    expect(down(["9999-12-30", "9999-12-31"], 2)).toEqual([
      "9999-12-30",
      "9999-12-31",
    ]);
  });

  it("extends backwards by the same step", () => {
    expect(up(["2026-03-01", "2026-03-08"], 2)).toEqual([
      "2026-02-22",
      "2026-02-15",
    ]);
  });

  // The date column's grammar spells a year 0000 to 9999, and a date
  // constructor remaps a year from 0 to 99 into the twentieth century, so these
  // are the years where a fill built on one would answer a plausible wrong
  // value: 0099-12-31 plus a day would read as 2000-01-01.
  it("crosses the year 0100 without leaving the century it was in", () => {
    expect(down(["0099-12-31"], 3)).toEqual([
      "0100-01-01",
      "0100-01-02",
      "0100-01-03",
    ]);
  });

  it("steps back into the years below 100", () => {
    expect(up(["0100-01-02", "0100-01-03"], 4)).toEqual([
      "0100-01-01",
      "0099-12-31",
      "0099-12-30",
      "0099-12-29",
    ]);
  });

  it("carries a month series across the same boundary", () => {
    expect(down(["0099-11-15", "0099-12-15"], 3)).toEqual([
      "0100-01-15",
      "0100-02-15",
      "0100-03-15",
    ]);
  });

  it("steps a month series back below the year 100", () => {
    expect(up(["0100-02-10", "0100-03-10"], 3)).toEqual([
      "0100-01-10",
      "0099-12-10",
      "0099-11-10",
    ]);
  });

  it("keeps the first years of the calendar rather than shifting them", () => {
    expect(down(["0001-02-28"], 2)).toEqual(["0001-03-01", "0001-03-02"]);
    expect(down(["0004-02-28"], 1)).toEqual(["0004-02-29"]);
  });

  it("copies at the far end, where the year has no spelling", () => {
    // 9999-12-31 plus a day is the year 10000, which the grammar cannot write,
    // so the line copies rather than filling values the column would refuse.
    expect(down(["9999-12-31"], 2)).toEqual(["9999-12-31", "9999-12-31"]);
    expect(down(["9999-11-30", "9999-12-30"], 1)).toEqual(["9999-11-30"]);
  });

  it("copies at the near end, where the year would go below zero", () => {
    expect(up(["0000-01-01"], 2)).toEqual(["0000-01-01", "0000-01-01"]);
  });
});

describe("fillLine: text", () => {
  it("increments a trailing integer from a single value", () => {
    expect(down(["Item 1"])).toEqual(["Item 2", "Item 3", "Item 4"]);
  });

  it("keeps the zero padding it was given", () => {
    expect(down(["A-007"])).toEqual(["A-008", "A-009", "A-010"]);
  });

  it("lets the digits grow past the padding", () => {
    expect(down(["A-099"])).toEqual(["A-100", "A-101", "A-102"]);
  });

  it("takes the step from two values sharing a prefix", () => {
    expect(down(["Row 1", "Row 3"])).toEqual(["Row 5", "Row 7", "Row 9"]);
  });

  it("copies plain text with no trailing integer", () => {
    expect(down(["North", "South"])).toEqual(["North", "South", "North"]);
  });

  it("copies when the prefixes differ", () => {
    expect(down(["North 1", "South 2"])).toEqual([
      "North 1",
      "South 2",
      "North 1",
    ]);
  });

  it("pads to the width of the value the fill continues from", () => {
    expect(up(["A-08", "A-09"], 2)).toEqual(["A-07", "A-06"]);
  });

  it("extends backwards to zero", () => {
    expect(up(["Row 1", "Row 2"], 1)).toEqual(["Row 0"]);
  });

  it("copies a line whose counter would go below zero", () => {
    // The counter is the unsigned digits on the end, so there is no spelling
    // for -1 here: "Row -1" would read back as the counter 1 under the prefix
    // "Row -". One unspellable value hands the whole line to the copy rule.
    expect(up(["Row 1", "Row 2"], 3)).toEqual(["Row 2", "Row 1", "Row 2"]);
  });

  it("does not double a sign the prefix already carries", () => {
    // "Row -2", "Row -1" is the counter 2 then 1 under the prefix "Row -", a
    // series stepping down by one. It continues to "Row -0" and then has
    // nowhere to go, so the line copies rather than writing "Row --1".
    expect(down(["Row -3", "Row -2"], 2)).toEqual(["Row -1", "Row -0"]);
    expect(down(["Row -2", "Row -1"])).toEqual(["Row -2", "Row -1", "Row -2"]);
  });
});

describe("fillLine: everything else copies", () => {
  it("copies booleans rather than alternating them", () => {
    expect(down(["true"])).toEqual(["true", "true", "true"]);
    expect(down(["true", "true", "false"])).toEqual(["true", "true", "false"]);
  });

  it("copies a line holding an empty cell", () => {
    expect(down([""])).toEqual(["", "", ""]);
    expect(down(["1", ""])).toEqual(["1", "", "1"]);
  });

  it("copies a line of mixed kinds", () => {
    expect(down(["1", "North"])).toEqual(["1", "North", "1"]);
  });

  it("copies cyclically in both directions", () => {
    expect(up(["a", "b", "c"], 4)).toEqual(["c", "b", "a", "c"]);
  });

  it("copies when the caller says the column takes no series", () => {
    // A foreign key stores Record IDs. Reading the trailing integer as a series
    // would point rows at records nobody chose, and they may well exist.
    expect(
      fillLine(["REG-0001", "REG-0002"], [2, 3, 4], { copyOnly: true }),
    ).toEqual(["REG-0001", "REG-0002", "REG-0001"]);
  });

  it("answers an empty source with empty cells", () => {
    expect(fillLine([], [0, 1])).toEqual(["", ""]);
  });
});
