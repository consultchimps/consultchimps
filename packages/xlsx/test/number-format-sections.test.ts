import { describe, expect, it } from "vitest";

import {
  activeNumberFormatSection,
  parseNumberFormatSections,
} from "../src/stream/number-format-sections.js";

describe("number-format sections", () => {
  it("keeps escaped, quoted, spacing, fill, and bracket semicolons in one section", () => {
    expect(
      parseNumberFormatSections('0\\;0;"negative;label"yyyy-mm-dd;0').map(
        (section) => section.code,
      ),
    ).toEqual(["0\\;0", '"negative;label"yyyy-mm-dd', "0"]);
    expect(
      parseNumberFormatSections("0_;0;yyyy-mm-dd").map(
        (section) => section.code,
      ),
    ).toEqual(["0_;0", "yyyy-mm-dd"]);
    expect(
      parseNumberFormatSections("0*;0;yyyy-mm-dd").map(
        (section) => section.code,
      ),
    ).toEqual(["0*;0", "yyyy-mm-dd"]);
    expect(
      parseNumberFormatSections("[Blue;Green]0;yyyy-mm-dd").map(
        (section) => section.code,
      ),
    ).toEqual(["[Blue;Green]0", "yyyy-mm-dd"]);
  });

  it("uses sign sections and supported first-section conditions", () => {
    const signed = parseNumberFormatSections("positive;negative;zero");
    expect(activeNumberFormatSection(signed, 1)?.code).toBe("positive");
    expect(activeNumberFormatSection(signed, -1)?.code).toBe("negative");
    expect(activeNumberFormatSection(signed, 0)?.code).toBe("zero");

    const oneCondition = parseNumberFormatSections("[>=1]date;number");
    expect(activeNumberFormatSection(oneCondition, 1)?.code).toBe("[>=1]date");
    expect(activeNumberFormatSection(oneCondition, 0)?.code).toBe("number");

    const twoConditions = parseNumberFormatSections(
      "[>=1]positive;[<0]negative;zero",
    );
    expect(activeNumberFormatSection(twoConditions, 1)?.code).toBe(
      "[>=1]positive",
    );
    expect(activeNumberFormatSection(twoConditions, -1)?.code).toBe(
      "[<0]negative",
    );
    expect(activeNumberFormatSection(twoConditions, 0)?.code).toBe("zero");

    const exhaustiveConditions = parseNumberFormatSections(
      "[>=0]nonnegative;[<0]negative",
    );
    expect(activeNumberFormatSection(exhaustiveConditions, 0)?.code).toBe(
      "[>=0]nonnegative",
    );
    expect(activeNumberFormatSection(exhaustiveConditions, -1)?.code).toBe(
      "[<0]negative",
    );
  });

  it.each(["number;[<0]date", "number;number;[=0]date"])(
    "rejects unsupported conditional layout %j",
    (format) => {
      expect(() => parseNumberFormatSections(format)).toThrow(/condition/iu);
    },
  );
});
