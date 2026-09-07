import { describe, expect, it } from "vitest";

import {
  contrastRatio,
  NEUTRAL_PALETTE,
  parseHexColor,
  ThemeError,
  validateCategorical,
  validatePalette,
  type Palette,
} from "../src/index.js";

describe("contrastRatio", () => {
  it("computes the WCAG contrast of black on white as 21", () => {
    expect(contrastRatio("#000000", "#ffffff")).toBeCloseTo(21, 5);
  });

  it("is symmetric and returns 1 for identical colours", () => {
    expect(contrastRatio("#2a78d6", "#fcfcfb")).toBeCloseTo(
      contrastRatio("#fcfcfb", "#2a78d6"),
      10,
    );
    expect(contrastRatio("#2a78d6", "#2a78d6")).toBeCloseTo(1, 10);
  });

  it("throws a coded error on a malformed hex value", () => {
    try {
      parseHexColor("not-a-colour");
      expect.unreachable("malformed hex should throw");
    } catch (error) {
      expect(error).toBeInstanceOf(ThemeError);
      expect((error as ThemeError).code).toBe("THEME_INVALID_COLOR");
    }
  });
});

describe("validatePalette", () => {
  it("passes the neutral palette in both modes", () => {
    expect(validatePalette(NEUTRAL_PALETTE, "light").valid).toBe(true);
    expect(validatePalette(NEUTRAL_PALETTE, "dark").valid).toBe(true);
  });

  it("reports an incomplete runtime palette instead of throwing", () => {
    const malformed = {
      name: "partial",
      // Missing the dark value a runtime palette might omit.
      surface: { light: "#fcfcfb" },
      ink: NEUTRAL_PALETTE.ink,
      categorical: [{ light: "#2a78d6" }],
      sequential: NEUTRAL_PALETTE.sequential,
      semantic: NEUTRAL_PALETTE.semantic,
    } as unknown as Palette;
    const report = validatePalette(malformed, "dark");
    expect(report.valid).toBe(false);
    expect(report.issues.some((issue) => issue.check === "invalid-color")).toBe(
      true,
    );
  });

  it("reports a palette missing its categorical colours as a shape issue", () => {
    const malformed = {
      name: "no-categorical",
      surface: NEUTRAL_PALETTE.surface,
      ink: NEUTRAL_PALETTE.ink,
      sequential: NEUTRAL_PALETTE.sequential,
      semantic: NEUTRAL_PALETTE.semantic,
    } as unknown as Palette;
    const report = validatePalette(malformed, "light");
    expect(report.valid).toBe(false);
    expect(report.issues.some((issue) => issue.check === "invalid-shape")).toBe(
      true,
    );
  });

  it("reports light-mode contrast shortfalls as warnings, not errors", () => {
    const report = validatePalette(NEUTRAL_PALETTE, "light");
    const contrastIssues = report.issues.filter(
      (issue) => issue.check === "surface-contrast",
    );
    expect(contrastIssues.length).toBeGreaterThan(0);
    expect(contrastIssues.every((issue) => issue.severity === "warning")).toBe(
      true,
    );
  });
});

describe("validateCategorical", () => {
  it("flags a near-duplicate adjacent pair as an error", () => {
    const report = validateCategorical(["#2a78d6", "#2b79d7"], "light", {
      surface: "#fcfcfb",
    });
    expect(report.valid).toBe(false);
    expect(
      report.issues.some(
        (issue) =>
          issue.check === "normal-vision-distinctness" &&
          issue.severity === "error",
      ),
    ).toBe(true);
  });

  it("reports a malformed colour as a structured issue instead of throwing", () => {
    const report = validateCategorical(
      ["#2a78d6", "#fff", "not-a-colour"],
      "light",
      {
        surface: "#fcfcfb",
      },
    );
    expect(report.valid).toBe(false);
    const invalid = report.issues.filter(
      (issue) => issue.check === "invalid-color",
    );
    expect(invalid).toHaveLength(2);
    expect(invalid.every((issue) => issue.severity === "error")).toBe(true);
  });

  it("reports a malformed surface without throwing", () => {
    const report = validateCategorical(["#2a78d6"], "light", {
      surface: "#zzz",
    });
    expect(report.valid).toBe(false);
    expect(
      report.issues.some(
        (issue) =>
          issue.check === "invalid-color" && "surface" in issue.details,
      ),
    ).toBe(true);
  });

  it("flags a washed-out hue as below the chroma floor", () => {
    const report = validateCategorical(["#8a8a88"], "light", {
      surface: "#fcfcfb",
    });
    expect(report.valid).toBe(false);
    expect(report.issues.some((issue) => issue.check === "chroma-floor")).toBe(
      true,
    );
  });
});
