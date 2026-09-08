import { describe, expect, it } from "vitest";

import {
  categoricalCount,
  isThemeError,
  NEUTRAL_PALETTE,
  NEUTRAL_PALETTES,
  resolveCategorical,
  resolveInk,
  resolveSemantic,
  resolveSequential,
  resolveSurface,
  ThemeError,
} from "../src/index.js";

describe("resolveSurface and resolveInk", () => {
  it("returns the value for the requested mode", () => {
    expect(resolveSurface(NEUTRAL_PALETTE, "light")).toBe("#fcfcfb");
    expect(resolveSurface(NEUTRAL_PALETTE, "dark")).toBe("#1a1a19");
    expect(resolveInk(NEUTRAL_PALETTE, "primary", "light")).toBe("#0b0b0b");
    expect(resolveInk(NEUTRAL_PALETTE, "primary", "dark")).toBe("#ffffff");
  });
});

describe("resolveCategorical", () => {
  it("resolves slots in order for both modes", () => {
    expect(resolveCategorical(NEUTRAL_PALETTE, 0, "light")).toBe("#2a78d6");
    expect(resolveCategorical(NEUTRAL_PALETTE, 0, "dark")).toBe("#3987e5");
    expect(
      resolveCategorical(
        NEUTRAL_PALETTE,
        categoricalCount(NEUTRAL_PALETTE) - 1,
        "light",
      ),
    ).toBe("#e34948");
  });

  it("throws a coded error rather than cycling past the last slot", () => {
    const count = categoricalCount(NEUTRAL_PALETTE);
    try {
      resolveCategorical(NEUTRAL_PALETTE, count, "light");
      expect.unreachable("out-of-range slot should throw");
    } catch (error) {
      expect(isThemeError(error)).toBe(true);
      expect((error as ThemeError).code).toBe(
        "THEME_CATEGORICAL_INDEX_OUT_OF_RANGE",
      );
    }
  });

  it("rejects a fractional or negative slot", () => {
    expect(() => resolveCategorical(NEUTRAL_PALETTE, -1, "light")).toThrow(
      ThemeError,
    );
    expect(() => resolveCategorical(NEUTRAL_PALETTE, 1.5, "light")).toThrow(
      ThemeError,
    );
  });
});

describe("resolveSequential", () => {
  it("snaps a position to the nearest step and clamps out-of-range positions", () => {
    const steps = NEUTRAL_PALETTE.sequential.light;
    expect(resolveSequential(NEUTRAL_PALETTE, 0, "light")).toBe(steps[0]);
    expect(resolveSequential(NEUTRAL_PALETTE, 1, "light")).toBe(
      steps[steps.length - 1],
    );
    expect(resolveSequential(NEUTRAL_PALETTE, 2, "light")).toBe(
      steps[steps.length - 1],
    );
    expect(resolveSequential(NEUTRAL_PALETTE, -1, "light")).toBe(steps[0]);
  });

  it("orients the dark ramp so low magnitude recedes toward the dark surface", () => {
    expect(resolveSequential(NEUTRAL_PALETTE, 0, "dark")).toBe("#0d366b");
    expect(resolveSequential(NEUTRAL_PALETTE, 1, "dark")).toBe("#cde2fb");
  });
});

describe("resolveSemantic", () => {
  it("resolves the reserved status scale", () => {
    expect(resolveSemantic(NEUTRAL_PALETTE, "good", "light")).toBe("#0ca30c");
    expect(resolveSemantic(NEUTRAL_PALETTE, "critical", "dark")).toBe(
      "#d03b3b",
    );
  });
});

describe("NEUTRAL_PALETTES", () => {
  it("indexes the default palette by name", () => {
    expect(NEUTRAL_PALETTES["neutral"]).toBe(NEUTRAL_PALETTE);
  });
});
