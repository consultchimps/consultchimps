import type { Palette } from "./palette.js";

/**
 * Neutral placeholder palettes. These carry no brand or client identity: they
 * exist so the model, the validation pass, and downstream consumers have a
 * working default to build and test against. A client's real colours are
 * supplied at runtime by the consumer and are never committed here.
 *
 * The values reproduce the ConsultChimps `dataviz` skill's validated default
 * palette, which was chosen and ordered to pass the same contrast and
 * categorical-distinctness checks this package exposes.
 */

/**
 * The default neutral palette: eight categorical slots, a blue magnitude ramp,
 * and the reserved status scale, each with light and dark values.
 */
export const NEUTRAL_PALETTE: Palette = {
  name: "neutral",
  surface: { light: "#fcfcfb", dark: "#1a1a19" },
  ink: {
    primary: { light: "#0b0b0b", dark: "#ffffff" },
    secondary: { light: "#52514e", dark: "#c3c2b7" },
    muted: { light: "#898781", dark: "#898781" },
  },
  categorical: [
    { light: "#2a78d6", dark: "#3987e5" },
    { light: "#eb6834", dark: "#d95926" },
    { light: "#1baf7a", dark: "#199e70" },
    { light: "#eda100", dark: "#c98500" },
    { light: "#e87ba4", dark: "#d55181" },
    { light: "#008300", dark: "#008300" },
    { light: "#4a3aa7", dark: "#9085e9" },
    { light: "#e34948", dark: "#e66767" },
  ],
  sequential: {
    // Light mode: low magnitude is pale and recedes toward the light surface.
    light: [
      "#cde2fb",
      "#b7d3f6",
      "#9ec5f4",
      "#86b6ef",
      "#6da7ec",
      "#5598e7",
      "#3987e5",
      "#2a78d6",
      "#256abf",
      "#1c5cab",
      "#184f95",
      "#104281",
      "#0d366b",
    ],
    // Dark mode: the same steps oriented so low magnitude recedes toward the
    // dark surface and high magnitude reads bright.
    dark: [
      "#0d366b",
      "#104281",
      "#184f95",
      "#1c5cab",
      "#256abf",
      "#2a78d6",
      "#3987e5",
      "#5598e7",
      "#6da7ec",
      "#86b6ef",
      "#9ec5f4",
      "#b7d3f6",
      "#cde2fb",
    ],
  },
  // Status is a fixed scale with reserved meaning; the same steps clear the
  // 3:1 mark on both surfaces, so light and dark share them.
  semantic: {
    good: { light: "#0ca30c", dark: "#0ca30c" },
    warning: { light: "#fab219", dark: "#fab219" },
    serious: { light: "#ec835a", dark: "#ec835a" },
    critical: { light: "#d03b3b", dark: "#d03b3b" },
  },
};

/**
 * Every neutral placeholder palette this package ships, keyed by name.
 */
export const NEUTRAL_PALETTES: Readonly<Record<string, Palette>> = {
  [NEUTRAL_PALETTE.name]: NEUTRAL_PALETTE,
};
