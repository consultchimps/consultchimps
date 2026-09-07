import { ThemeError } from "./errors.js";

/**
 * Colour maths behind the validation checks. The conversions and the
 * colour-vision-deficiency (CVD) simulation reproduce the ConsultChimps
 * `dataviz` skill's palette validator so a palette that passes here passes
 * there too: sRGB to linear light, linear light to OKLab and OKLCH, the WCAG
 * relative-luminance contrast ratio, and the Machado, Oliveira and Fernandes
 * (2009) CVD transforms at severity 1.0. Distances reported as "delta E" are
 * Euclidean distance in OKLab multiplied by 100, matching the skill.
 *
 * Everything here is pure and platform-neutral: no dependency, no global state.
 */

export type RgbTriple = [number, number, number];

/**
 * A `min(protanopia, deuteranopia)` pair is what the categorical-distinctness
 * check gates on; tritanopia is reported alongside for context.
 */
export type CvdKind = "protan" | "deutan" | "tritan";

// Machado, Oliveira and Fernandes (2009) CVD transforms at severity 1.0, in
// linear RGB. The distinctness thresholds are calibrated to this model, so the
// model is part of the check rather than an implementation detail.
const MACHADO: Record<CvdKind, readonly RgbTriple[]> = {
  protan: [
    [0.152286, 1.052583, -0.204868],
    [0.114503, 0.786281, 0.099216],
    [-0.003882, -0.048116, 1.051998],
  ],
  deutan: [
    [0.367322, 0.860646, -0.227968],
    [0.280085, 0.672501, 0.047413],
    [-0.01182, 0.04294, 0.968881],
  ],
  tritan: [
    [1.255528, -0.076749, -0.178779],
    [-0.078411, 0.930809, 0.147602],
    [0.004733, 0.691367, 0.3039],
  ],
};

/**
 * Whether a string is a six-digit hex colour (`#rrggbb` or `rrggbb`). The
 * validation pass uses this to report a malformed colour as a structured issue
 * rather than letting `parseHexColor` throw.
 */
export function isHexColor(value: string): boolean {
  return /^#?[0-9a-fA-F]{6}$/.test(value.trim());
}

/**
 * Parse `#rrggbb` (or `rrggbb`) to sRGB components in the 0 to 1 range. Throws a
 * `ThemeError` on anything else so a typo never propagates a silent `NaN`
 * through the checks, which would let a run pass by accident.
 */
export function parseHexColor(hex: string): RgbTriple {
  const trimmed = hex.trim().replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6}$/.test(trimmed)) {
    throw new ThemeError(
      "THEME_INVALID_COLOR",
      `Colour "${hex}" is not a six-digit hex value such as "#2a78d6".`,
      { color: hex },
    );
  }
  return [0, 2, 4].map(
    (offset) => parseInt(trimmed.slice(offset, offset + 2), 16) / 255,
  ) as RgbTriple;
}

function srgbToLinearComponent(channel: number): number {
  return channel <= 0.04045
    ? channel / 12.92
    : ((channel + 0.055) / 1.055) ** 2.4;
}

function toLinear(hex: string): RgbTriple {
  return parseHexColor(hex).map(srgbToLinearComponent) as RgbTriple;
}

function relativeLuminanceOf([r, g, b]: RgbTriple): number {
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/**
 * The WCAG relative luminance of a colour, used by the contrast ratio.
 */
export function relativeLuminance(hex: string): number {
  return relativeLuminanceOf(toLinear(hex));
}

/**
 * The WCAG contrast ratio between two colours, from 1 (identical) to 21 (black
 * on white). Order does not matter.
 */
export function contrastRatio(a: string, b: string): number {
  const [high, low] = [relativeLuminance(a), relativeLuminance(b)].sort(
    (x, y) => y - x,
  ) as [number, number];
  return (high + 0.05) / (low + 0.05);
}

function oklabFromLinear([r, g, b]: RgbTriple): RgbTriple {
  const l = Math.cbrt(0.4122214708 * r + 0.5363325363 * g + 0.0514459929 * b);
  const m = Math.cbrt(0.2119034982 * r + 0.6806995451 * g + 0.1073969566 * b);
  const s = Math.cbrt(0.0883024619 * r + 0.2817188376 * g + 0.6299787005 * b);
  return [
    0.2104542553 * l + 0.793617785 * m - 0.0040720468 * s,
    1.9779984951 * l - 2.428592205 * m + 0.4505937099 * s,
    0.0259040371 * l + 0.7827717662 * m - 0.808675766 * s,
  ];
}

/**
 * The OKLCH lightness (0 to 1) and chroma of a colour. Lightness drives the
 * per-mode band check; chroma below the floor reads as grey and stops carrying
 * identity.
 */
export function oklchLightnessChroma(hex: string): {
  lightness: number;
  chroma: number;
} {
  const [lightness, a, b] = oklabFromLinear(toLinear(hex));
  return { lightness, chroma: Math.hypot(a, b) };
}

function simulateCvd(hex: string, kind: CvdKind): RgbTriple {
  const [r, g, b] = toLinear(hex);
  const matrix = MACHADO[kind];
  const clamp = (channel: number): number => Math.max(0, Math.min(1, channel));
  return [
    clamp(matrix[0]![0] * r + matrix[0]![1] * g + matrix[0]![2] * b),
    clamp(matrix[1]![0] * r + matrix[1]![1] * g + matrix[1]![2] * b),
    clamp(matrix[2]![0] * r + matrix[2]![1] * g + matrix[2]![2] * b),
  ];
}

/**
 * Perceptual distance between two colours as OKLab delta E times 100. Pass a
 * `kind` to measure the distance as a reader with that colour-vision deficiency
 * would perceive it; omit it for unsimulated (normal) vision.
 */
export function deltaE(a: string, b: string, kind?: CvdKind): number {
  const first = oklabFromLinear(kind ? simulateCvd(a, kind) : toLinear(a));
  const second = oklabFromLinear(kind ? simulateCvd(b, kind) : toLinear(b));
  return (
    100 *
    Math.hypot(first[0] - second[0], first[1] - second[1], first[2] - second[2])
  );
}
