import {
  contrastRatio,
  deltaE,
  isHexColor,
  oklchLightnessChroma,
  type CvdKind,
} from "./color.js";
import { type ColorMode, type ModeColor, type Palette } from "./palette.js";

/**
 * Which check produced an issue.
 */
export type ValidationCheck =
  | "invalid-shape"
  | "invalid-color"
  | "lightness-band"
  | "chroma-floor"
  | "categorical-distinctness"
  | "normal-vision-distinctness"
  | "surface-contrast";

/**
 * `error` fails the pass; `warning` is a condition that is acceptable only with
 * a secondary encoding (direct labels, gaps, or texture) and does not fail it.
 */
export type ValidationSeverity = "error" | "warning";

export interface ValidationIssue {
  check: ValidationCheck;
  severity: ValidationSeverity;
  message: string;
  details: Record<string, unknown>;
}

/**
 * The structured outcome of a validation pass. `valid` is true when no
 * error-severity issue was found; warnings may still be present.
 */
export interface ValidationReport {
  valid: boolean;
  issues: ValidationIssue[];
}

/**
 * OKLCH lightness band per mode. Outside it a categorical hue is too pale or too
 * dark to sit as a mark on that mode's surface.
 */
export const LIGHTNESS_BAND: Readonly<
  Record<ColorMode, readonly [number, number]>
> = {
  light: [0.43, 0.77],
  dark: [0.48, 0.67],
};

/** OKLCH chroma floor. Below it a hue reads as grey and stops doing identity work. */
export const CHROMA_FLOOR = 0.1;

/** Adjacent-pair CVD delta E target: at or above it the pair is distinct. */
export const CVD_DISTINCTNESS_TARGET = 8;

/**
 * Adjacent-pair CVD delta E floor: between it and the target the pair is a
 * warning (legal only with a secondary encoding); below it, an error.
 */
export const CVD_DISTINCTNESS_FLOOR = 6;

/** Worst adjacent-pair delta E under normal vision: below it, an error. */
export const NORMAL_VISION_FLOOR = 15;

/** WCAG contrast ratio a mark should clear against its surface. */
export const SURFACE_CONTRAST_MINIMUM = 3;

export interface ValidateCategoricalOptions {
  /** The surface the colours render on, for the contrast check. */
  surface: string;
}

/**
 * Validate an ordered list of categorical colours, already resolved for one
 * mode, against the computable checks: the lightness band and chroma floor per
 * colour, adjacent-pair distinctness under colour-vision deficiency and under
 * normal vision, and contrast against the surface. Returns a structured report
 * and never throws: a malformed hex value (a colour or the surface) is reported
 * as an error issue and left out of the numeric checks, so a runtime-supplied
 * client palette with a typo is reported rather than crashing the pass.
 */
export function validateCategorical(
  colors: string[],
  mode: ColorMode,
  options: ValidateCategoricalOptions,
): ValidationReport {
  const issues: ValidationIssue[] = [];
  const [low, high] = LIGHTNESS_BAND[mode];

  // Malformed colours cannot be measured, so report each as an error and run
  // the numeric checks over the well-formed ones only.
  const validColors: string[] = [];
  for (const color of colors) {
    if (isHexColor(color)) {
      validColors.push(color);
    } else {
      issues.push({
        check: "invalid-color",
        severity: "error",
        message: `Colour "${color}" is not a six-digit hex value such as "#2a78d6".`,
        details: { color },
      });
    }
  }

  const surfaceValid = isHexColor(options.surface);
  if (!surfaceValid) {
    issues.push({
      check: "invalid-color",
      severity: "error",
      message: `Surface "${options.surface}" is not a six-digit hex value such as "#fcfcfb".`,
      details: { surface: options.surface },
    });
  }

  for (const color of validColors) {
    const { lightness, chroma } = oklchLightnessChroma(color);
    if (lightness < low || lightness > high) {
      issues.push({
        check: "lightness-band",
        severity: "error",
        message: `Colour ${color} has OKLCH lightness ${lightness.toFixed(3)}, outside the ${mode} band ${low} to ${high}.`,
        details: { color, lightness, mode, band: [low, high] },
      });
    }
    if (chroma < CHROMA_FLOOR) {
      issues.push({
        check: "chroma-floor",
        severity: "error",
        message: `Colour ${color} has OKLCH chroma ${chroma.toFixed(3)}, below the floor ${CHROMA_FLOOR}, so it reads as grey.`,
        details: { color, chroma },
      });
    }
  }

  for (let i = 0; i + 1 < validColors.length; i += 1) {
    const a = validColors[i]!;
    const b = validColors[i + 1]!;

    const cvdKinds: CvdKind[] = ["protan", "deutan"];
    let worstCvd = Infinity;
    let worstKind: CvdKind = "protan";
    for (const kind of cvdKinds) {
      const distance = deltaE(a, b, kind);
      if (distance < worstCvd) {
        worstCvd = distance;
        worstKind = kind;
      }
    }
    if (worstCvd < CVD_DISTINCTNESS_FLOOR) {
      issues.push({
        check: "categorical-distinctness",
        severity: "error",
        message: `Adjacent slots ${a} and ${b} are ${worstCvd.toFixed(1)} apart under ${worstKind} vision, below the floor ${CVD_DISTINCTNESS_FLOOR}.`,
        details: { a, b, deltaE: worstCvd, kind: worstKind },
      });
    } else if (worstCvd < CVD_DISTINCTNESS_TARGET) {
      issues.push({
        check: "categorical-distinctness",
        severity: "warning",
        message: `Adjacent slots ${a} and ${b} are ${worstCvd.toFixed(1)} apart under ${worstKind} vision, between the floor ${CVD_DISTINCTNESS_FLOOR} and the target ${CVD_DISTINCTNESS_TARGET}; a secondary encoding is required.`,
        details: { a, b, deltaE: worstCvd, kind: worstKind },
      });
    }

    const normal = deltaE(a, b);
    if (normal < NORMAL_VISION_FLOOR) {
      issues.push({
        check: "normal-vision-distinctness",
        severity: "error",
        message: `Adjacent slots ${a} and ${b} are ${normal.toFixed(1)} apart under normal vision, below the floor ${NORMAL_VISION_FLOOR}.`,
        details: { a, b, deltaE: normal },
      });
    }
  }

  if (surfaceValid) {
    for (const color of validColors) {
      const ratio = contrastRatio(color, options.surface);
      if (ratio < SURFACE_CONTRAST_MINIMUM) {
        issues.push({
          check: "surface-contrast",
          severity: "warning",
          message: `Colour ${color} sits at ${ratio.toFixed(2)}:1 against the surface ${options.surface}, below ${SURFACE_CONTRAST_MINIMUM}:1; visible labels or a table view are required.`,
          details: { color, surface: options.surface, ratio },
        });
      }
    }
  }

  return { valid: issues.every((issue) => issue.severity !== "error"), issues };
}

/**
 * Validate a palette's categorical slots for one mode against its own surface. A
 * runtime-supplied palette may be incomplete, so each value is read defensively:
 * a missing or non-string entry becomes an invalid-color issue rather than a
 * thrown error, keeping the never-throws contract of the validation pass.
 */
export function validatePalette(
  palette: Palette,
  mode: ColorMode,
): ValidationReport {
  // A runtime palette parsed from JSON could be null or not an object at all.
  const candidate: unknown = palette;
  if (typeof candidate !== "object" || candidate === null) {
    return {
      valid: false,
      issues: [
        {
          check: "invalid-shape",
          severity: "error",
          message: "The palette is not an object.",
          details: {},
        },
      ],
    };
  }
  const readMode = (color: ModeColor | undefined): string => {
    const value = color?.[mode];
    return typeof value === "string" ? value : "";
  };
  // A runtime palette that omits its categorical colours, gives a non-array, or
  // gives an empty array is a shape error: it has no usable series slot, and
  // every resolveCategorical call on it would throw.
  if (!Array.isArray(palette.categorical) || palette.categorical.length === 0) {
    return {
      valid: false,
      issues: [
        {
          check: "invalid-shape",
          severity: "error",
          message: "The palette has no categorical colours to validate.",
          details: {},
        },
      ],
    };
  }
  const colors = palette.categorical.map((entry) => readMode(entry));
  return validateCategorical(colors, mode, {
    surface: readMode(palette.surface),
  });
}
