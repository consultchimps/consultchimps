import { ThemeError } from "./errors.js";

/**
 * Which set of values a palette resolves to. Every colour role carries a light
 * and a dark value; `mode` selects between them.
 */
export type ColorMode = "light" | "dark";

/**
 * Reserved semantic roles for state. Status colours are a fixed, small scale
 * with reserved meaning; they never stand in for a categorical series.
 */
export type SemanticRole = "good" | "warning" | "serious" | "critical";

/**
 * Text and chrome ink roles that sit on the palette surface.
 */
export type InkRole = "primary" | "secondary" | "muted";

/**
 * A single colour role in both modes. The two values are chosen together for
 * their own surface, not derived from one another at runtime.
 */
export interface ModeColor {
  light: string;
  dark: string;
}

/**
 * A magnitude ramp for continuous or ordered data. Steps run from the lowest
 * magnitude to the highest. Each mode has its own orientation so the low end
 * recedes toward that mode's surface.
 */
export interface SequentialRamp {
  light: string[];
  dark: string[];
}

/**
 * A named palette. `categorical` is an ordered list of series slots, assigned in
 * order and never cycled; `sequential` encodes magnitude; `semantic` holds the
 * reserved status scale. A consumer supplies a client's real palette at runtime;
 * only neutral placeholder palettes are committed to this repository.
 */
export interface Palette {
  name: string;
  surface: ModeColor;
  ink: Record<InkRole, ModeColor>;
  categorical: ModeColor[];
  sequential: SequentialRamp;
  semantic: Record<SemanticRole, ModeColor>;
}

/**
 * Resolve a single mode-aware colour to its value for the given mode.
 */
export function resolveMode(color: ModeColor, mode: ColorMode): string {
  return color[mode];
}

/**
 * The palette surface (chart or panel background) for a mode.
 */
export function resolveSurface(palette: Palette, mode: ColorMode): string {
  return palette.surface[mode];
}

/**
 * An ink role (text or chrome) for a mode.
 */
export function resolveInk(
  palette: Palette,
  role: InkRole,
  mode: ColorMode,
): string {
  return palette.ink[role][mode];
}

/**
 * A reserved status colour for a mode.
 */
export function resolveSemantic(
  palette: Palette,
  role: SemanticRole,
  mode: ColorMode,
): string {
  return palette.semantic[role][mode];
}

/**
 * How many categorical series slots the palette defines.
 */
export function categoricalCount(palette: Palette): number {
  return palette.categorical.length;
}

/**
 * The categorical series colour at a zero-based slot for a mode. Slots are
 * assigned in order and never cycled, so an out-of-range index is a caller
 * mistake (a series past the palette's capacity folds into "Other" instead) and
 * throws rather than silently wrapping.
 */
export function resolveCategorical(
  palette: Palette,
  index: number,
  mode: ColorMode,
): string {
  const count = palette.categorical.length;
  if (!Number.isInteger(index) || index < 0 || index >= count) {
    throw new ThemeError(
      "THEME_CATEGORICAL_INDEX_OUT_OF_RANGE",
      `Categorical slot ${index} is out of range for palette "${palette.name}", which defines ${count} slots (0 to ${count - 1}).`,
      { index, count, palette: palette.name },
    );
  }
  return palette.categorical[index]![mode];
}

/**
 * A sequential ramp colour for a position from 0 (lowest magnitude) to 1
 * (highest), snapped to the nearest defined step for the mode. Positions outside
 * 0 to 1 are clamped.
 */
export function resolveSequential(
  palette: Palette,
  position: number,
  mode: ColorMode,
): string {
  const steps = palette.sequential[mode];
  if (steps.length === 0) {
    throw new ThemeError(
      "THEME_SEQUENTIAL_RAMP_EMPTY",
      `Palette "${palette.name}" has no ${mode} sequential steps.`,
      { palette: palette.name, mode },
    );
  }
  const clamped = Number.isNaN(position)
    ? 0
    : Math.max(0, Math.min(1, position));
  const stepIndex = Math.round(clamped * (steps.length - 1));
  return steps[stepIndex]!;
}
