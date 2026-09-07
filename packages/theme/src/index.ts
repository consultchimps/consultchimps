export { ThemeError, isThemeError } from "./errors.js";
export {
  contrastRatio,
  deltaE,
  oklchLightnessChroma,
  parseHexColor,
  relativeLuminance,
  type CvdKind,
  type RgbTriple,
} from "./color.js";
export {
  categoricalCount,
  resolveCategorical,
  resolveInk,
  resolveMode,
  resolveSemantic,
  resolveSequential,
  resolveSurface,
  type ColorMode,
  type InkRole,
  type ModeColor,
  type Palette,
  type SemanticRole,
  type SequentialRamp,
} from "./palette.js";
export { NEUTRAL_PALETTE, NEUTRAL_PALETTES } from "./palettes.js";
export {
  validateCategorical,
  validatePalette,
  CHROMA_FLOOR,
  CVD_DISTINCTNESS_FLOOR,
  CVD_DISTINCTNESS_TARGET,
  LIGHTNESS_BAND,
  NORMAL_VISION_FLOOR,
  SURFACE_CONTRAST_MINIMUM,
  type ValidateCategoricalOptions,
  type ValidationCheck,
  type ValidationIssue,
  type ValidationReport,
  type ValidationSeverity,
} from "./validation.js";
