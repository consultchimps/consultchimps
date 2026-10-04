/**
 * Two corrections to numfmt 3.2.6's output, kept until upstream fixes them
 * (ADR 0006). Each applies only to the inputs numfmt gets wrong; everything
 * else is numfmt's answer. `test/numfmt-guards.test.ts` pins the upstream
 * behaviour and fails once a numfmt release fixes it: delete that guard then.
 *
 * TODO(numfmt > 3.2.6): link the upstream issue or pull request once filed.
 */

/**
 * Guard (a): General truncates a number with 10 or 11 integer digits and a
 * fraction (`4403928373.5` shows `4403928373`). Excel shows at most 11
 * characters, so it rounds such a number to an integer, and to `1E+11` when
 * rounding reaches 12 digits. Returns undefined when numfmt is right.
 */
export function generalRoundingGuard(
  pattern: string,
  value: number | string | boolean,
): string | undefined {
  if (typeof value !== "number" || pattern.toLowerCase() !== "general") {
    return undefined;
  }
  const magnitude = Math.abs(value);
  if (magnitude < 1e9 || magnitude >= 1e11 || Number.isInteger(value)) {
    return undefined;
  }
  const sign = value < 0 ? "-" : "";
  const rounded = Math.round(magnitude);
  return rounded >= 1e11 ? `${sign}1E+11` : `${sign}${String(rounded)}`;
}

/**
 * Whether a format shows a time: an hour or second token, which a minute
 * token always sits beside. A format showing only a date drops the time
 * rather than rounding it, so nothing carries into the next day.
 */
const SHOWS_TIME = /[hs]/iu;

/** A conditional section such as `[<1]`. */
const CONDITION = /\[\s*[<>=]/u;

/** The text of a format before its first unquoted, unescaped `;`. */
function firstSection(pattern: string): string {
  let quoted = false;
  for (let index = 0; index < pattern.length; index += 1) {
    const character = pattern[index];
    if (character === "\\" && !quoted) index += 1;
    else if (character === '"') quoted = !quoted;
    else if (character === ";" && !quoted) return pattern.slice(0, index);
  }
  return pattern;
}

/** The decimals of a format's seconds, as numfmt reads them for rounding. */
const SECOND_DECIMALS = /s+\.(0{1,3})/iu;

/**
 * Guard (b): a time that rounds up to midnight keeps the old date
 * (`45292.9999999` under `m/d/yy h:mm` shows `1/1/24 0:00`). numfmt rounds
 * the time to its smallest shown unit, and when that carries into the next
 * day the time resets but the date does not. Returns the serial numfmt should
 * have formatted, the next day's midnight, or undefined when no carry occurs.
 * Only called for a date format.
 */
export function midnightCarryGuard(
  pattern: string,
  value: number,
): number | undefined {
  if (!(value >= 0)) return undefined;
  // A positive value takes the first section. A conditional format picks its
  // section itself, which this guard does not second-guess.
  if (CONDITION.test(pattern)) return undefined;
  const section = firstSection(pattern);
  // Literal text, escaped characters and bracketed sections are not tokens.
  const bare = section.replace(/"[^"]*"|\\.|\[[^\]]*\]/gu, "");
  if (!SHOWS_TIME.test(bare)) return undefined;
  const day = Math.trunc(value);
  const seconds = 86_400 * (value - day);
  if (Math.floor(seconds) !== 86_399) return undefined;
  const fraction = seconds - 86_399;
  // numfmt's thresholds: a near-whole second always carries; otherwise the
  // second rounds up when the fraction rounds away at the shown precision.
  const decimals = SECOND_DECIMALS.exec(section.replace(/"[^"]*"/gu, ""))?.[1]
    ?.length;
  const threshold =
    decimals === 3
      ? 0.9995
      : decimals === 2
        ? 0.995
        : decimals === 1
          ? 0.95
          : 0.5;
  const carries =
    fraction > 0.9999 ||
    (decimals === undefined ? fraction >= threshold : fraction > threshold);
  return carries ? day + 1 : undefined;
}
