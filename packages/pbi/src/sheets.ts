import { MAX_SHEET_NAME_UNITS } from "./values.js";

/**
 * Deterministic worksheet-name allocation under ADR 0004 Decision 5. Every
 * comparison uses `String.prototype.toLowerCase()` with no locale argument, so
 * `I` then `i` allocates `I` then `i_2` under any host locale.
 */

const FORBIDDEN = /[[\]:*?/\\]/g;

/** Excel refuses this name outright, so it is treated as already claimed. */
const RESERVED = "history";

function sanitize(name: string): string {
  let out = "";
  for (let index = 0; index < name.length; index++) {
    const unit = name.charCodeAt(index);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const low = name.charCodeAt(index + 1);
      if (low >= 0xdc00 && low <= 0xdfff) {
        out += name.slice(index, index + 2);
        index++;
        continue;
      }
      out += "_";
      continue;
    }
    if (unit >= 0xdc00 && unit <= 0xdfff) {
      out += "_";
      continue;
    }
    if (
      unit <= 0x1f ||
      (unit >= 0x7f && unit <= 0x9f) ||
      unit === 0xfffe ||
      unit === 0xffff
    ) {
      out += "_";
      continue;
    }
    out += name[index];
  }
  out = out.replace(FORBIDDEN, "_");
  // Replacements happen before trimming, so a name of only forbidden
  // characters trims to nothing and takes the fallback.
  out = out.replace(/^[\s']+|[\s']+$/g, "");
  return out.length === 0 ? "Sheet" : out;
}

/** Keep the longest prefix that does not split a surrogate pair. */
function shorten(base: string, units: number): string {
  if (base.length <= units) return base;
  let cut = base.slice(0, units);
  const last = cut.charCodeAt(cut.length - 1);
  if (last >= 0xd800 && last <= 0xdbff) cut = cut.slice(0, -1);
  return cut.replace(/[\s']+$/g, "");
}

export class SheetNameAllocator {
  readonly #claimed = new Set<string>([RESERVED]);

  /**
   * Claim one worksheet name for a table part. `part` is one-based; the first
   * part carries no part suffix. Suffix space is reserved before truncation,
   * and each attempt starts from the original sanitized base.
   */
  claim(name: string, part: number): string {
    const base = sanitize(name);
    const partSuffix = part <= 1 ? "" : `_${part}`;
    for (let collision = 1; ; collision++) {
      const suffix = partSuffix + (collision === 1 ? "" : `_${collision}`);
      const candidate =
        shorten(base, MAX_SHEET_NAME_UNITS - suffix.length) + suffix;
      const key = candidate.toLowerCase();
      if (!this.#claimed.has(key)) {
        this.#claimed.add(key);
        return candidate;
      }
    }
  }
}
