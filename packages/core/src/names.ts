/**
 * Portable filename sanitization shared by every operation that derives an
 * output name from caller-supplied text. This module must stay free of
 * node:fs and node:path imports so browser entry points can use it.
 */

const UNSAFE_NAME_CHARACTERS = /[<>:"/\\|?*]+/gu;
const WINDOWS_RESERVED_FILENAME =
  /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\..*)?$/iu;

// Linear-time replacement for a trailing [. ]+ regex, which CodeQL flags as
// polynomial on adversarial inputs with long runs of spaces.
function trimTrailingDotsAndSpaces(value: string): string {
  let end = value.length;
  while (end > 0) {
    const character = value[end - 1];
    if (character !== "." && character !== " ") {
      break;
    }
    end -= 1;
  }
  return value.slice(0, end);
}

/**
 * Shorten a string so its UTF-8 encoding fits within maxBytes, cutting only at
 * code-point boundaries so no character is split into invalid bytes.
 */
export function truncateToUtf8Bytes(value: string, maxBytes: number): string {
  const encoder = new TextEncoder();
  let result = "";
  let total = 0;
  for (const character of value) {
    const size = encoder.encode(character).length;
    if (total + size > maxBytes) {
      break;
    }
    result += character;
    total += size;
  }
  return result;
}

function withoutControlCharacters(value: string): string {
  let result = "";
  for (const character of value) {
    const codePoint = character.codePointAt(0);
    if (codePoint !== undefined && codePoint >= 32 && codePoint !== 127) {
      result += character;
    }
  }
  return result;
}

// Used when both the value and the fallback sanitize away to nothing, so the
// function can promise a non-empty, portable result for every input pair.
const LAST_RESORT_NAME = "file";

/**
 * Apply the character rules and the byte cap. The result is empty when the
 * input carries nothing usable, which is what makes the fallback chain in
 * safeNameFragment observable.
 */
function sanitizedFragment(value: string): string {
  const normalized = trimTrailingDotsAndSpaces(
    withoutControlCharacters(value.normalize("NFKC"))
      .replace(UNSAFE_NAME_CHARACTERS, "-")
      .replace(/\s+/gu, " ")
      .replace(/-+/gu, "-")
      .trim(),
  );
  // Cap the fragment by encoded size, truncating at code-point boundaries,
  // so generated names plus their page suffix stay well inside common
  // 255-byte filename limits even for multi-byte scripts and emoji.
  return trimTrailingDotsAndSpaces(truncateToUtf8Bytes(normalized, 80));
}

/**
 * Reduce a caller-supplied name to a portable filename fragment. Byte
 * operations never touch a filesystem, but their output names become
 * download and archive entries that must stay valid everywhere.
 *
 * The fallback goes through the same rules as the value, so a caller-supplied
 * fallback cannot smuggle an unsafe or over-long name past the guarantee. A
 * fallback that is already a clean fragment is returned unchanged.
 */
export function safeNameFragment(value: string, fallback: string): string {
  const safe =
    sanitizedFragment(value) || sanitizedFragment(fallback) || LAST_RESORT_NAME;
  return WINDOWS_RESERVED_FILENAME.test(safe) ? `_${safe}` : safe;
}
