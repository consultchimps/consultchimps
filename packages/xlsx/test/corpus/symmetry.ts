/**
 * Guardrail 1 of ARCHITECTURE.md: the symmetry harness.
 *
 * Corpus fixtures are authored in pairs - the same records as an Excel Table
 * and as a plain range - so "a capability implemented for only one shape is a
 * failing test, not a review comment". The existing corpus expresses that with
 * `it.each(SHAPES)` per assertion. This helper formalizes the same pairing one
 * level up, for suites whose setup differs per shape.
 *
 * Existing corpus tests are deliberately left on `it.each(SHAPES)`; this is
 * for tests written from here on.
 */
import { describe } from "vitest";

import type { CorpusShape } from "./fixtures.js";

/** Both authored shapes, in the order the corpus reports them. */
export const SHAPES: readonly CorpusShape[] = ["table", "range"];

/**
 * Run `fn` once per shape, each inside its own suite named `<name> (<shape>)`.
 *
 * Use it when a suite needs shape-specific setup (a differently built fixture,
 * a different expected geometry). When only the assertions differ, prefer the
 * corpus's existing `it.each(SHAPES)` form, which keeps the pair adjacent in
 * the reporter output.
 *
 * Asymmetric behavior still gets its own named pin rather than a branch inside
 * `fn`: per corpus/README.md, "the asymmetry gets its own named pin rather
 * than being hidden behind a branch".
 */
export function describeBothShapes(
  name: string,
  fn: (shape: CorpusShape) => void,
): void {
  for (const shape of SHAPES) {
    describe(`${name} (${shape})`, () => {
      fn(shape);
    });
  }
}
