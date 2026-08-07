/**
 * A trivial exercise of the symmetry harness itself: it must run its body once
 * per shape, in order, and give each run its own suite.
 */
import { describe, expect, it } from "vitest";

import { describeBothShapes, SHAPES } from "./symmetry.js";
import type { CorpusShape } from "./fixtures.js";

const observed: CorpusShape[] = [];

describeBothShapes("symmetry harness", (shape) => {
  observed.push(shape);

  it("receives its own shape", () => {
    expect(SHAPES).toContain(shape);
  });
});

describe("symmetry harness: registration", () => {
  it("ran the body once per authored shape, in corpus order", () => {
    expect(observed).toEqual(["table", "range"]);
  });
});
