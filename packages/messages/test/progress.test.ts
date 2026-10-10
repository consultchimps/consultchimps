import { expect, test } from "vitest";

import { progressStageLabel } from "../src/index.js";

test("a stage reads as words, naming inputs and outputs for workbook stages", () => {
  expect(progressStageLabel("reading-workbooks")).toBe("Reading inputs");
  expect(progressStageLabel("building-workbooks")).toBe("Building outputs");
  expect(progressStageLabel("staging-workbooks")).toBe("Staging outputs");
  expect(progressStageLabel("merging-inputs")).toBe("Merging inputs");
  expect(progressStageLabel("constructor")).toBe("Constructor");
});
