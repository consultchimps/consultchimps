import { expect, test } from "vitest";

import { draftImportRecipe } from "../src/import/recipe.js";
import type { ImportSource } from "../src/import/types.js";
import { validateImportRecipe } from "../src/validators.js";

test("recipe drafting rejects sources with no selected regions", async () => {
  const source: ImportSource = {
    key: "hidden-workbook",
    readerVersion: "synthetic-no-selections-1",
    bytes: {
      name: "hidden.xlsx",
      size: 0,
      async readAt() {
        return new Uint8Array();
      },
    },
    selections: [],
  };

  await expect(draftImportRecipe({ sources: [source] })).rejects.toMatchObject({
    code: "DB_IMPORT_NO_SELECTIONS",
    message: expect.stringContaining("include hidden sheets"),
  });
});

test("an explicit empty recipe remains valid for exclusion workflows", () => {
  expect(() => validateImportRecipe({ version: 1, routes: [] })).not.toThrow();
});
