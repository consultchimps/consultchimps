import { expect, test } from "vitest";

import { draftImportProfile } from "../src/import/profile.js";
import type { ImportSource } from "../src/import/types.js";
import { validateImportProfile } from "../src/validators.js";

test("profile drafting rejects sources with no selected regions", async () => {
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

  await expect(draftImportProfile({ sources: [source] })).rejects.toMatchObject(
    {
      code: "DB_IMPORT_NO_SELECTIONS",
      message: expect.stringContaining("include hidden sheets"),
    },
  );
});

test("an explicit empty profile remains valid for exclusion workflows", () => {
  expect(() => validateImportProfile({ version: 1, routes: [] })).not.toThrow();
});
