import { expect, test } from "vitest";

import { draftImportRecipe } from "../src/import/recipe.js";
import type { ImportSource } from "../src/import/types.js";
import { validateImportRecipe } from "../src/validators.js";

function source(key: string): ImportSource {
  return {
    key,
    readerVersion: "synthetic-recipe-unicode-1",
    bytes: {
      name: "synthetic.xlsx",
      size: 0,
      async readAt() {
        return new Uint8Array();
      },
    },
    selections: [
      {
        key: "Data",
        label: "Data",
        async open() {
          throw new Error("Recipe drafting must not open a source selection");
        },
      },
    ],
  };
}

function inferredDestination(
  recipe: Awaited<ReturnType<typeof draftImportRecipe>>,
) {
  validateImportRecipe(recipe);
  const destination = recipe.routes[0]?.destination;
  expect(destination?.kind).toBe("new-table-infer");
  if (destination?.kind !== "new-table-infer") {
    throw new Error("Drafted route did not infer a new table");
  }
  return destination;
}

test("drafts a valid prefix from a leading supplementary letter", async () => {
  const destination = inferredDestination(
    await draftImportRecipe({ sources: [source("𐐀Data")] }),
  );
  expect(destination).toMatchObject({
    name: "𐐀Data",
    recordId: { prefix: "𐐀", padding: 6 },
  });
});

test("drafts multiword prefixes by code point and keeps ASCII behavior", async () => {
  const unicode = inferredDestination(
    await draftImportRecipe({
      sources: [source("source")],
      into: "𐐀 𐐁 𐐂 𐐃 𐐄 𐐅 𐐆 𐐇 𐐈",
    }),
  );
  expect(Array.from(unicode.recordId.prefix)).toEqual([
    "𐐀",
    "𐐁",
    "𐐂",
    "𐐃",
    "𐐄",
    "𐐅",
    "𐐆",
    "𐐇",
  ]);

  const ascii = inferredDestination(
    await draftImportRecipe({
      sources: [source("source")],
      into: "north sales data",
    }),
  );
  expect(ascii).toMatchObject({
    name: "north_sales_data",
    recordId: { prefix: "NSD", padding: 6 },
  });
});

test("truncates table suggestions without splitting a supplementary letter", async () => {
  const included = `A${"a".repeat(197)}𐐀`;
  expect(included.length).toBe(200);
  expect(
    inferredDestination(
      await draftImportRecipe({
        sources: [source("source")],
        into: included,
      }),
    ).name,
  ).toBe(included);

  const excluded = `A${"a".repeat(198)}𐐀`;
  expect(excluded.length).toBe(201);
  const truncated = inferredDestination(
    await draftImportRecipe({
      sources: [source("source")],
      into: excluded,
    }),
  ).name;
  expect(truncated).toBe(`A${"a".repeat(198)}`);
  expect(truncated.length).toBe(199);
});
