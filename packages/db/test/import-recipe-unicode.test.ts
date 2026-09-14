import { expect, test } from "vitest";

import { draftImportProfile } from "../src/import/profile.js";
import type { ImportSource } from "../src/import/types.js";
import { validateImportProfile } from "../src/validators.js";

function source(key: string): ImportSource {
  return {
    key,
    readerVersion: "synthetic-profile-unicode-1",
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
          throw new Error("Profile drafting must not open a source selection");
        },
      },
    ],
  };
}

function inferredDestination(
  profile: Awaited<ReturnType<typeof draftImportProfile>>,
) {
  validateImportProfile(profile);
  const destination = profile.routes[0]?.destination;
  expect(destination?.kind).toBe("new-table-infer");
  if (destination?.kind !== "new-table-infer") {
    throw new Error("Drafted route did not infer a new table");
  }
  return destination;
}

test("drafts a valid prefix from a leading supplementary letter", async () => {
  const destination = inferredDestination(
    await draftImportProfile({ sources: [source("𐐀Data")] }),
  );
  expect(destination).toMatchObject({
    name: "𐐀Data",
    recordId: { prefix: "𐐀", padding: 6 },
  });
});

test("drafts multiword prefixes by code point and keeps ASCII behavior", async () => {
  const unicode = inferredDestination(
    await draftImportProfile({
      sources: [source("source")],
      naming: { kind: "single-table", name: "𐐀 𐐁 𐐂 𐐃 𐐄 𐐅 𐐆 𐐇 𐐈" },
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
    await draftImportProfile({
      sources: [source("source")],
      naming: { kind: "single-table", name: "north sales data" },
    }),
  );
  expect(ascii).toMatchObject({
    name: "north_sales_data",
    recordId: { prefix: "NSD", padding: 6 },
  });
});

test("selection-label table prefixes preserve whole code points before uppercasing", async () => {
  const input = source("source");
  const selection = input.selections[0];
  if (selection === undefined) {
    throw new Error("Expected a synthetic selection");
  }
  const profile = await draftImportProfile({
    sources: [
      {
        ...input,
        selections: [
          { ...selection, label: "AAAAAAAßtail" },
          { ...selection, key: "unicode", label: "ABCDEF𐐨GHI" },
        ],
      },
    ],
    naming: { kind: "selection-label" },
  });
  validateImportProfile(profile);

  expect(
    profile.routes.map((route) => {
      expect(route.destination.kind).toBe("new-table-infer");
      if (route.destination.kind !== "new-table-infer") {
        throw new Error("Drafted route did not infer a new table");
      }
      return route.destination.recordId.prefix;
    }),
  ).toEqual(["AAAAAAASS", "ABCDEF𐐀G"]);
});

test("truncates table suggestions without splitting a supplementary letter", async () => {
  const included = `A${"a".repeat(197)}𐐀`;
  expect(included.length).toBe(200);
  expect(
    inferredDestination(
      await draftImportProfile({
        sources: [source("source")],
        naming: { kind: "single-table", name: included },
      }),
    ).name,
  ).toBe(included);

  const excluded = `A${"a".repeat(198)}𐐀`;
  expect(excluded.length).toBe(201);
  const truncated = inferredDestination(
    await draftImportProfile({
      sources: [source("source")],
      naming: { kind: "single-table", name: excluded },
    }),
  ).name;
  expect(truncated).toBe(`A${"a".repeat(198)}`);
  expect(truncated.length).toBe(199);
});
