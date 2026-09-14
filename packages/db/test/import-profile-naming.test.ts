import { expect, test } from "vitest";

import { draftImportProfile } from "../src/import/profile.js";
import type { ImportSource } from "../src/import/types.js";

function source(
  key: string,
  selections: readonly { readonly key: string; readonly label: string }[],
  opened: () => void,
): ImportSource {
  return {
    key,
    readerVersion: "synthetic-profile-naming-1",
    bytes: {
      name: `${key}.synthetic`,
      size: 0,
      async readAt() {
        return new Uint8Array();
      },
    },
    selections: selections.map((selection) => ({
      ...selection,
      async open() {
        opened();
        throw new Error("Profile drafting must not open source rows");
      },
    })),
  };
}

function destinations(
  profile: Awaited<ReturnType<typeof draftImportProfile>>,
): readonly { readonly name: string; readonly prefix: string }[] {
  return profile.routes.map((route) => {
    expect(route.destination.kind).toBe("new-table-infer");
    if (route.destination.kind !== "new-table-infer") {
      throw new Error("Expected an inferred destination");
    }
    return {
      name: route.destination.name,
      prefix: route.destination.recordId.prefix,
    };
  });
}

test("source-or-selection uses the source for one selection and labels for several", async () => {
  let openCount = 0;
  const profile = await draftImportProfile({
    sources: [
      source(
        "North source.csv",
        [{ key: "only", label: "Ignored label" }],
        () => {
          openCount += 1;
        },
      ),
      source(
        "Workbook",
        [
          { key: "north", label: "North sales" },
          { key: "south", label: "South returns" },
        ],
        () => {
          openCount += 1;
        },
      ),
    ],
  });

  expect(destinations(profile)).toEqual([
    { name: "North_source_csv", prefix: "NSC" },
    { name: "North_sales", prefix: "NS" },
    { name: "South_returns", prefix: "SR" },
  ]);
  expect(openCount).toBe(0);
});

test("selection-label uses the sanitized table name for its record prefix", async () => {
  let openCount = 0;
  const profile = await draftImportProfile({
    sources: [
      source(
        "Ignored source",
        [
          { key: "region", label: "  Q1 / Orders  " },
          { key: "inventory", label: "inventory" },
        ],
        () => {
          openCount += 1;
        },
      ),
    ],
    naming: { kind: "selection-label" },
  });

  expect(destinations(profile)).toEqual([
    { name: "Q1_Orders", prefix: "Q1_ORDER" },
    { name: "inventory", prefix: "INVENTOR" },
  ]);
  expect(openCount).toBe(0);
});

test("single-table applies the same sanitizer and prefix policy", async () => {
  let openCount = 0;
  const profile = await draftImportProfile({
    sources: [
      source(
        "Ignored source",
        [{ key: "region", label: "Ignored label" }],
        () => {
          openCount += 1;
        },
      ),
    ],
    naming: { kind: "single-table", name: "  Current / Orders  " },
  });

  expect(destinations(profile)).toEqual([
    { name: "Current_Orders", prefix: "CO" },
  ]);
  expect(openCount).toBe(0);
});

test("single-table rejects multiple selections without opening them", async () => {
  let openCount = 0;
  await expect(
    draftImportProfile({
      sources: [
        source(
          "Workbook",
          [
            { key: "north", label: "North" },
            { key: "south", label: "South" },
          ],
          () => {
            openCount += 1;
          },
        ),
      ],
      naming: { kind: "single-table", name: "Orders" },
    }),
  ).rejects.toMatchObject({ code: "DB_IMPORT_INTO_AMBIGUOUS" });
  expect(openCount).toBe(0);
});

test.each([
  null,
  { kind: "unsupported" },
  { kind: "single-table" },
  { kind: "single-table", name: 42 },
])(
  "rejects invalid runtime naming %# before opening source rows",
  async (naming) => {
    let openCount = 0;
    const input = {
      sources: [
        source("Source", [{ key: "region", label: "Region" }], () => {
          openCount += 1;
        }),
      ],
      naming,
    };

    await expect(
      Reflect.apply(draftImportProfile, undefined, [input]),
    ).rejects.toMatchObject({ code: "DB_INVALID_IMPORT_NAMING" });
    expect(openCount).toBe(0);
  },
);
