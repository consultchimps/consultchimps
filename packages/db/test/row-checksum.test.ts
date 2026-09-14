import { expect, test } from "vitest";

import { createCaptureRowChecksum } from "../src/import/row-checksum.js";

type Row = readonly [sourceRow: bigint, valuesJson: string];

function checksum(groups: readonly (readonly Row[])[]): string {
  const result = createCaptureRowChecksum();
  for (const group of groups) {
    for (const [sourceRow, valuesJson] of group) {
      result.update(sourceRow, valuesJson);
    }
  }
  return result.digest();
}

test("defines a stable lowercase digest for an empty capture", () => {
  const digest = checksum([]);
  expect(digest).toMatch(/^[0-9a-f]{64}$/u);
  expect(digest).toBe(
    "e4459720b9abecbaa31d303e073bce51db984289deca370fa8ef220b0f1b4377",
  );
});

test("is independent of caller batch grouping and preserves Unicode bytes", () => {
  const rows: Row[] = [
    [2n, '{"Region":{"kind":"string","value":"North"}}'],
    [3n, '{"Region":{"kind":"string","value":"北部 🐒"}}'],
    [10n, '{"Amount":{"kind":"number","raw":"1.20"}}'],
  ];
  expect(checksum([rows])).toBe(checksum([[rows[0]!], rows.slice(1)]));
  expect(checksum([rows])).toBe(
    "c9efc1ab7f49d7936efa7b660fd4ea56e9d96d5a2bf404a327e29987ec341910",
  );
});

test("frames row coordinates and payload lengths without concatenation collisions", () => {
  expect(checksum([[[1n, "23"]]])).not.toBe(checksum([[[12n, "3"]]]));
  expect(
    checksum([
      [
        [1n, "a"],
        [2n, "bc"],
      ],
    ]),
  ).not.toBe(
    checksum([
      [
        [1n, "ab"],
        [2n, "c"],
      ],
    ]),
  );
});

test("changes when a coordinate, value, or row order changes", () => {
  const original: Row[] = [
    [2n, '{"Value":{"kind":"string","value":"North"}}'],
    [3n, '{"Value":{"kind":"string","value":"South"}}'],
  ];
  const digest = checksum([original]);
  expect(checksum([[[4n, original[0]![1]], original[1]!]])).not.toBe(digest);
  expect(
    checksum([[original[0]!, [original[1]![0], `${original[1]![1]} `]]]),
  ).not.toBe(digest);
  expect(checksum([[original[1]!, original[0]!]])).not.toBe(digest);
});

test("rejects coordinates outside the validated source-row domain", () => {
  for (const sourceRow of [0n, -1n, BigInt(Number.MAX_SAFE_INTEGER) + 1n]) {
    const result = createCaptureRowChecksum();
    expect(() => result.update(sourceRow, "{}"))
      .toThrowErrorMatchingInlineSnapshot(`
        [RangeError: The capture row checksum requires a positive safe source row.]
      `);
  }
});
