import { describe, expect, it } from "vitest";
import {
  decodeSegmentIds,
  parseDictionary,
  parseIdf,
  parseIdfmeta,
  VertipaqAlignmentError,
  VertipaqError,
  XM_FIRST_DATA_ID,
} from "../src/vertipaq.js";
import {
  idf,
  idfmeta,
  numericDictionary,
  stringDictionary,
  unsupportedDictionary,
} from "./vertipaq-bytes.js";

/**
 * Section C, the column-decoding rows. Every allocation in this layer is sized
 * from a field in an untrusted member, so each test lies about one field and
 * asserts that the reader refuses with its own error before it allocates,
 * without a `RangeError` and without repeating the number it was told.
 */

function refusal(run: () => unknown): VertipaqError {
  try {
    run();
  } catch (error) {
    expect(error).toBeInstanceOf(VertipaqError);
    expect(error).not.toBeInstanceOf(RangeError);
    return error as VertipaqError;
  }
  throw new Error("expected a refusal");
}

function carriesNoDeclaredValue(error: Error, declared: number): void {
  expect(declared).toBeGreaterThan(1000);
  const text = `${error.message} ${JSON.stringify(error)}`;
  expect(text).not.toContain(String(declared));
}

describe(".idfmeta descriptors", () => {
  it("reads a well-formed descriptor", () => {
    const descriptors = parseIdfmeta(idfmeta([{ records: 9 }]), 9);
    expect(descriptors).toHaveLength(1);
    expect(descriptors[0]!.records).toBe(9);
    expect(descriptors[0]!.bitWidth).toBe(2);
  });

  it("refuses a record count past the catalog's row count", () => {
    // The reference implementation would push this many nulls or size a vector
    // from it. It is an alignment failure, not a corrupt member.
    const member = idfmeta([{ records: 9_007_199_254_740_991 }]);
    const error = refusal(() => parseIdfmeta(member, 9));
    expect(error).toBeInstanceOf(VertipaqAlignmentError);
    carriesNoDeclaredValue(error, 9_007_199_254_740_991);
  });

  it("refuses segments whose records sum past the row count", () => {
    const member = idfmeta([{ records: 6 }, { records: 6 }]);
    expect(parseIdfmeta(member, 12)).toHaveLength(2);
    expect(refusal(() => parseIdfmeta(member, 11))).toBeInstanceOf(
      VertipaqAlignmentError,
    );
  });

  it("refuses a subsegment count past the row count", () => {
    const member = idfmeta([{ records: 4, subsegmentRecords: 40_000 }]);
    const error = refusal(() => parseIdfmeta(member, 4));
    expect(error).toBeInstanceOf(VertipaqAlignmentError);
    carriesNoDeclaredValue(error, 40_000);
  });

  it("refuses a segment count the member cannot hold", () => {
    const member = idfmeta([{ records: 1 }]);
    const view = new DataView(member.buffer);
    view.setBigUint64(6, 5_000_000n, true);
    const error = refusal(() => parseIdfmeta(member, 1));
    expect(error.message).toContain("segment count");
    carriesNoDeclaredValue(error, 5_000_000);
  });

  it("refuses a truncated member", () => {
    const member = idfmeta([{ records: 4 }]);
    refusal(() => parseIdfmeta(member.subarray(0, member.length - 10), 4));
  });
});

describe(".idf run-length segments", () => {
  it("reads well-formed entries", () => {
    const segments = parseIdf(
      idf([{ runs: [[5, 3] as const, [6, 2] as const] }]),
    );
    expect(segments).toHaveLength(1);
    expect([...segments[0]!.dataValues]).toEqual([5, 6]);
    expect([...segments[0]!.repeatValues]).toEqual([3, 2]);
  });

  it("refuses an entry count the member's bytes cannot supply", () => {
    // Eight bytes per entry, so a member of a few dozen bytes cannot hold
    // millions of them. Before the ceiling this allocated two typed arrays of
    // that length before reading a single entry.
    const member = idf([
      { runs: [[1, 1] as const], declaredEntries: 4_000_000 },
    ]);
    const error = refusal(() => parseIdf(member));
    expect(error.message).toContain("entry count");
    carriesNoDeclaredValue(error, 4_000_000);
  });

  it("refuses the largest count the field can hold", () => {
    const member = idf([{ runs: [], declaredEntries: 9_007_199_254_740_991 }]);
    carriesNoDeclaredValue(
      refusal(() => parseIdf(member)),
      9_007_199_254_740_991,
    );
  });

  it("refuses a sub-word count the member cannot supply", () => {
    const member = idf([{ runs: [[1, 1] as const], subWords: 2_000_000 }]);
    const error = refusal(() => parseIdf(member));
    expect(error.message).toContain("word count");
    carriesNoDeclaredValue(error, 2_000_000);
  });
});

describe("segment decoding", () => {
  it("expands run-length entries to the declared records", () => {
    const descriptor = parseIdfmeta(idfmeta([{ records: 5 }]), 5)[0]!;
    const segment = parseIdf(
      idf([{ runs: [[7, 3] as const, [8, 2] as const] }]),
    )[0]!;
    expect([...decodeSegmentIds(segment, descriptor)]).toEqual([7, 7, 7, 8, 8]);
  });

  it("caps the expansion at the declared records", () => {
    const descriptor = parseIdfmeta(idfmeta([{ records: 4 }]), 4)[0]!;
    const segment = parseIdf(
      idf([{ runs: [[7, 3] as const, [8, 9] as const] }]),
    )[0]!;
    expect([...decodeSegmentIds(segment, descriptor)]).toEqual([7, 7, 7, 8]);
  });

  it("refuses a segment that declares no record count", () => {
    // Without a cap the sum of four-byte repeat values would size the vector.
    const descriptor = parseIdfmeta(idfmeta([{ records: 0 }]), 0)[0]!;
    const segment = parseIdf(
      idf([{ runs: [[1, 4_000_000_000] as const] }]),
    )[0]!;
    const error = refusal(() => decodeSegmentIds(segment, descriptor));
    expect(error.message).toContain("record count");
    carriesNoDeclaredValue(error, 4_000_000_000);
  });

  it("refuses a bit-packed subsegment whose width is unknowable", () => {
    // An unknown compression class leaves the width at zero. Before the guard
    // this reached a Float64Array sized by an infinite per-word count.
    const descriptor = parseIdfmeta(
      idfmeta([
        {
          records: 4,
          compressionClass: 999_999,
          subCompressionClass: 999_999,
          subsegmentRecords: 4,
        },
      ]),
      4,
    )[0]!;
    expect(descriptor.bitWidth).toBe(0);
    const segment = parseIdf(
      idf([{ runs: [[0xffffffff, 4] as const], sub: new Uint8Array(8) }]),
    )[0]!;
    refusal(() => decodeSegmentIds(segment, descriptor));
  });
});

describe("dictionaries", () => {
  it("reads a numeric dictionary and keeps its bigints", () => {
    const dictionary = parseDictionary(
      numericDictionary(0, [1n, 9_007_199_254_740_993n]),
      XM_FIRST_DATA_ID,
    );
    expect(dictionary!.values).toEqual([1n, 9_007_199_254_740_993n]);
    expect(dictionary!.isString).toBe(false);
  });

  it("refuses a value count the member's bytes cannot supply", () => {
    const member = numericDictionary(0, [1n], { declaredCount: 3_000_000 });
    const error = refusal(() => parseDictionary(member, XM_FIRST_DATA_ID));
    expect(error.message).toContain("value count");
    carriesNoDeclaredValue(error, 3_000_000);
  });

  it("returns null for a dictionary type it does not support", () => {
    // The caller turns this into PBI_COLUMN_UNSUPPORTED_ENCODING.
    expect(
      parseDictionary(unsupportedDictionary(7), XM_FIRST_DATA_ID),
    ).toBeNull();
  });

  it("reads an uncompressed string page", () => {
    const dictionary = parseDictionary(
      stringDictionary([{ text: ["alpha", "beta"] }], []),
      XM_FIRST_DATA_ID,
    );
    expect(dictionary!.values).toEqual(["alpha", "beta"]);
    expect(dictionary!.isString).toBe(true);
  });

  it("refuses a page count the member cannot hold", () => {
    const member = stringDictionary([{ text: ["a"] }], []);
    new DataView(member.buffer).setBigInt64(45, 900_000n, true);
    const error = refusal(() => parseDictionary(member, XM_FIRST_DATA_ID));
    expect(error.message).toContain("page count");
    carriesNoDeclaredValue(error, 900_000);
  });

  it("refuses a Huffman page whose declared bits run past its buffer", () => {
    const encodeArray = new Uint8Array(128);
    encodeArray[0] = 0x11; // a complete two-symbol code
    const member = stringDictionary(
      [
        {
          huffman: {
            encodeArray,
            buffer: new Uint8Array(4),
            totalBits: 5_000_000,
          },
        },
      ],
      [[0, 0]],
    );
    const error = refusal(() => parseDictionary(member, XM_FIRST_DATA_ID));
    // The Huffman refusal is converted before it leaves this layer.
    expect(error.message).toContain("string page");
    carriesNoDeclaredValue(error, 5_000_000);
  });

  it("refuses a Huffman page with an incomplete code", () => {
    const encodeArray = new Uint8Array(128);
    encodeArray[0] = 0x01; // one symbol of length one: incomplete
    const member = stringDictionary(
      [
        {
          huffman: { encodeArray, buffer: new Uint8Array(4), totalBits: 8 },
        },
      ],
      [[0, 0]],
    );
    refusal(() => parseDictionary(member, XM_FIRST_DATA_ID));
  });
});

describe("an exhausted bit-packed subsegment", () => {
  it("refuses rather than filling the rest of the column with zeros", () => {
    // The descriptor claims four bit-packed entries and the subsegment holds
    // one word. Filling the shortfall with zeros would keep the row count and
    // export fabricated ids as if they were real values.
    const descriptor = parseIdfmeta(
      idfmeta([
        {
          records: 200,
          compressionClass: 703066,
          subCompressionClass: 703062, // 32 bits: two values per word
          subsegmentRecords: 200,
        },
      ]),
      200,
    )[0]!;
    const segment = parseIdf(
      idf([{ runs: [[0xffffffff, 200] as const], sub: new Uint8Array(16) }]),
    )[0]!;
    const error = refusal(() => decodeSegmentIds(segment, descriptor));
    expect(error.message).toContain("exhausted");
  });

  it("decodes a subsegment that holds every entry it claims", () => {
    const descriptor = parseIdfmeta(
      idfmeta([
        {
          records: 4,
          compressionClass: 703066,
          subCompressionClass: 703062,
          subsegmentRecords: 4,
          minDataId: 0,
        },
      ]),
      4,
    )[0]!;
    const segment = parseIdf(
      idf([{ runs: [[0xffffffff, 4] as const], sub: new Uint8Array(16) }]),
    )[0]!;
    expect([...decodeSegmentIds(segment, descriptor)]).toHaveLength(4);
  });
});

describe("the hybrid run-length placeholder marker", () => {
  /**
   * The marker is `0xFFFFFFFF` minus the number of packed values already
   * consumed, not minus the run's own index. The corpus settles it: across the
   * nine fixtures, 497 segments carry a bit-packed subsegment, 142 of them
   * place their first marker after one or more ordinary runs, and in every one
   * of those the marker is the bare `0xFFFFFFFF` regardless of its index.
   * `f-2026-corpspend`'s "Country/Region ID" has its first marker at run 8 with
   * `0xFFFFFFFF`, and its second at run 11 with `0xFFFFFFE3`, which is
   * `0xFFFFFFFF` less the 28 values the first consumed.
   */
  function hybrid(
    runs: readonly (readonly [number, number])[],
    packedValues: number,
  ) {
    const descriptor = parseIdfmeta(
      idfmeta([
        {
          records: runs.reduce((total, [, repeat]) => total + repeat, 0),
          compressionClass: 703066,
          subCompressionClass: 703062, // 32 bits: two values per 64-bit word
          subsegmentRecords: packedValues,
          minDataId: 0,
        },
      ]),
      1000,
    )[0]!;
    // Two packed values per word, ascending so each one is identifiable.
    const sub = new Uint8Array(Math.ceil(packedValues / 2) * 8);
    const view = new DataView(sub.buffer);
    for (let index = 0; index < packedValues; index++)
      view.setUint32(index * 4, 100 + index, true);
    const segment = parseIdf(idf([{ runs, sub }]))[0]!;
    return [...decodeSegmentIds(segment, descriptor)];
  }

  it("recognises a placeholder that follows an ordinary run", () => {
    // One ordinary run first, then the marker. Under a rule keyed on the run's
    // index the marker would have to be 0xFFFFFFFE here; it is not, and the
    // packed values would decode as one huge id and then a null.
    expect(
      hybrid(
        [
          [7, 3],
          [0xffffffff, 4],
        ],
        4,
      ),
    ).toEqual([7, 7, 7, 100, 101, 102, 103]);
  });

  it("recognises a placeholder that comes first", () => {
    expect(
      hybrid(
        [
          [0xffffffff, 4],
          [7, 2],
        ],
        4,
      ),
    ).toEqual([100, 101, 102, 103, 7, 7]);
  });

  it("offsets a later marker by the values already consumed", () => {
    // The shape of "Country/Region ID": a marker after ordinary runs, then a
    // second marker reduced by the first marker's own count.
    expect(
      hybrid(
        [
          [5, 2],
          [0xffffffff, 3],
          [6, 1],
          [0xffffffff - 3, 2],
        ],
        5,
      ),
    ).toEqual([5, 5, 100, 101, 102, 6, 103, 104]);
  });

  it("treats a value that is not the marker as an ordinary id", () => {
    // 0xFFFFFFFE with nothing consumed is data, not a placeholder.
    expect(hybrid([[0xfffffffe, 2]], 0)).toEqual([0xfffffffe, 0xfffffffe]);
  });
});
