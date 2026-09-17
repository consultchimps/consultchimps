import { describe, expect, it } from "vitest";
import {
  buildTable,
  decodePage,
  expandEncodeArray,
  HuffmanError,
  swapPairs,
} from "../src/huffman.js";

/**
 * Section C, the Huffman rows. Every case here is a crafted page: the assertion
 * is always that the reader refuses with its own error, that no `RangeError` or
 * plain `Error` escapes, and that no number the file declared appears in what
 * the reader says.
 */

/** A complete two-symbol code: symbols 0 and 1, one bit each. */
function twoSymbolTable(): ReturnType<typeof buildTable> {
  const packed = new Uint8Array(128);
  packed[0] = 0x11;
  return buildTable(expandEncodeArray(packed));
}

function refusal(run: () => unknown): HuffmanError {
  try {
    run();
  } catch (error) {
    // A RangeError or a bare Error here is the defect, not the refusal.
    expect(error).toBeInstanceOf(HuffmanError);
    return error as HuffmanError;
  }
  throw new Error("expected a refusal");
}

/** Nothing the file declared may appear in the message a caller could log. */
function carriesNoDeclaredValue(error: Error, declared: number): void {
  const text = `${error.message} ${JSON.stringify(error)}`;
  expect(declared).toBeGreaterThan(1000);
  expect(text).not.toContain(String(declared));
  expect(text).not.toContain(declared.toString(16));
}

describe("a page cannot decode past its own bytes", () => {
  it("refuses the reviewer's amplification probe", () => {
    // Four bytes of page declaring five million bits. Before the ceiling this
    // decoded 5,000,000 bytes in 125 ms, an amplification of 1,250,000 to one,
    // with no refusal and no budget.
    const page = swapPairs(new Uint8Array(4));
    const error = refusal(() =>
      decodePage(page, twoSymbolTable(), [0], 5_000_000, 0),
    );
    expect(error.message).toContain("page bit count");
    carriesNoDeclaredValue(error, 5_000_000);
  });

  it("refuses the largest value the field can hold", () => {
    const page = swapPairs(new Uint8Array(4));
    const error = refusal(() =>
      decodePage(page, twoSymbolTable(), [0], 0xffffffff, 0),
    );
    carriesNoDeclaredValue(error, 0xffffffff);
  });

  it("refuses a bit count one past the page", () => {
    const page = swapPairs(new Uint8Array(4));
    expect(() => decodePage(page, twoSymbolTable(), [0], 32, 0)).not.toThrow();
    refusal(() => decodePage(page, twoSymbolTable(), [0], 33, 0));
  });

  it("refuses an offset past the declared bit count", () => {
    const page = swapPairs(new Uint8Array(4));
    refusal(() => decodePage(page, twoSymbolTable(), [0, 40], 32, 0));
    refusal(() => decodePage(page, twoSymbolTable(), [8, 0], 32, 0));
    refusal(() => decodePage(page, twoSymbolTable(), [-1], 32, 0));
  });

  it("caps the run at one symbol per bit, charset byte included", () => {
    // A page of all-zero bytes decodes to symbol 0 repeatedly. Sixteen bits of
    // page with a charset byte is at most thirty-two output bytes, never more.
    const page = swapPairs(new Uint8Array(2));
    const runs = decodePage(page, twoSymbolTable(), [0], 16, 0x04);
    expect(runs[0]!.length).toBe(32);
  });
});

describe("the code table", () => {
  it("refuses an over-subscribed code", () => {
    const packed = new Uint8Array(128);
    // Three symbols of length one cannot fit a one-bit alphabet.
    packed[0] = 0x11;
    packed[1] = 0x01;
    refusal(() => buildTable(expandEncodeArray(packed)));
  });

  it("refuses an incomplete code", () => {
    const packed = new Uint8Array(128);
    packed[0] = 0x01; // one symbol of length one leaves a codeword unassigned
    refusal(() => buildTable(expandEncodeArray(packed)));
  });

  it("returns an empty table when no symbol has a length", () => {
    expect(buildTable(expandEncodeArray(new Uint8Array(128))).maxLength).toBe(
      0,
    );
  });

  it("refuses a corrupt bitstream that reaches an unassigned codeword", () => {
    // A complete four-symbol code, then a page whose bits select a codeword the
    // table does not carry is impossible; an incomplete table is the reachable
    // shape, and buildTable refuses it first. This asserts the guard directly.
    const table = { table: new Uint16Array(2), maxLength: 1 };
    const page = swapPairs(new Uint8Array(2));
    refusal(() => decodePage(page, table, [0], 16, 0));
  });
});

describe("the bit-order helpers", () => {
  it("swaps byte pairs and copies a trailing odd byte", () => {
    expect([...swapPairs(new Uint8Array([1, 2, 3, 4, 5]))]).toEqual([
      2, 1, 4, 3, 5,
    ]);
  });

  it("expands 128 nibble-packed bytes into 256 lengths", () => {
    const packed = new Uint8Array(128);
    packed[0] = 0x21;
    const lengths = expandEncodeArray(packed);
    expect(lengths.length).toBe(256);
    expect(lengths[0]).toBe(1);
    expect(lengths[1]).toBe(2);
  });
});
