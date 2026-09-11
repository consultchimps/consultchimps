import { describe, expect, it } from "vitest";

import { cellText, encodeTsv, parseTsv } from "./workspace-tsv";

// The grammar both directions of the clipboard share. The line-ending cases are
// the reason this module exists: Excel on Windows writes CRLF, and a parser
// that splits on "\n" alone leaves a "\r" on the last field of every row, which
// is how a multi-row paste corrupts its last column.

describe("cellText", () => {
  it("renders a stored value as the text the clipboard carries", () => {
    expect(cellText("North")).toBe("North");
    expect(cellText(12)).toBe("12");
    expect(cellText(1.5)).toBe("1.5");
    expect(cellText(true)).toBe("true");
    expect(cellText(false)).toBe("false");
    expect(cellText("2026-01-31")).toBe("2026-01-31");
  });

  it("renders nothing for an empty cell", () => {
    expect(cellText(null)).toBe("");
    expect(cellText(undefined)).toBe("");
  });
});

describe("encodeTsv", () => {
  it("joins cells with tabs and terminates every row with CRLF", () => {
    expect(
      encodeTsv([
        ["a", "b"],
        ["c", "d"],
      ]),
    ).toBe("a\tb\r\nc\td\r\n");
  });

  it("terminates a single row too, which is what Excel's own clipboard does", () => {
    // The terminator is what makes a block ending in a blank row readable back:
    // without it, "North\r\n" would have to mean both one row and two.
    expect(encodeTsv([["North"]])).toBe("North\r\n");
  });

  it("writes a block of no rows as no text", () => {
    expect(encodeTsv([])).toBe("");
  });

  it("quotes a field holding a tab, a newline, a carriage return, or a quote", () => {
    expect(encodeTsv([["a\tb"]])).toBe('"a\tb"\r\n');
    expect(encodeTsv([["a\nb"]])).toBe('"a\nb"\r\n');
    expect(encodeTsv([["a\rb"]])).toBe('"a\rb"\r\n');
    expect(encodeTsv([['say "hi"']])).toBe('"say ""hi"""\r\n');
  });

  it("leaves an ordinary field unquoted", () => {
    expect(encodeTsv([["1 North Street", ""]])).toBe("1 North Street\t\r\n");
  });
});

describe("parseTsv", () => {
  it("reads Windows line endings, which is what Excel writes", () => {
    expect(parseTsv("a\tb\r\nc\td\r\n")).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  it("reads a line break as the end of the row before it, as Excel means it", () => {
    // A single cell copied from Excel arrives terminated. It is one row, not a
    // row and a blank one, and this is the case the rule is chosen for.
    expect(parseTsv("A\r\n")).toEqual([["A"]]);
    expect(parseTsv("A\tB\r\n")).toEqual([["A", "B"]]);
  });

  it("reads a second line break as a blank row, because the first ended a row", () => {
    expect(parseTsv("A\r\n\r\n")).toEqual([["A"], [""]]);
    expect(parseTsv("\r\n")).toEqual([[""]]);
    expect(parseTsv("\r\nA\r\n")).toEqual([[""], ["A"]]);
  });

  it("reads a bare carriage return and a bare newline alike", () => {
    expect(parseTsv("a\rb")).toEqual([["a"], ["b"]]);
    expect(parseTsv("a\nb")).toEqual([["a"], ["b"]]);
    expect(parseTsv("a\r\nb\nc\rd")).toEqual([["a"], ["b"], ["c"], ["d"]]);
  });

  it("keeps a row the text did not terminate", () => {
    expect(parseTsv("a")).toEqual([["a"]]);
    expect(parseTsv("a\r\nb")).toEqual([["a"], ["b"]]);
  });

  it("keeps an empty trailing field", () => {
    expect(parseTsv("a\t")).toEqual([["a", ""]]);
    expect(parseTsv("\ta")).toEqual([["", "a"]]);
  });

  it("reads a quoted field holding a tab or a newline as one field", () => {
    expect(parseTsv('"a\tb"\tc')).toEqual([["a\tb", "c"]]);
    expect(parseTsv('"line one\nline two"\tc')).toEqual([
      ["line one\nline two", "c"],
    ]);
    expect(parseTsv('"line one\r\nline two"')).toEqual([
      ["line one\r\nline two"],
    ]);
  });

  it("reads a doubled quote inside a quoted field as one quote", () => {
    expect(parseTsv('"say ""hi"""')).toEqual([['say "hi"']]);
    expect(parseTsv('"""quoted"""\tplain')).toEqual([['"quoted"', "plain"]]);
  });

  it("leaves a quote in the middle of an unquoted field alone", () => {
    expect(parseTsv('a"b\tc')).toEqual([['a"b', "c"]]);
    expect(parseTsv('12" pipe')).toEqual([['12" pipe']]);
  });

  it("takes the rest of the text as the field when a quote is never closed", () => {
    expect(parseTsv('"a\tb\nc')).toEqual([["a\tb\nc"]]);
  });

  it("reads text after a closing quote as part of the same field", () => {
    expect(parseTsv('"a"b\tc')).toEqual([["ab", "c"]]);
  });

  it("reads empty text as nothing at all", () => {
    expect(parseTsv("")).toEqual([]);
  });

  it("round-trips every block the encoder can produce", () => {
    const block = [
      ["North", "a\tb", 'say "hi"'],
      ["", "line one\r\nline two", "12"],
    ];

    expect(parseTsv(encodeTsv(block))).toEqual(block);
  });
});

describe("the round trip", () => {
  // The property the two halves exist to keep: whatever the grid copies, a
  // paste of it is the same block. The blank cases are the ones a copy runs
  // into and the ones the earlier encoding lost, so they are named rather than
  // left to a general claim.
  const blocks: ReadonlyArray<{
    readonly what: string;
    readonly block: string[][];
  }> = [
    { what: "one value", block: [["North"]] },
    {
      what: "a rectangle",
      block: [
        ["a", "b"],
        ["c", "d"],
      ],
    },
    { what: "a blank cell", block: [[""]] },
    { what: "a blank row after a value", block: [["A"], [""]] },
    { what: "a blank row before a value", block: [[""], ["A"]] },
    {
      what: "a blank last column",
      block: [
        ["A", ""],
        ["B", ""],
      ],
    },
    {
      what: "a blank first column",
      block: [
        ["", "A"],
        ["", "B"],
      ],
    },
    {
      what: "nothing but blanks",
      block: [
        ["", ""],
        ["", ""],
      ],
    },
    { what: "no rows at all", block: [] },
    {
      what: "fields holding the separators themselves",
      block: [
        ["a\tb", "c\r\nd"],
        ['say "hi"', ""],
      ],
    },
    {
      what: "a trailing blank row of several cells",
      block: [
        ["a", "b"],
        ["", ""],
      ],
    },
  ];

  for (const { what, block } of blocks) {
    it(`survives a copy and a paste of ${what}`, () => {
      expect(parseTsv(encodeTsv(block))).toEqual(block);
    });
  }

  it("stays the same block however many times it goes round", () => {
    const block = [["A"], [""]];
    const once = encodeTsv(block);

    expect(encodeTsv(parseTsv(once))).toBe(once);
    expect(parseTsv(encodeTsv(parseTsv(once)))).toEqual(block);
  });
});
