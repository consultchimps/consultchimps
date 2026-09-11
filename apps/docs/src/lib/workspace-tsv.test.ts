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
  it("joins cells with tabs and rows with CRLF, with no trailing newline", () => {
    expect(
      encodeTsv([
        ["a", "b"],
        ["c", "d"],
      ]),
    ).toBe("a\tb\r\nc\td");
  });

  it("writes a single cell as itself", () => {
    expect(encodeTsv([["North"]])).toBe("North");
  });

  it("quotes a field holding a tab, a newline, a carriage return, or a quote", () => {
    expect(encodeTsv([["a\tb"]])).toBe('"a\tb"');
    expect(encodeTsv([["a\nb"]])).toBe('"a\nb"');
    expect(encodeTsv([["a\rb"]])).toBe('"a\rb"');
    expect(encodeTsv([['say "hi"']])).toBe('"say ""hi"""');
  });

  it("leaves an ordinary field unquoted", () => {
    expect(encodeTsv([["1 North Street", ""]])).toBe("1 North Street\t");
  });
});

describe("parseTsv", () => {
  it("reads Windows line endings, which is what Excel writes", () => {
    expect(parseTsv("a\tb\r\nc\td\r\n")).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  it("reads a bare carriage return and a bare newline alike", () => {
    expect(parseTsv("a\rb")).toEqual([["a"], ["b"]]);
    expect(parseTsv("a\nb")).toEqual([["a"], ["b"]]);
    expect(parseTsv("a\r\nb\nc\rd")).toEqual([["a"], ["b"], ["c"], ["d"]]);
  });

  it("drops one trailing row separator and no more", () => {
    expect(parseTsv("a\r\n")).toEqual([["a"]]);
    expect(parseTsv("a\r\n\r\n")).toEqual([["a"], [""]]);
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
