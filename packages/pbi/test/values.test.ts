import { describe, expect, it } from "vitest";
import {
  currencyText,
  dateSerialText,
  exactMilliseconds,
  oversizedBinaryValues,
  significantDigits,
  toBase64,
  toCell,
  truncateCellText,
  dateColumnReadable,
} from "../src/values.js";

/** Section C, "Value typing", and Decision 6's three checks. */

const cell = (
  value: Parameters<typeof toCell>[0],
  type: Parameters<typeof toCell>[1],
) => toCell(value, type);

describe("numeric cells under the fifteen-digit rule", () => {
  it("writes an ordinary int64 as a number", () => {
    expect(cell(42n, "int64")).toEqual({
      cell: { kind: "number", text: "42", style: 0 },
    });
  });

  it("writes a sixteen-digit integer as exact text", () => {
    expect(cell(1234567890123456n, "int64")).toEqual({
      cell: { kind: "text", text: "1234567890123456" },
      reason: "PBI_NUMERIC_AS_TEXT",
    });
  });

  it("keeps an int64 past the safe-integer range exactly, as text", () => {
    const value = 9007199254740993n;
    expect(cell(value, "int64")).toEqual({
      cell: { kind: "text", text: "9007199254740993" },
      reason: "PBI_NUMERIC_AS_TEXT",
    });
  });

  it("writes currency with four decimals and a four-decimal format", () => {
    expect(cell(12_3456n, "currency")).toEqual({
      cell: { kind: "number", text: "12.3456", style: 2 },
    });
    expect(currencyText(-5000n)).toBe("-0.5000");
    expect(currencyText(0n)).toBe("0.0000");
  });

  it("writes a sixteen-digit currency value as exact text", () => {
    // 123456789012.3456 has sixteen significant digits.
    expect(cell(1234567890123456n, "currency")).toEqual({
      cell: { kind: "text", text: "123456789012.3456" },
      reason: "PBI_NUMERIC_AS_TEXT",
    });
  });

  it("writes out-of-range doubles as text", () => {
    expect(cell(1e308, "double")).toEqual({
      cell: { kind: "text", text: "1e+308" },
      reason: "PBI_NUMERIC_AS_TEXT",
    });
    expect(cell(1e-310, "double").reason).toBe("PBI_NUMERIC_AS_TEXT");
    expect(cell(2.2251e-308, "double").cell.kind).toBe("number");
  });

  it("writes negative zero as the text -0", () => {
    expect(cell(-0, "double")).toEqual({
      cell: { kind: "text", text: "-0" },
      reason: "PBI_NUMERIC_AS_TEXT",
    });
    expect(cell(0, "double").cell.kind).toBe("number");
  });

  it("writes non-finite doubles with the defined spellings", () => {
    expect(cell(Number.NaN, "double")).toEqual({
      cell: { kind: "text", text: "NaN" },
      reason: "PBI_NONFINITE_AS_TEXT",
    });
    expect(cell(Infinity, "double").cell).toEqual({
      kind: "text",
      text: "Infinity",
    });
    expect(cell(-Infinity, "double").cell).toEqual({
      kind: "text",
      text: "-Infinity",
    });
  });

  it("counts significant digits without leading or trailing zeros or exponent", () => {
    expect(significantDigits("0.00012300")).toBe(3);
    expect(significantDigits("1.5e+308")).toBe(2);
    expect(significantDigits("-1000")).toBe(1);
  });

  it("leaves null a blank cell that counts nothing", () => {
    expect(cell(null, "double")).toEqual({ cell: { kind: "blank" } });
  });
});

describe("date cells", () => {
  it("rounds the exact product once, half to the later instant", () => {
    expect(exactMilliseconds(1).milliseconds).toBe(86_400_000n);
    expect(exactMilliseconds(0.5)).toEqual({
      milliseconds: 43_200_000n,
      exact: true,
    });
    // 86,400,000 is 2^10 times 84,375, so an exact half millisecond needs a
    // serial of an odd multiple of 2^-11. 1/2048 of a day is 42,187.5 ms.
    const result = exactMilliseconds(1 / 2048);
    expect(result.exact).toBe(false);
    expect(result.milliseconds).toBe(42_188n);
  });

  it("writes an ordinary date as a serial with the date format", () => {
    // 43831 days from 1899-12-30 is 2020-01-01; both sides of 1900-03-01 agree.
    const result = cell(43831, "dateTimeSerial");
    expect(result.cell).toEqual({ kind: "number", text: "43831", style: 1 });
    expect(result.reason).toBeUndefined();
  });

  it("subtracts one day before 1900-03-01 and never generates 1900-02-29", () => {
    // Day 2 from the model epoch is 1900-01-01, Excel serial 1.
    expect(cell(2, "dateTimeSerial").cell).toEqual({
      kind: "number",
      text: "1",
      style: 1,
    });
    // Day 60 is 1900-02-28, which must not land on Excel's fictitious 60.
    expect(cell(60, "dateTimeSerial").cell).toEqual({
      kind: "number",
      text: "59",
      style: 1,
    });
    expect(cell(61, "dateTimeSerial").cell).toEqual({
      kind: "number",
      text: "61",
      style: 1,
    });
  });

  it("selects text below serial 2 without rounding", () => {
    expect(cell(1.5, "dateTimeSerial")).toEqual({
      cell: { kind: "text", text: "pbi-date-serial:1.5" },
      reason: "PBI_DATE_AS_TEXT",
    });
    expect(cell(-0, "dateTimeSerial").cell).toEqual({
      kind: "text",
      text: "pbi-date-serial:-0",
    });
  });

  it("selects text past year 9999", () => {
    expect(cell(2_958_466, "dateTimeSerial").reason).toBe("PBI_DATE_AS_TEXT");
    expect(cell(2_958_465, "dateTimeSerial").reason).toBeUndefined();
  });

  it("counts a changed millisecond as rounded, never also as text", () => {
    const serial = 43831 + 1 / 2048;
    const result = cell(serial, "dateTimeSerial");
    expect(result.reason).toBe("PBI_DATE_ROUNDED");
    expect(result.cell.kind).toBe("number");
  });

  it("preserves the fallback serial bit for bit", () => {
    const serial = 1.1000000000000001;
    const text = dateSerialText(serial);
    expect(Number(text.slice("pbi-date-serial:".length))).toBe(serial);
  });

  it("marks a column with a non-finite serial unreadable", () => {
    expect(dateColumnReadable([1, null, 2])).toBe(true);
    expect(dateColumnReadable([1, Number.NaN])).toBe(false);
    expect(dateColumnReadable([Infinity])).toBe(false);
  });
});

describe("text and binary cells", () => {
  it("truncates at the code-unit limit without splitting a surrogate pair", () => {
    const value = `${"a".repeat(32_766)}\u{1f600}`;
    const { text, truncated } = truncateCellText(value);
    expect(truncated).toBe(true);
    expect(text.length).toBe(32_766);
  });

  it("stops before the 254th line break even below the code-unit limit", () => {
    const value = "x\n".repeat(300);
    const { text, truncated } = truncateCellText(value);
    expect(truncated).toBe(true);
    expect([...text].filter((ch) => ch === "\n")).toHaveLength(253);
  });

  it("never splits a CRLF pair", () => {
    const value = `${"y".repeat(32_766)}\r\nz`;
    const { text } = truncateCellText(value);
    expect(text.endsWith("\r")).toBe(false);
    expect(text.length).toBe(32_766);
  });

  it("counts one truncation per changed cell", () => {
    const result = cell("z".repeat(40_000), "string");
    expect(result.reason).toBe("PBI_TEXT_TRUNCATED");
    expect((result.cell as { text: string }).text.length).toBe(32_767);
  });

  it("encodes binary as standard padded base64", () => {
    expect(toBase64(new Uint8Array([]))).toBe("");
    expect(toBase64(new Uint8Array([102, 111, 111]))).toBe("Zm9v");
    expect(toBase64(new Uint8Array([102, 111]))).toBe("Zm8=");
    expect(toBase64(new Uint8Array([102]))).toBe("Zg==");
    expect(cell(new Uint8Array([1, 2]), "binary")).toEqual({
      cell: { kind: "text", text: "AQI=" },
      reason: "PBI_BINARY_AS_BASE64",
    });
  });

  it("counts an oversized base64 value rather than truncating it", () => {
    // 8,191 base64 quartets hold 24,573 bytes in 32,764 characters; one more
    // byte needs a 8,192nd quartet and passes the 32,767-character limit.
    expect(oversizedBinaryValues([new Uint8Array(24_573)])).toBe(0);
    expect(oversizedBinaryValues([new Uint8Array(24_574), null])).toBe(1);
  });

  it("writes booleans as native cells and null as blank", () => {
    expect(cell(true, "boolean")).toEqual({
      cell: { kind: "boolean", value: true },
    });
    expect(cell(false, "boolean").reason).toBeUndefined();
    expect(cell(null, "boolean")).toEqual({ cell: { kind: "blank" } });
  });
});
