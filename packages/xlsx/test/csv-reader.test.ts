/**
 * The CSV reader (ADR 0007): how bytes become text, rows and cell values.
 */
import { ConsultChimpsError } from "@consultchimps/core";
import { describe, expect, it } from "vitest";

import { settleCsvOptions, type CsvReadOptions } from "../src/csv/options.js";
import {
  CSV_MAX_OPEN_ROW_CHARS,
  CSV_PIECE_BYTES,
  CsvWorkbook,
  csvSheetName,
  isCsvName,
} from "../src/csv/reader.js";
import { CsvNumberReader, csvDate } from "../src/csv/typing.js";
import type { StreamedCell } from "../src/operations/consolidate/reader.js";
import { bytesSource } from "../src/package/index.js";

const utf8 = (text: string): Uint8Array => new TextEncoder().encode(text);

function utf16(text: string, littleEndian: boolean): Uint8Array {
  const bytes = new Uint8Array(2 + text.length * 2);
  bytes.set(littleEndian ? [0xff, 0xfe] : [0xfe, 0xff]);
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    const [first, second] = littleEndian
      ? [code & 0xff, code >> 8]
      : [code >> 8, code & 0xff];
    bytes[2 + index * 2] = first;
    bytes[3 + index * 2] = second;
  }
  return bytes;
}

/** Text as Windows-1252 bytes, for the characters these tests use. */
function windows1252(text: string): Uint8Array {
  const special: Record<string, number> = { "€": 0x80, "’": 0x92 };
  return Uint8Array.from(
    [...text].map((char) => special[char] ?? char.charCodeAt(0)),
  );
}

interface Read {
  workbook: CsvWorkbook;
  rows: Array<[number, StreamedCell[]]>;
  values: unknown[][];
}

async function read(
  bytes: Uint8Array,
  options?: CsvReadOptions,
  readOptions: { text?: boolean; occupancy?: boolean } = {},
  name = "data.csv",
): Promise<Read & { range: unknown; occupied?: boolean }> {
  const workbook = await CsvWorkbook.open(
    bytesSource(name, bytes),
    { file: name, source: name, details: { source: name } },
    options,
  );
  const rows: Array<[number, StreamedCell[]]> = [];
  const result = await workbook.readWorksheet(
    workbook.sheets[0]!,
    {
      begin: () => {
        rows.length = 0;
      },
      row: (row, cells) => {
        rows.push([row, [...cells]]);
      },
    },
    readOptions,
  );
  // Each row as a dense list of its values, blank cells as undefined.
  const values: unknown[][] = [];
  for (const [row, cells] of rows) {
    const dense: unknown[] = [];
    for (const cell of cells) dense[cell.column] = cell.value;
    values[row] = dense;
  }
  return {
    workbook,
    rows,
    values,
    range: result.range,
    ...(result.occupied === undefined ? {} : { occupied: result.occupied }),
  };
}

async function failure(
  bytes: Uint8Array,
  options?: CsvReadOptions,
): Promise<ConsultChimpsError> {
  try {
    await read(bytes, options);
  } catch (error) {
    if (error instanceof ConsultChimpsError) return error;
    throw error;
  }
  throw new Error("The read was expected to fail.");
}

describe("CSV structure", () => {
  it("keeps quoted line breaks, doubled quotes and delimiters in their fields", async () => {
    const { values } = await read(
      utf8(
        'id,note\r\n1,"line one\r\nline two"\r\n2,"say ""hi"""\r\n3,"a,b"\r\n',
      ),
    );
    expect(values).toEqual([
      ["id", "note"],
      ["1", "line one\r\nline two"],
      ["2", 'say "hi"'],
      ["3", "a,b"],
    ]);
  });

  it("reads a quoted field that spans pieces", async () => {
    const long = "x".repeat(CSV_PIECE_BYTES + 5000);
    const { values } = await read(
      utf8(`id,note\n1,"${long}\n${long}"\n2,end\n`),
    );
    expect(values).toEqual([
      ["id", "note"],
      ["1", `${long}\n${long}`],
      ["2", "end"],
    ]);
  });

  it("decodes a character split between pieces whole", async () => {
    const padding = "y".repeat(CSV_PIECE_BYTES - 3);
    const { values } = await read(utf8(`a\n${padding}é\n`));
    expect(values).toEqual([["a"], [`${padding}é`]]);
  });

  it("reads every row the same wherever a piece ends", async () => {
    const tail = 'p,"q\r\nr",s\r\n"t""u",v\nw,"x"\r\n"y" ,z\r\n';
    const expected = [
      ["p", "q\r\nr", "s"],
      ['t"u', "v"],
      ["w", "x"],
      ["y", "z"],
    ];
    for (let shift = 0; shift <= tail.length; shift += 1) {
      const padding = `${"z".repeat(CSV_PIECE_BYTES - 2 - shift)}\r\n`;
      const { values } = await read(utf8(`${padding}${tail}`));
      expect(
        values.slice(1),
        `piece ends ${String(shift)} characters in`,
      ).toEqual(expected);
    }
  });

  it("reads ragged rows, blank lines and a missing final line break", async () => {
    const { values, range } = await read(utf8("a,b,c\n1\n\n1,2,3,4\n\n5,6"));
    expect(values).toEqual([
      ["a", "b", "c"],
      ["1"],
      undefined,
      ["1", "2", "3", "4"],
      undefined,
      ["5", "6"],
    ]);
    expect(range).toEqual({
      startRow: 0,
      endRow: 5,
      startColumn: 0,
      endColumn: 3,
    });
  });

  it("reads a file mixing CRLF and LF row by row", async () => {
    const { values } = await read(utf8("a,b\r\n1,2\n3,4\r\n5,6\n"));
    expect(values).toEqual([
      ["a", "b"],
      ["1", "2"],
      ["3", "4"],
      ["5", "6"],
    ]);
  });

  it("splits a file whose rows end in CR alone on CR", async () => {
    const { values } = await read(utf8('a,b\r1,"two\rlines"\r3,4\r'));
    expect(values).toEqual([
      ["a", "b"],
      ["1", "two\rlines"],
      ["3", "4"],
    ]);
  });

  it("decides the line ending by the first one in the file", async () => {
    // A CR file whose quoted text holds an LF, and an LF file whose first row
    // runs past the first piece.
    const crFile = await read(utf8('a,b\r1,"two\nlines"\r3,4\r'));
    expect(crFile.values).toEqual([
      ["a", "b"],
      ["1", "two\nlines"],
      ["3", "4"],
    ]);
    const long = "x".repeat(CSV_PIECE_BYTES + 10);
    const lfFile = await read(utf8(`a,"${long}"\n1,2\n`));
    expect(lfFile.values).toEqual([
      ["a", long],
      ["1", "2"],
    ]);
  });

  it("refuses a CR on its own in a file split on LF", async () => {
    const error = await failure(utf8("a,b\n1,2\rc,d\ne,f\n"));
    expect(error.code).toBe("XLSX_CSV_MALFORMED");
    expect(error.message).toContain("row 2 holds a carriage return on its own");
  });

  it("reads a CR ending a quoted last field as part of the line ending", async () => {
    const { values } = await read(utf8('a,b\nc,"x\r"\nd,e\n'));
    expect(values).toEqual([
      ["a", "b"],
      ["c", "x"],
      ["d", "e"],
    ]);
  });

  it("leaves empty fields blank and keeps spaces", async () => {
    const { rows } = await read(utf8('a,,"", b \n'));
    expect(rows).toEqual([
      [
        0,
        [
          { column: 0, value: "a" },
          { column: 3, value: " b " },
        ],
      ],
    ]);
  });

  it("describes an empty file as holding nothing", async () => {
    const { values, range } = await read(new Uint8Array());
    expect(values).toEqual([]);
    expect(range).toBeUndefined();
  });

  it("names its one visible worksheet after the file", async () => {
    const { workbook } = await read(
      utf8("a\n"),
      undefined,
      {},
      "North sales.CSV",
    );
    expect(workbook.sheets).toEqual([
      {
        name: "North sales",
        visible: true,
        visibility: "visible",
        part: "csv",
      },
    ]);
    expect(workbook.tables).toEqual([]);
    expect(workbook.names).toEqual([]);
    expect(csvSheetName("inputs/2025/q1.csv")).toBe("q1");
    expect(csvSheetName(".csv")).toBe("Sheet1");
    expect(csvSheetName("[Q1]: a*b?.csv")).toBe("_Q1__ a_b_");
    expect(csvSheetName("'quoted'.csv")).toBe("quoted");
    expect(csvSheetName("History.csv")).toBe("History_");
    expect(csvSheetName("''.csv")).toBe("Sheet1");
    expect(csvSheetName(`${"a".repeat(40)}.csv`)).toBe("a".repeat(31));
    expect(csvSheetName(`${"a".repeat(30)}'b.csv`)).toBe("a".repeat(30));
    expect(
      csvSheetName(`${"a".repeat(30)}${String.fromCodePoint(0x1f600)}.csv`),
    ).toBe("a".repeat(30));
    expect(isCsvName("a.CSV")).toBe(true);
    expect(isCsvName("a.xlsx")).toBe(false);
  });

  it("fingerprints the bytes", async () => {
    const first = await read(utf8("a,b\n1,2\n"));
    const same = await read(utf8("a,b\n1,2\n"));
    const changed = await read(utf8("a,b\n1,3\n"));
    expect(same.workbook.fingerprint).toBe(first.workbook.fingerprint);
    expect(changed.workbook.fingerprint).not.toBe(first.workbook.fingerprint);
  });
});

describe("CSV delimiters", () => {
  it.each([
    ["comma", "a,b\n1,2\n", ","],
    ["semicolon", "a;b;c\n1;2,5;3\n4;5;6\n", ";"],
    ["tab", "a\tb\n1\t2\n", "\t"],
    ["pipe", "a|b\n1|2\n", "|"],
    ["one column", "name\nalpha\nbeta\n", ","],
  ])("guesses a %s delimiter", async (_name, text, delimiter) => {
    const { workbook } = await read(utf8(text));
    expect(workbook.delimiter).toBe(delimiter);
  });

  it("looks past title lines that hold no delimiter", async () => {
    const titles = `Report title\n`.repeat(12);
    const { workbook } = await read(utf8(`${titles}a;b\n1;2\n`));
    expect(workbook.delimiter).toBe(";");
  });

  it("takes a chosen delimiter over the guess", async () => {
    const { workbook, values } = await read(utf8("a;b,c\n1;2,3\n"), {
      delimiter: ",",
    });
    expect(workbook.delimiter).toBe(",");
    expect(values).toEqual([
      ["a;b", "c"],
      ["1;2", "3"],
    ]);
  });
});

describe("CSV encodings", () => {
  const text = "name,city\r\nZoë,München\r\n";
  const expected = [
    ["name", "city"],
    ["Zoë", "München"],
  ];

  it("reads UTF-8 without a byte order mark", async () => {
    const { workbook, values } = await read(utf8(text));
    expect([workbook.encoding, workbook.encodingSource]).toEqual([
      "utf-8",
      "valid-utf-8",
    ]);
    expect(values).toEqual(expected);
    expect(workbook.warnings).toEqual([]);
  });

  it.each([
    ["UTF-8", Uint8Array.from([0xef, 0xbb, 0xbf, ...utf8(text)]), "utf-8"],
    ["UTF-16 LE", utf16(text, true), "utf-16le"],
    ["UTF-16 BE", utf16(text, false), "utf-16be"],
  ])("reads %s from its byte order mark", async (_name, bytes, encoding) => {
    const { workbook, values } = await read(bytes);
    expect([workbook.encoding, workbook.encodingSource]).toEqual([
      encoding,
      "byte-order-mark",
    ]);
    expect(values).toEqual(expected);
  });

  it("falls back to Windows-1252, with a warning, when the bytes are not UTF-8", async () => {
    const { workbook, values } = await read(
      windows1252("item;price\r\nCafé;€3,50\r\nIt’s;1\r\n"),
    );
    expect([workbook.encoding, workbook.encodingSource]).toEqual([
      "windows-1252",
      "fallback",
    ]);
    expect(values).toEqual([
      ["item", "price"],
      ["Café", "€3,50"],
      ["It’s", "1"],
    ]);
    expect(workbook.warnings).toEqual([
      "data.csv is not valid UTF-8 and has no byte order mark, so it was read as Windows-1252. If its accented letters look wrong, choose its encoding and run again.",
    ]);
  });

  it("takes a chosen encoding over the detected one", async () => {
    const { workbook, values } = await read(utf8("a\né\n"), {
      encoding: "windows-1252",
    });
    expect(workbook.encodingSource).toBe("chosen");
    expect(values).toEqual([["a"], ["Ã©"]]);
  });

  it("lets a byte order mark win over a chosen encoding, with a warning", async () => {
    const { workbook, values } = await read(
      Uint8Array.from([0xef, 0xbb, 0xbf, ...utf8(text)]),
      { encoding: "windows-1252" },
    );
    expect([workbook.encoding, workbook.encodingSource]).toEqual([
      "utf-8",
      "byte-order-mark",
    ]);
    expect(values).toEqual(expected);
    expect(workbook.warnings).toEqual([
      "data.csv starts with a UTF-8 byte order mark, so it was read as UTF-8 rather than the chosen Windows-1252.",
    ]);
  });

  it("refuses UTF-32, whose mark begins as UTF-16 LE's does", async () => {
    for (const mark of [
      [0xff, 0xfe, 0, 0],
      [0, 0, 0xfe, 0xff],
    ]) {
      const error = await failure(Uint8Array.from([...mark, 0x61, 0, 0, 0]));
      expect(error.code).toBe("XLSX_CSV_ENCODING_UNKNOWN");
      expect(error.message).toContain("UTF-32");
    }
  });

  it("refuses zero bytes without a byte order mark", async () => {
    const bytes = utf16(text, true).subarray(2);
    const error = await failure(bytes);
    expect(error.code).toBe("XLSX_CSV_ENCODING_UNKNOWN");
    expect(error.message).toContain("most likely UTF-16");
    const { values } = await read(bytes, { encoding: "utf-16le" });
    expect(values).toEqual(expected);
  });
});

describe("malformed CSV", () => {
  it("refuses a quoted field that is never closed, naming its row", async () => {
    const error = await failure(utf8('a,b\n1,2\n3,"open\n4,5\n'));
    expect(error.code).toBe("XLSX_CSV_MALFORMED");
    expect(error.message).toContain(
      "row 3 opens a quoted field that is never closed",
    );
    expect(error.details).toMatchObject({ row: 3 });
  });

  it("refuses text after a closing quote", async () => {
    const error = await failure(utf8('a,b\n1,"two"x\n3,4\n'));
    expect(error.code).toBe("XLSX_CSV_MALFORMED");
    expect(error.message).toContain(
      "row 2 has a quoted field with text after its closing quote",
    );
  });

  it("refuses an unclosed quote across pieces, naming its row", async () => {
    const error = await failure(
      utf8(`a\n"${"x".repeat(CSV_PIECE_BYTES * 2)}\n`),
    );
    expect(error.code).toBe("XLSX_CSV_MALFORMED");
    expect(error.message).toContain("row 2");
  });
});

describe("CSV typing", () => {
  it("keeps every field text by default", async () => {
    const { values } = await read(utf8("n,d\n1234.5,2025-01-31\n"));
    expect(values[1]).toEqual(["1234.5", "2025-01-31"]);
  });

  it("reads numbers only when they are only a number", async () => {
    const { values } = await read(
      utf8(
        'n\n42\n-1.5\n+7\n"1,234.50"\n0.25\n0\n-0\n007\n1e5\n" 3"\n$4\n1234567890123456\n123456789012345\n"1,23"\n.5\n',
      ),
      { numbers: true },
    );
    expect(values.slice(1).map((row) => row[0])).toEqual([
      42,
      -1.5,
      7,
      1234.5,
      0.25,
      0,
      0,
      "007",
      "1e5",
      " 3",
      "$4",
      "1234567890123456",
      123456789012345,
      "1,23",
      ".5",
    ]);
  });

  it("reads semicolon-locale numbers with their separators", async () => {
    const { values } = await read(utf8("n;m\n1.234,5;3,25\n12,5;1234,5\n"), {
      numbers: true,
      decimalSeparator: ",",
    });
    expect(values.slice(1)).toEqual([
      [1234.5, 3.25],
      [12.5, 1234.5],
    ]);
  });

  it("reads numbers with no thousands separator", () => {
    const reader = new CsvNumberReader(".", "");
    expect(reader.read("1234.5")).toBe(1234.5);
    expect(reader.read("1,234")).toBeUndefined();
  });

  it("reads ISO dates only when asked, and only whole real days", async () => {
    const { values } = await read(
      utf8(
        "d\n2025-01-31\n2025-01-31 13:45\n2025-01-31T13:45:30.5\n2025-02-30\n2025-01-31Z\n1899-12-31\n31/01/2025\n2025-1-31\n",
      ),
      { dates: "iso" },
    );
    expect(values.slice(1).map((row) => row[0])).toEqual([
      "2025-01-31T00:00:00.000Z",
      "2025-01-31T13:45:00.000Z",
      "2025-01-31T13:45:30.500Z",
      "2025-02-30",
      "2025-01-31Z",
      "1899-12-31",
      "31/01/2025",
      "2025-1-31",
    ]);
  });

  it("reads a named order and never guesses it", () => {
    expect(csvDate("03/04/2025", "dmy")).toBe("2025-04-03T00:00:00.000Z");
    expect(csvDate("03/04/2025", "mdy")).toBe("2025-03-04T00:00:00.000Z");
    expect(csvDate("03/04/2025", "iso")).toBeUndefined();
    expect(csvDate("3.4.2025 9:05", "dmy")).toBe("2025-04-03T09:05:00.000Z");
    expect(csvDate("13/31/2025", "dmy")).toBeUndefined();
    expect(csvDate("13/31/2025", "mdy")).toBeUndefined();
    expect(csvDate("31/12/2025", "mdy")).toBeUndefined();
    expect(csvDate("03/04-2025", "dmy")).toBeUndefined();
    expect(csvDate("03/04/25", "dmy")).toBeUndefined();
    expect(csvDate("2025/4/3", "ymd")).toBe("2025-04-03T00:00:00.000Z");
    expect(csvDate("2025-04-03", "dmy")).toBe("2025-04-03T00:00:00.000Z");
    expect(csvDate("2025/4/3 24:00", "ymd")).toBeUndefined();
    expect(csvDate("2025/4/3 1:00 PM", "ymd")).toBeUndefined();
  });

  it("gives each cell its field as text, and a date the field it came from", async () => {
    const { rows } = await read(
      utf8("1,234,2025-01-31\n"),
      { numbers: true, dates: "iso", delimiter: "," },
      { text: true, occupancy: true },
    );
    expect(rows).toEqual([
      [
        0,
        [
          { column: 0, value: 1, text: "1" },
          { column: 1, value: 234, text: "234" },
          {
            column: 2,
            value: "2025-01-31T00:00:00.000Z",
            text: "2025-01-31",
            stored: "2025-01-31",
          },
        ],
      ],
    ]);
  });
});

describe("CSV options", () => {
  it.each([
    [{ delimiter: ";;" }, "delimiter"],
    [{ delimiter: '"' }, "delimiter"],
    [{ encoding: "latin9" }, "encoding"],
    [{ dates: "dym" }, "dates"],
    [{ decimalSeparator: ";" }, "decimalSeparator"],
    [{ thousandsSeparator: "_" }, "thousandsSeparator"],
    [
      { numbers: true, decimalSeparator: ",", thousandsSeparator: "," },
      "thousandsSeparator",
    ],
    [{ decimalSeparator: "," }, "decimalSeparator"],
    [{ thousandsSeparator: "none" }, "thousandsSeparator"],
  ])("refuses %j before reading", (options, option) => {
    let caught: unknown;
    try {
      settleCsvOptions(options as CsvReadOptions);
    } catch (error) {
      caught = error;
    }
    expect(caught).toBeInstanceOf(ConsultChimpsError);
    expect((caught as ConsultChimpsError).code).toBe("XLSX_CSV_INVALID_OPTION");
    expect((caught as ConsultChimpsError).details).toEqual({ option });
  });

  it("fills in the separators", () => {
    expect(settleCsvOptions({ numbers: true }).numbers).toEqual({
      decimal: ".",
      thousands: ",",
    });
    expect(
      settleCsvOptions({ numbers: true, decimalSeparator: "," }).numbers,
    ).toEqual({ decimal: ",", thousands: "." });
    expect(
      settleCsvOptions({ numbers: true, thousandsSeparator: "" }).numbers,
    ).toEqual({ decimal: ".", thousands: "" });
    expect(settleCsvOptions(undefined).numbers).toBeUndefined();
  });
});

describe("a row that never ends", () => {
  it("is refused once it runs past 16 MB, before the file is read to its end", async () => {
    const open = `a\n"${"x".repeat(CSV_MAX_OPEN_ROW_CHARS + CSV_PIECE_BYTES)}`;
    const bytes = utf8(`${open}\n${"y,z\n".repeat(CSV_PIECE_BYTES)}`);
    let furthest = 0;
    const source = bytesSource("data.csv", bytes);
    const watched = {
      ...source,
      readAt: (offset: number, length: number) => {
        furthest = Math.max(furthest, offset + length);
        return source.readAt(offset, length);
      },
    };
    const workbook = await CsvWorkbook.open(watched, {
      file: "data.csv",
      source: "data.csv",
      details: {},
    });
    furthest = 0;
    let caught: unknown;
    try {
      await workbook.readWorksheet(workbook.sheets[0]!, {
        begin: () => undefined,
        row: () => undefined,
      });
    } catch (error) {
      caught = error;
    }
    expect((caught as ConsultChimpsError).code).toBe("XLSX_CSV_MALFORMED");
    expect((caught as ConsultChimpsError).message).toContain(
      "row 2 is longer than 16 MB, the most one row may hold here",
    );
    expect(furthest).toBeLessThan(bytes.length);
  });
});
