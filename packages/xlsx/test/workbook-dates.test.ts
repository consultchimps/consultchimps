/**
 * The invariant behind #244: a value read from a workbook as a date is written
 * back as a date serial with a date format, by every writer that writes cells.
 * Each writer below runs on the same inputs, and every date the inputs hold
 * must read back from its output as the same date, stored as a number rather
 * than as text. A new writer joins the list.
 */
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import {
  consolidateWorkbooksBytes,
  mergeWorkbooksBytes,
  readWorkbookTablesBytes,
  splitWorkbookBytes,
} from "../src/bytes.js";
import { consolidateWorkbooks, writeTable } from "../src/index.js";
import {
  calendarIsoParts,
  serial1900,
  serialCalendarParts,
} from "../src/model/calendar.js";
import { columnLetters } from "../src/package/table-writer.js";
import { WORKBOOK_DATE_TEXT } from "../src/shared.js";
import { sheetNamesOf, sheetRows, sheetXml } from "./support/read-workbook.js";
import { buildWorkbookFixture } from "./support/workbook-fixture.js";

interface Input {
  readonly name: string;
  readonly bytes: Uint8Array;
}

const DATES_1900 = [
  "2024-01-16T00:00:00.000Z",
  "2024-01-16T12:00:00.000Z",
  "2024-02-29T08:30:15.250Z",
  "1900-01-15T00:00:00.000Z",
  "1900-03-01T00:00:00.000Z",
];
const DATE_1904 = "2024-01-17T00:00:00.000Z";
/** Before the 1900 system's first day: no serial, so it stays text. */
const BEFORE_1900 = "1850-06-01T00:00:00.000Z";

async function inputs(): Promise<{ modern: Input; old: Input }> {
  const modern = await buildWorkbookFixture({
    sheets: [
      {
        name: "Data",
        rows: [
          ["Case", "Group", "Opened"],
          ["A", "x", { value: 45307, format: "dd/mm/yyyy" }],
          ["B", "x", { value: 45307.5, format: 22 }],
          ["C", "x", { date: "2024-02-29T08:30:15.250" }],
          ["D", "x", { value: 15, format: 14 }],
          ["E", "x", { value: 61, format: "d mmm yyyy" }],
          ["F", "x", { date: "1850-06-01" }],
          ["G", "x", "not a date"],
        ],
      },
    ],
  });
  // 1904 system: serial 43846 is 17 January 2024.
  const old = await buildWorkbookFixture({
    date1904: true,
    sheets: [
      {
        name: "Old",
        rows: [
          ["Case", "Group", "Opened"],
          ["H", "x", { value: 43846, format: 14 }],
        ],
      },
    ],
  });
  return {
    modern: { name: "modern.xlsx", bytes: modern },
    old: { name: "old.xlsx", bytes: old },
  };
}

type Writer = (modern: Input, old: Input) => Promise<Uint8Array[]>;

async function inDirectory<T>(
  run: (directory: string) => Promise<T>,
): Promise<T> {
  const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-dates-"));
  try {
    return await run(directory);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
}

const WRITERS: Record<string, Writer> = {
  "consolidate (bytes)": async (modern, old) =>
    (await consolidateWorkbooksBytes({ inputs: [modern, old] })).outputs.map(
      (output) => output.bytes,
    ),
  "consolidate (file)": (modern, old) =>
    inDirectory(async (directory) => {
      const paths: string[] = [];
      for (const input of [modern, old]) {
        const file = path.join(directory, input.name);
        await writeFile(file, input.bytes);
        paths.push(file);
      }
      const output = path.join(directory, "out.xlsx");
      await consolidateWorkbooks({ inputs: paths, output });
      return [new Uint8Array(await readFile(output))];
    }),
  "compact split": (modern, old) =>
    Promise.all(
      [modern, old].map(async (input) => {
        const { outputs } = await splitWorkbookBytes({
          input,
          column: "Group",
          sheet: input === modern ? "Data" : "Old",
          preserveWorkbook: false,
        });
        expect(outputs).toHaveLength(1);
        return outputs[0]!.bytes;
      }),
    ),
  "write table": (modern, old) =>
    inDirectory(async (directory) => {
      const outputs: Uint8Array[] = [];
      for (const input of [modern, old]) {
        const [table] = await readWorkbookTablesBytes(input);
        const output = path.join(directory, `table-${input.name}`);
        await writeTable(output, table!);
        outputs.push(new Uint8Array(await readFile(output)));
      }
      return outputs;
    }),
  // Merge copies cells and styles as they are. Inputs in different date
  // systems are a separate defect, so each system is merged on its own here.
  merge: (modern, old) =>
    Promise.all(
      [modern, old].map(async (input) => {
        const { outputs } = await mergeWorkbooksBytes({
          inputs: [input, input],
        });
        return outputs[0]!.bytes;
      }),
    ),
};

/** Every date a workbook reads as, with whether its cell is stored as text. */
async function datesOf(
  bytes: Uint8Array,
): Promise<Array<{ value: string; text: boolean }>> {
  const found: Array<{ value: string; text: boolean }> = [];
  for (const sheet of await sheetNamesOf(bytes)) {
    const rows = await sheetRows(bytes, sheet);
    const xml = await sheetXml(bytes, sheet);
    rows.forEach((row, rowIndex) => {
      row.forEach((value, columnIndex) => {
        if (typeof value !== "string" || !WORKBOOK_DATE_TEXT.test(value)) {
          return;
        }
        const ref = `${columnLetters(columnIndex)}${rowIndex + 1}`;
        const cell = new RegExp(`<c r="${ref}"([^>]*)>`, "u").exec(xml);
        expect(cell, ref).not.toBeNull();
        found.push({
          value,
          text: /\bt="(?:s|str|inlineStr)"/u.test(cell![1]!),
        });
      });
    });
  }
  return found;
}

describe("a date read from a workbook is written back as a date", () => {
  it.each(Object.entries(WRITERS))("%s", async (_, writer) => {
    const { modern, old } = await inputs();
    const outputs = await writer(modern, old);
    const expected = [...DATES_1900, BEFORE_1900, DATE_1904].sort();

    const read: Array<{ value: string; text: boolean }> = [];
    for (const output of outputs) read.push(...(await datesOf(output)));
    const values = [...new Set(read.map((entry) => entry.value))].sort();
    expect(values).toEqual(expected);
    for (const entry of read) {
      // A day before 1900 has no serial in the 1900 system, so a writer that
      // writes serials keeps it as text; merge keeps its declared date cell.
      if (entry.value !== BEFORE_1900) {
        expect(entry.text, entry.value).toBe(false);
      }
    }
  });

  it("writes the 1900 serial with a date or a date-time format", async () => {
    const { modern, old } = await inputs();
    const { outputs } = await consolidateWorkbooksBytes({
      inputs: [modern, old],
      addSourceColumns: false,
    });
    const zip = await JSZip.loadAsync(outputs[0]!.bytes);
    const sheet = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    const styles = await zip.file("xl/styles.xml")!.async("string");
    const workbook = await zip.file("xl/workbook.xml")!.async("string");

    expect(workbook).not.toContain("date1904");
    expect(styles).toContain(
      `<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm:ss"/>`,
    );
    expect(styles).toContain(
      `<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>`,
    );
    expect(sheet).toContain(`<c r="C2" s="1"><v>45307</v></c>`);
    expect(sheet).toContain(`<c r="C3" s="2"><v>45307.5</v></c>`);
    expect(sheet).toContain(`<c r="C5" s="1"><v>15</v></c>`);
    expect(sheet).toContain(`<c r="C6" s="1"><v>61</v></c>`);
    expect(sheet).toContain(
      `<c r="C7" t="inlineStr"><is><t>${BEFORE_1900}</t></is></c>`,
    );
    expect(sheet).toContain(
      `<c r="C8" t="inlineStr"><is><t>not a date</t></is></c>`,
    );
    // The 1904 input's 17 January 2024, recounted in the 1900 system.
    expect(sheet).toContain(`<c r="C9" s="1"><v>45308</v></c>`);
  });
});

describe("text in the workbook date spelling", () => {
  it("is written as a date, the accepted conflation", async () => {
    const bytes = await buildWorkbookFixture({
      sheets: [
        {
          name: "Data",
          rows: [["Opened"], ["2024-01-16T00:00:00.000Z"]],
        },
      ],
    });
    const { outputs } = await consolidateWorkbooksBytes({
      inputs: [{ name: "text.xlsx", bytes }],
      addSourceColumns: false,
    });
    const zip = await JSZip.loadAsync(outputs[0]!.bytes);
    const sheet = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    expect(sheet).toContain(`<c r="A2" s="1"><v>45307</v></c>`);
  });
});

describe("the 1900 serial of a calendar date", () => {
  it("skips the invented 29 February 1900 and stops before 1900", () => {
    const serial = (text: string): number | undefined =>
      serial1900(calendarIsoParts(text)!);
    expect(serial("1899-12-31T00:00:00.000Z")).toBeUndefined();
    expect(serial("1900-01-01T00:00:00.000Z")).toBe(1);
    expect(serial("1900-02-28T00:00:00.000Z")).toBe(59);
    expect(serial("1900-03-01T00:00:00.000Z")).toBe(61);
    expect(serial("9999-12-31T00:00:00.000Z")).toBe(2_958_465);
    expect(calendarIsoParts("2023-02-29T00:00:00.000Z")).toBeUndefined();
    expect(calendarIsoParts("2024-01-16")).toBeUndefined();
    const last = serial("9999-12-31T23:59:59.999Z")!;
    expect(last).toBeLessThan(2_958_466);
    expect(serialCalendarParts(last, false)).toEqual(
      calendarIsoParts("9999-12-31T23:59:59.999Z"),
    );
  });

  it("is the inverse of reading a serial, to the millisecond", () => {
    for (let serial = 1; serial < 2_958_466; serial += 7919.123_456) {
      const parts = serialCalendarParts(serial, false)!;
      const back = serialCalendarParts(serial1900(parts)!, false);
      expect(back, String(serial)).toEqual(parts);
    }
  });
});
