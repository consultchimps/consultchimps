/**
 * The invariant from #162: a formula cell with no cached value is never
 * silently dropped. Every operation that reads cell values counts it in
 * `formulaCellsWithoutCachedValues` and names it in a warning, so the reader
 * knows to open and save the workbook in Excel. A new value-reading operation
 * belongs in this table.
 */
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";
import { afterEach, describe, expect, it } from "vitest";

import {
  consolidateWorkbooksBytes,
  describeWorkbookBytes,
  mergeWorkbooksBytes,
  readWorksheetRecordsBytes,
  splitWorkbookBytes,
} from "../src/bytes.js";
import {
  consolidateWorkbooks,
  describeWorkbook,
  mergeWorkbooks,
  readWorksheetRecords,
  splitWorkbookByColumn,
} from "../src/index.js";
import {
  buildCorpusWorkbook,
  cleanupCorpusDirectories,
  CORPUS_PARTS,
  CORPUS_SHEET,
  CORPUS_SPLIT_COLUMN,
  CORPUS_TABLE_NAME,
  createCorpusDirectory,
} from "./corpus/fixtures.js";
import { buildWorkbookFixture } from "./support/workbook-fixture.js";

afterEach(cleanupCorpusDirectories);

/**
 * The records reader refuses the corpus's unnamed side column, so it reads a
 * plain sheet with the same cell left uncalculated.
 */
async function recordsWorkbook(): Promise<Uint8Array> {
  return buildWorkbookFixture({
    sheets: [
      {
        name: CORPUS_SHEET,
        rows: [
          ["Report"],
          [],
          ["Record", "Client", "Group"],
          [1, "Client A", "Alpha"],
          [2, "Client B", { formula: '"Beta"' }],
        ],
      },
    ],
  });
}

/** The split-column cell of the second data row, which every read covers. */
const UNCACHED = `${CORPUS_SHEET}!C5`;

/** The corpus workbook with Data!C5 turned into a formula never calculated. */
async function uncachedWorkbook(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(
    // Structured references, which a preserved table split accepts.
    await buildCorpusWorkbook({ formulas: "structured", shape: "table" }),
  );
  const sheet = await zip.file(CORPUS_PARTS.dataSheet)!.async("string");
  const cell = /<c r="C5"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/u;
  expect(sheet).toMatch(cell);
  zip.file(
    CORPUS_PARTS.dataSheet,
    sheet.replace(cell, '<c r="C5"><f>"Beta"</f></c>'),
  );
  return zip.generateAsync({ type: "uint8array" });
}

interface Reported {
  count: number;
  warnings: readonly string[];
}

type Operation = (
  input: string,
  bytes: Uint8Array,
  out: string,
) => Promise<Reported>;

const reported = (result: {
  metrics: { formulaCellsWithoutCachedValues: number };
  warnings: string[];
}): Reported => ({
  count: result.metrics.formulaCellsWithoutCachedValues,
  warnings: result.warnings,
});

const named = { name: "input.xlsx" };

const OPERATIONS: Record<string, Operation> = {
  consolidate: async (input, _bytes, out) =>
    reported(
      await consolidateWorkbooks({
        inputs: [input],
        output: path.join(out, "all.xlsx"),
      }),
    ),
  "consolidate (bytes)": async (_input, bytes) =>
    reported(
      (await consolidateWorkbooksBytes({ inputs: [{ ...named, bytes }] }))
        .result,
    ),
  "split, every worksheet": async (input, _bytes, out) =>
    reported(
      await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: out,
      }),
    ),
  "split, every worksheet, values only": async (input, _bytes, out) =>
    reported(
      await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: out,
        values: true,
      }),
    ),
  "split, compact worksheet": async (input, _bytes, out) =>
    reported(
      await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        headerRow: 3,
        input,
        outputDirectory: out,
        preserveWorkbook: false,
        sheet: CORPUS_SHEET,
      }),
    ),
  "split, compact worksheet (bytes)": async (_input, bytes) =>
    reported(
      (
        await splitWorkbookBytes({
          column: CORPUS_SPLIT_COLUMN,
          headerRow: 3,
          input: { ...named, bytes },
          preserveWorkbook: false,
          sheet: CORPUS_SHEET,
        })
      ).result,
    ),
  "split, preserved Excel Table": async (input, _bytes, out) =>
    reported(
      await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: out,
        preserveWorkbook: true,
        table: CORPUS_TABLE_NAME,
      }),
    ),
  "split, preserved Excel Table, values only": async (input, _bytes, out) =>
    reported(
      await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: out,
        preserveWorkbook: true,
        table: CORPUS_TABLE_NAME,
        values: true,
      }),
    ),
  "merge, values only": async (input, _bytes, out) =>
    reported(
      await mergeWorkbooks([input], path.join(out, "merged.xlsx"), {
        values: true,
      }),
    ),
  "merge, values only (bytes)": async (_input, bytes) =>
    reported(
      (
        await mergeWorkbooksBytes({
          inputs: [{ ...named, bytes }],
          values: true,
        })
      ).result,
    ),
  describe: async (input) => reported((await describeWorkbook(input)).result),
  "describe (bytes)": async (_input, bytes) =>
    reported((await describeWorkbookBytes({ ...named, bytes })).result),
  // The records carry the cells; the populate operation reports them.
  records: async (_input, _bytes, out) => {
    const input = path.join(out, "records.xlsx");
    await mkdir(out, { recursive: true });
    await writeFile(input, await recordsWorkbook());
    const records = await readWorksheetRecords(input, {
      headerRow: 3,
      worksheet: CORPUS_SHEET,
    });
    return {
      count: records.uncachedFormulas.length,
      warnings: records.uncachedFormulas,
    };
  },
  "records (bytes)": async () => {
    const records = await readWorksheetRecordsBytes(
      { ...named, bytes: await recordsWorkbook() },
      { headerRow: 3, worksheet: CORPUS_SHEET },
    );
    return {
      count: records.uncachedFormulas.length,
      warnings: records.uncachedFormulas,
    };
  },
};

describe("formula cells with no cached value", () => {
  it.each(Object.keys(OPERATIONS))(
    "are counted and named by %s",
    async (name) => {
      const directory = await createCorpusDirectory();
      const bytes = await uncachedWorkbook();
      const input = path.join(directory, "input.xlsx");
      await writeFile(input, bytes);

      const result = await OPERATIONS[name]!(
        input,
        bytes,
        path.join(directory, "out"),
      );

      if (name === "split, every worksheet, values only") {
        // Like its other per-output counts, this one counts each output that
        // loses the cell, and the stale aggregates the split blanks.
        expect(result.count).toBeGreaterThanOrEqual(1);
      } else {
        expect(result.count).toBe(1);
      }
      expect(
        result.warnings.some((warning) => warning.includes(UNCACHED)),
      ).toBe(true);
    },
  );
});
