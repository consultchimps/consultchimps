/**
 * The two-pass consolidation (ADR 0006, #225) against the in-memory pipeline
 * it replaced: the same tables read by the table reader, folded by the same
 * mapping, stacked by the same union and written by the same writer must give
 * the same bytes. Then the changes made on purpose, and what a failed or
 * cancelled run leaves behind.
 */
import { writeFileSync } from "node:fs";
import { mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import type { RandomAccessSource } from "@consultchimps/core";
import {
  applyColumnMappingToTables,
  unionTables,
  type ColumnMapping,
} from "@consultchimps/tabular";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import {
  blobSource,
  consolidateWorkbookSources,
  consolidateWorkbooksBytes,
  readWorkbookTablesBytes,
  type ByteSink,
  type WorkbookInputBytes,
} from "../src/bytes.js";
import { consolidateWorkbooks } from "../src/index.js";
import { buildTableWorkbookBytes } from "../src/shared.js";
import {
  buildWorkbookFixture,
  type FixtureValue,
} from "./support/workbook-fixture.js";

type Rows = unknown[][];

/** A date as a declared date cell (`t="d"`), as SheetJS wrote one. */
function dateCell(date: Date): FixtureValue {
  return { date: date.toISOString() };
}

async function fixtureInput(
  name: string,
  sheets: Array<{
    name: string;
    rows: Rows;
    merges?: string[];
    hidden?: boolean;
  }>,
  options: { sharedStrings?: boolean } = {},
): Promise<WorkbookInputBytes> {
  return {
    name,
    bytes: await buildWorkbookFixture({
      // Inline strings unless asked, so both string stores stay covered.
      inlineStrings: options.sharedStrings !== true,
      sheets: sheets.map((sheet) => ({
        name: sheet.name,
        rows: sheet.rows.map((row) =>
          row.map((value) =>
            value instanceof Date ? dateCell(value) : (value as FixtureValue),
          ),
        ),
        ...(sheet.merges ? { merges: sheet.merges } : {}),
        ...(sheet.hidden ? { state: "hidden" as const } : {}),
      })),
    }),
  };
}

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** One worksheet written by hand, for markup no library writes. */
async function handInput(
  name: string,
  sheetXml: string,
): Promise<WorkbookInputBytes> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Hand" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<worksheet xmlns="${MAIN}">${sheetXml}</worksheet>`,
  );
  return { name, bytes: await zip.generateAsync({ type: "uint8array" }) };
}

interface RunOptions {
  addSourceColumns?: boolean;
  headerRow?: number;
  includeHiddenSheets?: boolean;
  mapping?: ColumnMapping;
  normalizeHeaders?: boolean;
  sheets?: string[];
}

/** The pipeline consolidation ran before the two passes, in memory. */
async function inMemoryConsolidation(
  inputs: WorkbookInputBytes[],
  options: RunOptions,
): Promise<Uint8Array> {
  const tables = [];
  for (const input of inputs) {
    tables.push(...(await readWorkbookTablesBytes(input, options)));
  }
  const mapped = options.mapping
    ? applyColumnMappingToTables(tables, options.mapping).tables
    : tables;
  return buildTableWorkbookBytes(
    unionTables(mapped, {
      addSourceColumns: options.addSourceColumns,
      normalizeHeaders: options.normalizeHeaders,
    }),
    "Consolidated",
  );
}

const reviewInputs = (): Promise<WorkbookInputBytes[]> =>
  Promise.all([
    fixtureInput("north.xlsx", [
      {
        name: "Review Log",
        rows: [
          ["Quarterly review"],
          [],
          ["Case ID", "Failed Checks", null, "Opened", "Flag", "Note"],
          ["R-1", 5, null, new Date(Date.UTC(2024, 0, 31)), true, "a & b"],
          ["R-2", 1e21, null, null, false, "  padded  "],
          [
            null,
            -0.000001,
            null,
            new Date(Date.UTC(1999, 11, 31, 12)),
            null,
            "",
          ],
        ],
      },
      { name: "Empty", rows: [] },
      {
        name: "Hidden",
        rows: [
          ["Case ID", "Secret"],
          ["H-1", 1],
        ],
        hidden: true,
      },
    ]),
    fixtureInput(
      "south.xlsx",
      [
        {
          name: "vF",
          rows: [
            ["Report title", null, null, null],
            ["case_id", "Failed_Checks", null, "Region"],
            ["S-1", 7, "unnamed", "South"],
            ["S-2", 8, null, "South"],
          ],
          merges: ["A1:D1"],
        },
        {
          name: "Dupes",
          rows: [
            ["A", "A", "A_2", null],
            [1, 2, 3, 4],
          ],
        },
      ],
      { sharedStrings: true },
    ),
  ]);

describe("two-pass consolidation gives the in-memory pipeline's bytes", () => {
  it.each<[string, RunOptions]>([
    ["by default", {}],
    [
      "with normalized headers and hidden sheets",
      { normalizeHeaders: true, includeHiddenSheets: true },
    ],
    [
      "without source columns, from selected sheets",
      { addSourceColumns: false, sheets: ["review log", "VF"] },
    ],
    ["from a declared header row", { headerRow: 2 }],
  ])("%s", async (_, options) => {
    const inputs = await reviewInputs();
    const { outputs } = await consolidateWorkbooksBytes({ inputs, ...options });
    expect(Buffer.from(outputs[0]!.bytes)).toEqual(
      Buffer.from(await inMemoryConsolidation(inputs, options)),
    );
  });

  it("with a mapping that folds, coerces and adds constants", async () => {
    const inputs = await Promise.all([
      fixtureInput("north.xlsx", [
        {
          name: "Cases",
          rows: [
            ["Case ID", "Total", "Opened", "Region"],
            ["R-1", "1.234,50", "03/01/2024", "north"],
            ["R-2", "12.345.678,25", "3/1/2024", "north"],
          ],
        },
      ]),
      fixtureInput("south.xlsx", [
        {
          name: "Cases",
          rows: [
            ["reference", "total", "Opened", "Area"],
            ["S-1", "7,25", "", "south"],
          ],
        },
      ]),
    ]);
    const mapping: ColumnMapping = {
      version: 1,
      columns: [
        { name: "Case_ID", aliases: ["Reference"] },
        {
          name: "Amount",
          aliases: ["Total"],
          coercion: {
            type: "number",
            decimalSeparator: ",",
            thousandsSeparator: ".",
          },
        },
        {
          name: "Opened On",
          aliases: ["Opened"],
          coercion: { type: "date", format: "D/M/YYYY" },
        },
      ],
      constants: { Programme: "Review 2024", Wave: 3 },
    };
    const { outputs, result } = await consolidateWorkbooksBytes({
      inputs,
      mapping,
    });
    expect(result.metrics.unmappedColumns).toBe(2);
    expect(Buffer.from(outputs[0]!.bytes)).toEqual(
      Buffer.from(await inMemoryConsolidation(inputs, { mapping })),
    );
  });

  it("with rows out of order, cells without references, and no dimension", async () => {
    const inputs = [
      await handInput(
        "hand.xlsx",
        `<sheetData><row r="3"><c r="A3" t="inlineStr"><is><t>R-2</t></is></c><c r="B3"><v>2</v></c></row><row r="1"><c t="inlineStr"><is><t>Case</t></is></c><c t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>R-1</t></is></c><c r="B2"><v>1</v></c></row><row r="7"/></sheetData>`,
      ),
    ];
    const { outputs, result } = await consolidateWorkbooksBytes({ inputs });
    expect(result.metrics.outputRows).toBe(2);
    expect(Buffer.from(outputs[0]!.bytes)).toEqual(
      Buffer.from(await inMemoryConsolidation(inputs, {})),
    );
  });
});

describe("two-pass consolidation: deliberate changes", () => {
  it("writes an error cell as an error and keeps a carriage return", async () => {
    const inputs = [
      await handInput(
        "hand.xlsx",
        `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Case</t></is></c><c r="B1" t="inlineStr"><is><t>Result</t></is></c><c r="C1" t="inlineStr"><is><t>Note</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>R-1</t></is></c><c r="B2" t="e"><v>#DIV/0!</v></c><c r="C2" t="inlineStr"><is><t>one_x000D_
two</t></is></c></row></sheetData>`,
      ),
    ];
    const { outputs } = await consolidateWorkbooksBytes({
      inputs,
      addSourceColumns: false,
    });
    const sheet = (await JSZip.loadAsync(outputs[0]!.bytes)).file(
      "xl/worksheets/sheet1.xml",
    );
    const xml = await sheet!.async("string");
    expect(xml).toContain(`<c r="B2" t="e"><v>#DIV/0!</v></c>`);
    expect(xml).toContain("one_x000D_\ntwo");
  });
});

describe("two-pass consolidation: what a failed run leaves behind", () => {
  it("writes nothing, not even a staging file, when an input changes between the passes", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));
    try {
      const [first, second] = await reviewInputs();
      const firstPath = path.join(directory, first!.name);
      const secondPath = path.join(directory, second!.name);
      await writeFile(firstPath, first!.bytes);
      await writeFile(secondPath, second!.bytes);
      const output = path.join(directory, "out", "consolidated.xlsx");

      await expect(
        consolidateWorkbooks({
          inputs: [firstPath, secondPath],
          output,
          onProgress: (progress) => {
            if (
              progress.stage === "reading-workbooks" &&
              progress.completed === 2
            ) {
              writeFileSync(firstPath, second!.bytes);
            }
          },
        }),
      ).rejects.toMatchObject({ code: "XLSX_READ_FAILED" });
      expect(await readdir(path.dirname(output))).toEqual([]);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("refuses an input changed between the passes even when its size is not", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));
    try {
      const sheet = (amount: string) =>
        `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2"><v>${amount}</v></c></row></sheetData>`;
      const before = await handInput("hand.xlsx", sheet("1111"));
      const after = await handInput("hand.xlsx", sheet("2222"));
      expect(after.bytes.length).toBe(before.bytes.length);
      const input = path.join(directory, "hand.xlsx");
      await writeFile(input, before.bytes);
      const output = path.join(directory, "consolidated.xlsx");
      await expect(
        consolidateWorkbooks({
          inputs: [input],
          output,
          onProgress: (progress) => {
            if (progress.stage === "reading-workbooks") {
              writeFileSync(input, after.bytes);
            }
          },
        }),
      ).rejects.toMatchObject({
        code: "XLSX_READ_FAILED",
        message: expect.stringContaining(
          "changed while it was being consolidated",
        ),
      });
      expect(await readdir(directory)).toEqual(["hand.xlsx"]);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("writes nothing when cancelled after the inputs were read", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));
    try {
      const inputs = await reviewInputs();
      const paths = [];
      for (const input of inputs) {
        const target = path.join(directory, input.name);
        await writeFile(target, input.bytes);
        paths.push(target);
      }
      const output = path.join(directory, "out", "consolidated.xlsx");
      const controller = new AbortController();
      await expect(
        consolidateWorkbooks({
          inputs: paths,
          output,
          signal: controller.signal,
          onProgress: (progress) => {
            if (progress.completed === inputs.length) controller.abort();
          },
        }),
      ).rejects.toMatchObject({ code: "OPERATION_ABORTED" });
      await expect(readFile(output)).rejects.toMatchObject({ code: "ENOENT" });
      expect(await readdir(path.dirname(output)).catch(() => [])).toEqual([]);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });
});

/** A source that fails a read larger than the zip reader's chunk. */
function guardedSource(input: WorkbookInputBytes): {
  source: RandomAccessSource;
  largestRead: () => number;
} {
  const inner = blobSource(input.name, new Blob([input.bytes]));
  let largest = 0;
  return {
    largestRead: () => largest,
    source: {
      name: inner.name,
      size: inner.size,
      readAt: (offset, length, signal) => {
        largest = Math.max(largest, length);
        if (length > 1024 * 1024) {
          throw new Error(`Read ${length} bytes of ${inner.size} at once`);
        }
        return inner.readAt(offset, length, signal);
      },
    },
  };
}

/** A sink that keeps each chunk and counts the flushes between them. */
function collectingSink(): {
  sink: ByteSink;
  bytes: () => Buffer;
  chunks: () => number;
} {
  const chunks: Uint8Array[] = [];
  return {
    sink: {
      write: (chunk) => {
        chunks.push(chunk);
      },
      flush: () => Promise.resolve(),
    },
    bytes: () => Buffer.concat(chunks),
    chunks: () => chunks.length,
  };
}

/** A worksheet of `rows` rows stored uncompressed, so the file is large. */
function largeSheetXml(rows: number): string {
  const cells = (row: number, values: string[]) =>
    values
      .map(
        (value, column) =>
          `<c r="${String.fromCharCode(65 + column)}${row}" t="inlineStr"><is><t>${value}</t></is></c>`,
      )
      .join("");
  const body = [`<row r="1">${cells(1, ["Case_ID", "Region", "Note"])}</row>`];
  for (let row = 2; row <= rows; row += 1) {
    body.push(
      `<row r="${row}">${cells(row, [`R-${row}`, row % 2 ? "North" : "South", `Note for row ${row}`])}</row>`,
    );
  }
  return `<sheetData>${body.join("")}</sheetData>`;
}

describe("consolidation over sources and a sink", () => {
  it("reads in pieces, writes in chunks, and gives the command line's bytes", async () => {
    const large = await handInput("large.xlsx", largeSheetXml(20_000));
    expect(large.bytes.length).toBeGreaterThan(2 * 1024 * 1024);
    const inputs = [...(await reviewInputs()), large];
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));
    try {
      const paths = [];
      for (const input of inputs) {
        const target = path.join(directory, input.name);
        await writeFile(target, input.bytes);
        paths.push(target);
      }
      const output = path.join(directory, "out.xlsx");
      const reference = await consolidateWorkbooks({ inputs: paths, output });

      const guarded = inputs.map(guardedSource);
      const collected = collectingSink();
      const outcome = await consolidateWorkbookSources({
        inputs: guarded.map((entry) => entry.source),
        output: collected.sink,
        outputName: "out",
      });

      expect(outcome.outputName).toBe("out.xlsx");
      expect(outcome.mappingDraft).toBeUndefined();
      expect(outcome.result.metrics).toEqual(reference.metrics);
      expect(outcome.result.metrics.outputRows).toBeGreaterThan(20_000);
      expect(collected.chunks()).toBeGreaterThan(1);
      expect(Math.max(...guarded.map((entry) => entry.largestRead()))).toBe(
        1024 * 1024,
      );
      expect(collected.bytes()).toEqual(await readFile(output));
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("returns the mapping draft beside the streamed workbook", async () => {
    const inputs = await reviewInputs();
    const collected = collectingSink();
    const outcome = await consolidateWorkbookSources({
      inputs: inputs.map((input) => guardedSource(input).source),
      output: collected.sink,
      suggestMapping: true,
    });
    const reference = await consolidateWorkbooksBytes({
      inputs,
      suggestMapping: true,
    });
    expect(collected.bytes()).toEqual(Buffer.from(reference.outputs[0]!.bytes));
    expect(outcome.mappingDraft).toEqual(reference.outputs[1]);
    expect(outcome.result).toEqual(reference.result);
  });
});
