/**
 * Conformance corpus: merge and consolidate.
 *
 * The merge is a worksheet-object copy through SheetJS, so everything outside
 * that object model is silently lost. Phase 1b rebuilds it as a part-level
 * transplant; until then the losses are pinned here AS CURRENT BEHAVIOUR so
 * that the rebuild has to change them deliberately.
 */
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { afterEach, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import { mergeWorkbooksBytes } from "../../src/bytes.js";
import {
  consolidateWorkbooks,
  mergeWorkbooks,
  readWorkbookTables,
} from "../../src/index.js";
import {
  buildCorpusWorkbook,
  cleanupCorpusDirectories,
  conditionalFormattingSqref,
  CORPUS_COMMENT_TEXT,
  CORPUS_LOCAL_NAME,
  CORPUS_PARTS,
  CORPUS_RANGE_NAME,
  CORPUS_TABLE_NAME,
  createCorpusDirectory,
  dataValidationSqref,
  hasPackagePart,
  hyperlinkReferences,
  mergedCellReferences,
  packagePartNames,
  readPackagePart,
  readWorkbookBytes,
  worksheetCellFormula,
  worksheetCellValue,
  writeCorpusWorkbook,
} from "./fixtures.js";

const MERGED_DATA_PART = "xl/worksheets/sheet1.xml";
const MERGED_INDEX_PART = "xl/worksheets/sheet9.xml";

afterEach(cleanupCorpusDirectories);

async function mergedCorpus(): Promise<{
  bytes: Uint8Array;
  warnings: string[];
}> {
  const directory = await createCorpusDirectory();
  const first = await writeCorpusWorkbook(directory, "first.xlsx", {
    shape: "table",
    pivot: true,
  });
  const second = await writeCorpusWorkbook(directory, "second.xlsx", {
    shape: "range",
  });
  const output = path.join(directory, "merged.xlsx");
  const result = await mergeWorkbooks([first, second], output);
  return { bytes: await readWorkbookBytes(output), warnings: result.warnings };
}

describe("corpus: merge", () => {
  it("invariant: merge preserves worksheet visibility and records it in the sheet index", async () => {
    const directory = await createCorpusDirectory();
    const first = await writeCorpusWorkbook(directory, "first.xlsx", {
      shape: "table",
    });
    const second = await writeCorpusWorkbook(directory, "second.xlsx", {
      shape: "range",
    });
    const output = path.join(directory, "merged.xlsx");

    const result = await mergeWorkbooks([first, second], output);

    expect(result.metrics).toEqual({
      hiddenSheets: 4,
      inputFiles: 2,
      outputSheets: 8,
    });
    expect(result.warnings.join("\n")).toMatch(
      /see the visible "Sheet Index"/u,
    );
    const workbook = XLSX.read(await readFile(output), { type: "buffer" });
    expect(workbook.SheetNames).toEqual([
      "Data",
      "Summary",
      "Hidden",
      "VeryHidden",
      "Data (2)",
      "Summary (2)",
      "Hidden (2)",
      "VeryHidden (2)",
      "Sheet Index",
    ]);
    expect(workbook.Workbook?.Sheets?.map((sheet) => sheet.Hidden)).toEqual([
      0, 0, 1, 2, 0, 0, 1, 2, 0,
    ]);
    const index = await readPackagePart(
      await readWorkbookBytes(output),
      MERGED_INDEX_PART,
    );
    expect(index).toContain("Very hidden");
    expect(index).toContain("VeryHidden (2)");
  });

  it("invariant: merge reserves the sheet-index name so a source sheet cannot take it", async () => {
    const directory = await createCorpusDirectory();
    const collidingWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      collidingWorkbook,
      XLSX.utils.aoa_to_sheet([["Value"], [1]]),
      "Sheet Index",
    );
    const colliding = path.join(directory, "colliding.xlsx");
    await writeFile(
      colliding,
      XLSX.write(collidingWorkbook, { bookType: "xlsx", type: "buffer" }),
    );
    const output = path.join(directory, "merged.xlsx");

    await mergeWorkbooks([colliding], output);

    const workbook = XLSX.read(await readFile(output), { type: "buffer" });
    expect(workbook.SheetNames).toEqual(["Sheet Index (2)", "Sheet Index"]);
  });

  it("invariant: merge keeps merged ranges, hyperlinks, comments and formulas", async () => {
    const merged = await mergedCorpus();
    const dataXml = await readPackagePart(merged.bytes, MERGED_DATA_PART);

    expect(mergedCellReferences(dataXml)).toEqual(["A1:F1", "H6:I6"]);
    expect(hyperlinkReferences(dataXml)).toEqual(["A6"]);
    expect(worksheetCellFormula(dataXml, "E4")).toBe("D4*2");
    expect(worksheetCellValue(dataXml, "E4")).toBe("20");
    expect(await hasPackagePart(merged.bytes, CORPUS_PARTS.comments)).toBe(
      true,
    );
    expect(
      await readPackagePart(merged.bytes, CORPUS_PARTS.comments),
    ).toContain(CORPUS_COMMENT_TEXT);
  });

  it("pins: merge currently drops workbook-level and sheet-level defined names", async () => {
    const merged = await mergedCorpus();
    const workbookXml = await readPackagePart(
      merged.bytes,
      CORPUS_PARTS.workbook,
    );

    expect(workbookXml).not.toContain("definedName");
    expect(workbookXml).not.toContain(CORPUS_RANGE_NAME);
    expect(workbookXml).not.toContain(CORPUS_LOCAL_NAME);
  });

  it("pins: merge currently drops Excel Table parts", async () => {
    const merged = await mergedCorpus();
    const parts = await packagePartNames(merged.bytes);

    expect(parts.some((part) => part.startsWith("xl/tables/"))).toBe(false);
    const dataXml = await readPackagePart(merged.bytes, MERGED_DATA_PART);
    expect(dataXml).not.toContain("tableParts");
    expect(dataXml).not.toContain(CORPUS_TABLE_NAME);
  });

  it("pins: merge currently drops pivot tables and their caches", async () => {
    const merged = await mergedCorpus();
    const parts = await packagePartNames(merged.bytes);

    expect(parts.some((part) => part.startsWith("xl/pivotTables/"))).toBe(
      false,
    );
    expect(parts.some((part) => part.startsWith("xl/pivotCache/"))).toBe(false);
    expect(
      await readPackagePart(merged.bytes, CORPUS_PARTS.workbook),
    ).not.toContain("pivotCache");
  });

  it("pins: merge currently drops conditional formatting and data validation", async () => {
    const merged = await mergedCorpus();
    const dataXml = await readPackagePart(merged.bytes, MERGED_DATA_PART);

    expect(conditionalFormattingSqref(dataXml)).toBeUndefined();
    expect(dataValidationSqref(dataXml)).toBeUndefined();
  });

  it("pins: merge currently drops the calculation chain and the shared string table", async () => {
    const merged = await mergedCorpus();
    const parts = await packagePartNames(merged.bytes);

    expect(parts).not.toContain(CORPUS_PARTS.calcChain);
    expect(parts).not.toContain(CORPUS_PARTS.sharedStrings);
    // Strings are re-emitted inline on every cell instead of being interned.
    const dataXml = await readPackagePart(merged.bytes, MERGED_DATA_PART);
    expect(dataXml).toContain('<c r="B4" t="str"><v>Client A</v></c>');
  });

  it("pins: a values-only merge replaces formulas with their cached values", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "first.xlsx", {
      shape: "range",
    });
    const output = path.join(directory, "merged.xlsx");

    await mergeWorkbooks([input], output, { values: true });

    const dataXml = await readPackagePart(
      await readWorkbookBytes(output),
      MERGED_DATA_PART,
    );
    expect(worksheetCellFormula(dataXml, "E4")).toBeUndefined();
    expect(worksheetCellValue(dataXml, "E4")).toBe("20");
  });

  it("pins: mergeWorkbooksBytes matches the file merge and can omit the sheet index", async () => {
    const outcome = await mergeWorkbooksBytes({
      includeSheetIndex: false,
      inputs: [
        {
          bytes: await buildCorpusWorkbook({ shape: "table" }),
          name: "first.xlsx",
        },
        {
          bytes: await buildCorpusWorkbook({ shape: "range" }),
          name: "second.xlsx",
        },
      ],
    });

    expect(outcome.outputs[0]!.name).toBe("merged.xlsx");
    expect(outcome.result.metrics).toEqual({
      hiddenSheets: 4,
      inputFiles: 2,
      outputSheets: 8,
    });
    const workbookXml = await readPackagePart(
      outcome.outputs[0]!.bytes,
      CORPUS_PARTS.workbook,
    );
    expect(workbookXml).not.toContain("Sheet Index");
    expect(workbookXml).toContain('state="veryHidden"');
    expect(outcome.result.warnings.join("\n")).toMatch(
      /source worksheets were hidden in the merged workbook/u,
    );
  });
});

describe("corpus: consolidate", () => {
  it("pins: consolidate flattens visible worksheets into one values-only table", async () => {
    const directory = await createCorpusDirectory();
    const first = await writeCorpusWorkbook(directory, "first.xlsx", {
      shape: "table",
    });
    const second = await writeCorpusWorkbook(directory, "second.xlsx", {
      shape: "range",
    });
    const output = path.join(directory, "consolidated.xlsx");

    const result = await consolidateWorkbooks({
      headerRow: 3,
      inputs: [first, second],
      output,
    });

    // Both Data worksheets contribute; the Summary worksheets have no rows
    // below row 3, so they produce no table at all.
    expect(result.metrics).toEqual({
      inputFiles: 2,
      inputTables: 2,
      // Six region columns, three unnamed columns for the cells outside it,
      // and the three provenance columns.
      outputColumns: 12,
      // The table shape contributes its totals row and footer block; the range
      // shape contributes only its footer block.
      outputRows: 15,
    });
    const consolidated = XLSX.read(await readFile(output), { type: "buffer" });
    expect(consolidated.SheetNames).toEqual(["Consolidated"]);
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(
      consolidated.Sheets.Consolidated!,
      { defval: null },
    );
    expect(rows[0]).toMatchObject({
      Amount: 10,
      Client: "Client A",
      Doubled: 20,
      Group: "Alpha",
      Record: 1,
      _source_file: "first.xlsx",
      _source_row: 4,
      _source_sheet: "Data",
    });
    // Consolidation reads parsed values, so no formula survives.
    const consolidatedXml = await readPackagePart(
      await readWorkbookBytes(output),
      "xl/worksheets/sheet1.xml",
    );
    expect(consolidatedXml).not.toContain("<f>");
  });

  it("pins: header detection without headerRow picks the first non-empty row", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "first.xlsx", {
      shape: "range",
    });

    const tables = await readWorkbookTables(input);

    expect(tables.map((table) => table.source?.sheet)).toEqual([
      "Data",
      "Summary",
    ]);
    // Row 1 holds the report title, so the title becomes the first column name
    // and the real header row becomes a data row.
    expect(tables[0]?.columns[0]).toBe("Corpus allocation report");
    expect(tables[0]?.source?.firstDataRow).toBe(2);
  });

  it("pins: hidden worksheets are excluded unless includeHiddenSheets is set", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "first.xlsx", {
      shape: "range",
    });

    const visible = await readWorkbookTables(input);
    const everything = await readWorkbookTables(input, {
      includeHiddenSheets: true,
    });

    expect(visible.map((table) => table.source?.sheet)).toEqual([
      "Data",
      "Summary",
    ]);
    expect(everything.map((table) => table.source?.sheet)).toEqual([
      "Data",
      "Summary",
      "Hidden",
      "VeryHidden",
    ]);
  });
});
