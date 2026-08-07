/**
 * Conformance corpus: what every split mode does to every tracked structure
 * TODAY. These tests pin current behaviour, lossy outcomes included, so that
 * Phase 1 can re-express the split engine on the layered model and prove it
 * changed nothing it did not mean to change.
 */
import { readFile } from "node:fs/promises";
import path from "node:path";

import { afterEach, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import { splitWorkbookBytes, planSplitWorkbookBytes } from "../../src/bytes.js";
import {
  planSplitWorkbookByColumn,
  splitWorkbookByColumn,
} from "../../src/index.js";
import {
  buildCorpusWorkbook,
  cleanupCorpusDirectories,
  conditionalFormattingSqref,
  CORPUS_PARTS,
  CORPUS_RANGE_NAME,
  CORPUS_SHEET,
  CORPUS_SIDE_NOTE,
  CORPUS_SPLIT_COLUMN,
  CORPUS_TABLE_NAME,
  CORPUS_VBA_BYTES,
  createCorpusDirectory,
  dataValidationSqref,
  hasPackagePart,
  packagePartNames,
  readPackagePart,
  readWorkbookBytes,
  worksheetCellFormula,
  worksheetCellValue,
  worksheetRowNumbers,
  writeCorpusWorkbook,
  type CorpusShape,
} from "./fixtures.js";

const SHAPES: readonly CorpusShape[] = ["table", "range"];

afterEach(cleanupCorpusDirectories);

describe("corpus: all-worksheet split", () => {
  it.each(SHAPES)(
    "pins: the %s binding filters every worksheet that carries the split column",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
        shape,
      });
      const outputDirectory = path.join(directory, "out");

      const result = await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory,
      });

      expect(
        result.artifacts.map((artifact) => path.basename(artifact.path)),
      ).toEqual(["Alpha.xlsx", "Beta.xlsx", "Gamma.xlsx"]);
      expect(result.summary).toMatchObject({
        column: CORPUS_SPLIT_COLUMN,
        copiedUnchangedSheets: ["Summary", "VeryHidden"],
        filteredSheets: ["Data", "Hidden"],
        valuesOnly: false,
      });
      expect(result.metrics).toMatchObject({
        groups: 3,
        outputFiles: 3,
        sheetsCopiedUnchanged: 2,
        sheetsFiltered: 2,
        valuesOnly: 0,
      });
      // The table binding stops at the table's own last data row; the range
      // binding treats every row below the header as data, so it also scans
      // the totals row and the footer block and reports them as skipped.
      expect(result.metrics.inputRows).toBe(shape === "table" ? 9 : 12);
      expect(result.metrics.skippedRows).toBe(shape === "table" ? 0 : 3);
    },
  );

  it.each(SHAPES)(
    "invariant (128a310): the %s binding renumbers surviving rows and their cell references",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
        shape,
      });
      await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, "out"),
      });

      const alpha = await readWorkbookBytes(
        path.join(directory, "out", "Alpha.xlsx"),
      );
      const dataXml = await readPackagePart(alpha, CORPUS_PARTS.dataSheet);
      // Source rows 4, 6 and 9 hold the Alpha records; they land on 4, 5 and 6.
      expect(worksheetRowNumbers(dataXml).slice(0, 4)).toEqual([1, 3, 4, 5]);
      expect(worksheetCellValue(dataXml, "D4")).toBe("10");
      expect(worksheetCellValue(dataXml, "D5")).toBe("30");
      expect(worksheetCellValue(dataXml, "D6")).toBe("60");
      // Records 1, 3 and 6 keep their identity while their rows compact.
      expect(worksheetCellValue(dataXml, "A4")).toBe("1");
      expect(worksheetCellValue(dataXml, "A5")).toBe("3");
      expect(worksheetCellValue(dataXml, "A6")).toBe("6");
    },
  );

  it.each(SHAPES)(
    "pins: the %s binding copies every source package part into every output",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
        shape,
        pivot: true,
      });
      await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, "out"),
      });

      const sourceParts = await packagePartNames(
        await readWorkbookBytes(input),
      );
      const alpha = await readWorkbookBytes(
        path.join(directory, "out", "Alpha.xlsx"),
      );
      // Every source part reaches the output except the pivot parts, which
      // the Tier-1 confidentiality pass removes because a pivot cache carries
      // a private copy of every group's rows (tier1-gaps.corpus.test.ts).
      expect(await packagePartNames(alpha)).toEqual(
        sourceParts.filter((part) => !part.startsWith("xl/pivot")),
      );
      // Comments, drawings and the calculation chain all survive, whether or
      // not they still describe the filtered rows.
      expect(await hasPackagePart(alpha, CORPUS_PARTS.comments)).toBe(true);
      expect(await hasPackagePart(alpha, CORPUS_PARTS.pivotTable)).toBe(false);
      expect(await hasPackagePart(alpha, CORPUS_PARTS.calcChain)).toBe(true);
    },
  );

  it.each(SHAPES)(
    "invariant: the %s binding preserves hidden and very-hidden worksheets",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
        shape,
      });
      await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, "out"),
      });

      const alpha = await readFile(path.join(directory, "out", "Alpha.xlsx"));
      const workbook = XLSX.read(alpha, { type: "buffer" });
      expect(workbook.SheetNames).toEqual([
        "Data",
        "Summary",
        "Hidden",
        "VeryHidden",
      ]);
      expect(workbook.Workbook?.Sheets?.map((sheet) => sheet.Hidden)).toEqual([
        0, 0, 1, 2,
      ]);
      // A hidden worksheet carrying the split column is filtered like any
      // other; visibility does not exclude it from an all-worksheet split.
      const hiddenXml = await readPackagePart(
        new Uint8Array(alpha),
        CORPUS_PARTS.hiddenSheet,
      );
      expect(worksheetRowNumbers(hiddenXml)).toEqual([1, 2, 3]);
    },
  );

  it.each(SHAPES)(
    "invariant: the %s binding round-trips a macro workbook as .xlsm",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const input = await writeCorpusWorkbook(directory, "corpus.xlsm", {
        shape,
        macro: true,
      });
      const result = await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, "out"),
      });

      expect(
        result.artifacts.every((artifact) => artifact.path.endsWith(".xlsm")),
      ).toBe(true);
      expect(
        result.artifacts.every(
          (artifact) =>
            artifact.mediaType ===
            "application/vnd.ms-excel.sheet.macroEnabled.12",
        ),
      ).toBe(true);
      const alpha = await readWorkbookBytes(result.artifacts[0]!.path);
      const archive = await readPackagePart(alpha, CORPUS_PARTS.contentTypes);
      expect(archive).toContain("macroEnabled.main+xml");
      expect(await hasPackagePart(alpha, CORPUS_PARTS.vbaProject)).toBe(true);
      const vba = await import("jszip").then(async (module) => {
        const zip = await module.default.loadAsync(alpha);
        return zip.file(CORPUS_PARTS.vbaProject)!.async("uint8array");
      });
      expect([...vba]).toEqual([...CORPUS_VBA_BYTES]);
    },
  );

  it("pins: the range binding deletes the footer block, the table binding keeps it", async () => {
    const directory = await createCorpusDirectory();
    const tableInput = await writeCorpusWorkbook(directory, "table.xlsx", {
      shape: "table",
    });
    const rangeInput = await writeCorpusWorkbook(directory, "range.xlsx", {
      shape: "range",
    });
    await splitWorkbookByColumn({
      column: CORPUS_SPLIT_COLUMN,
      input: tableInput,
      outputDirectory: path.join(directory, "table-out"),
    });
    await splitWorkbookByColumn({
      column: CORPUS_SPLIT_COLUMN,
      input: rangeInput,
      outputDirectory: path.join(directory, "range-out"),
    });

    const tableXml = await readPackagePart(
      await readWorkbookBytes(path.join(directory, "table-out", "Alpha.xlsx")),
      CORPUS_PARTS.dataSheet,
    );
    const rangeXml = await readPackagePart(
      await readWorkbookBytes(path.join(directory, "range-out", "Alpha.xlsx")),
      CORPUS_PARTS.dataSheet,
    );
    // The table binding stops filtering at the table's last data row, so the
    // totals row and the footer block survive, renumbered onto rows 7 and 9.
    expect(worksheetRowNumbers(tableXml)).toEqual([1, 3, 4, 5, 6, 7, 9]);
    // The range binding treats everything below the header as data, so the
    // totals row and the footer block are removed as unmatched rows.
    expect(worksheetRowNumbers(rangeXml)).toEqual([1, 3, 4, 5, 6]);
  });

  it("pins: A1 formulas force the table binding onto the plain row-removal path", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "table",
      formulas: "a1",
    });
    const result = await splitWorkbookByColumn({
      column: CORPUS_SPLIT_COLUMN,
      input,
      outputDirectory: path.join(directory, "out"),
    });

    expect(result.warnings.join("\n")).toMatch(
      /contained formulas tied to row positions/u,
    );
    const alpha = await readWorkbookBytes(
      path.join(directory, "out", "Alpha.xlsx"),
    );
    // The table part is left claiming the original range because the fallback
    // path never touches it.
    expect(await readPackagePart(alpha, CORPUS_PARTS.table)).toContain(
      'ref="A3:F10"',
    );
  });

  it("pins: structured-reference formulas let the table binding compact its table part", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "table",
      formulas: "structured",
    });
    const result = await splitWorkbookByColumn({
      column: CORPUS_SPLIT_COLUMN,
      input,
      outputDirectory: path.join(directory, "out"),
    });

    expect(result.warnings).toEqual([]);
    const alpha = await readWorkbookBytes(
      path.join(directory, "out", "Alpha.xlsx"),
    );
    const tableXml = await readPackagePart(alpha, CORPUS_PARTS.table);
    expect(tableXml).toContain('ref="A3:F7"');
    expect(tableXml).toContain('<autoFilter ref="A3:F6"/>');
    const dataXml = await readPackagePart(alpha, CORPUS_PARTS.dataSheet);
    // The compacting path leaves rows below the table where they were.
    expect(worksheetRowNumbers(dataXml)).toEqual([1, 3, 4, 5, 6, 7, 12]);
    expect(worksheetCellFormula(dataXml, "E5")).toBe(
      `${CORPUS_TABLE_NAME}[[#This Row],[Amount]]*2`,
    );
  });

  it("pins: a single array formula anywhere on the sheet withdraws table compaction", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      arrayFormula: true,
      formulas: "structured",
      shape: "table",
    });
    const result = await splitWorkbookByColumn({
      column: CORPUS_SPLIT_COLUMN,
      input,
      outputDirectory: path.join(directory, "out"),
    });

    // tableCanBeCompacted scans the whole worksheet part, not just the table
    // range, so one t="array" formula outside the table is enough to send the
    // sheet down the plain row-removal path.
    expect(result.warnings.join("\n")).toMatch(
      /contained formulas tied to row positions/u,
    );
    const alpha = await readWorkbookBytes(
      path.join(directory, "out", "Alpha.xlsx"),
    );
    expect(await readPackagePart(alpha, CORPUS_PARTS.table)).toContain(
      'ref="A3:F10"',
    );
    const dataXml = await readPackagePart(alpha, CORPUS_PARTS.dataSheet);
    // The array formula's own row survives, and its range reference and cached
    // value are copied through without adjustment.
    expect(dataXml).toContain('<f t="array" ref="G4:G4">SUM(D4:D9)</f>');
    expect(worksheetCellValue(dataXml, "G4")).toBe("210");
  });

  it.each(SHAPES)(
    "invariant: a values-only %s split bakes cached values and removes the calculation chain",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
        shape,
        uncachedFormula: true,
      });
      const result = await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, "out"),
        values: true,
      });

      expect(result.metrics.valuesOnly).toBe(1);
      expect(result.metrics.formulaCellsConverted).toBeGreaterThan(0);
      // Reported once per output workbook, because each output is converted.
      // Summary!B4 has never been calculated, so it contributes three. The
      // rest are results the Tier-1 pass cleared before the conversion because
      // they covered rows the group does not receive: Summary!B2 in every
      // output and Summary!B3 in the two that lose Data row 4, plus the totals
      // row and the footer aggregate on the table binding, whose data rows are
      // filtered while those two rows stay.
      expect(result.metrics.formulaCellsBlankedForRemovedRows).toBe(
        shape === "table" ? 11 : 5,
      );
      expect(result.metrics.formulaCellsWithoutCachedValues).toBe(
        shape === "table" ? 14 : 8,
      );
      expect(result.warnings.join("\n")).toMatch(/Summary!B4/u);

      const alpha = await readWorkbookBytes(
        path.join(directory, "out", "Alpha.xlsx"),
      );
      expect(await hasPackagePart(alpha, CORPUS_PARTS.calcChain)).toBe(false);
      const workbookRelationships = await readPackagePart(
        alpha,
        "xl/_rels/workbook.xml.rels",
      );
      expect(workbookRelationships).not.toContain("calcChain");
      const dataXml = await readPackagePart(alpha, CORPUS_PARTS.dataSheet);
      expect(worksheetCellFormula(dataXml, "E4")).toBeUndefined();
      expect(worksheetCellValue(dataXml, "E4")).toBe("20");
    },
  );

  it.each(SHAPES)(
    "pins: planning a %s all-worksheet split reports outputs without writing them",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
        shape,
      });
      const plan = await planSplitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, "out"),
      });

      expect(plan.operation).toBe("sheets.split-by-column");
      expect(plan.outputs.map((output) => path.basename(output.path))).toEqual([
        "Alpha.xlsx",
        "Beta.xlsx",
        "Gamma.xlsx",
      ]);
      expect(plan.outputs.every((output) => output.exists)).toBe(false);
      expect(plan.metrics.groups).toBe(3);
    },
  );
});

describe("corpus: table-selection split", () => {
  it("pins: a preserved table split refuses to relocate a row carrying A1 formulas", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "table",
      formulas: "a1",
    });

    await expect(
      splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, "out"),
        table: CORPUS_TABLE_NAME,
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_PRESERVE_FORMULA" });
  });

  it("pins: values-only conversion unblocks the preserved-split formula refusal", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "table",
      formulas: "a1",
    });

    const result = await splitWorkbookByColumn({
      column: CORPUS_SPLIT_COLUMN,
      input,
      outputDirectory: path.join(directory, "out"),
      table: CORPUS_TABLE_NAME,
      values: true,
    });
    expect(result.metrics.outputFiles).toBe(3);
    const alpha = await readWorkbookBytes(
      path.join(directory, "out", "corpus-Alpha.xlsx"),
    );
    const dataXml = await readPackagePart(alpha, CORPUS_PARTS.dataSheet);
    expect(worksheetCellFormula(dataXml, "E5")).toBeUndefined();
    expect(worksheetCellValue(dataXml, "E5")).toBe("60");
  });

  it("invariant (727239a): whole rows leave with the table row, including cells outside the table", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "table",
      formulas: "structured",
    });

    await splitWorkbookByColumn({
      column: CORPUS_SPLIT_COLUMN,
      input,
      outputDirectory: path.join(directory, "out"),
      table: CORPUS_TABLE_NAME,
    });

    const beta = await readWorkbookBytes(
      path.join(directory, "out", "corpus-Beta.xlsx"),
    );
    const betaSharedStrings = await readPackagePart(
      beta,
      CORPUS_PARTS.sharedStrings,
    );
    const betaData = await readPackagePart(beta, CORPUS_PARTS.dataSheet);
    // Data!H6 sits outside the table on an Alpha row. The row leaves whole, so
    // no styled shell and no out-of-table cell is left behind in the Beta
    // output. The string table is copied unchanged, so the note's text is
    // still interned there; what matters is that no cell references it.
    expect(betaSharedStrings).toContain(CORPUS_SIDE_NOTE);
    expect(betaData).not.toContain('r="H6"');
    expect(betaData).not.toContain('r="H5"');
    expect(worksheetRowNumbers(betaData)).toEqual([1, 3, 4, 5, 6, 12]);
  });

  it("pins: preserveWorkbook:false rebuilds a compact single-worksheet workbook", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "table",
      formulas: "a1",
      pivot: true,
    });

    const result = await splitWorkbookByColumn({
      column: CORPUS_SPLIT_COLUMN,
      input,
      outputDirectory: path.join(directory, "out"),
      preserveWorkbook: false,
      table: CORPUS_TABLE_NAME,
    });

    expect(result.metrics).toMatchObject({ groups: 3, inputRows: 6 });
    const alpha = await readWorkbookBytes(
      path.join(directory, "out", "corpus-Alpha.xlsx"),
    );
    const parts = await packagePartNames(alpha);
    // Everything the corpus tracks is gone: the table part, the pivot cache,
    // comments, conditional formatting, data validation and the other sheets.
    expect(parts).not.toContain(CORPUS_PARTS.table);
    expect(parts).not.toContain(CORPUS_PARTS.pivotCacheRecords);
    expect(parts).not.toContain(CORPUS_PARTS.comments);
    expect(parts).not.toContain(CORPUS_PARTS.calcChain);
    const workbook = XLSX.read(await readFile(result.artifacts[0]!.path), {
      type: "buffer",
    });
    expect(workbook.SheetNames).toEqual([CORPUS_SHEET]);
    const rebuilt = await readPackagePart(alpha, "xl/worksheets/sheet1.xml");
    expect(rebuilt).not.toContain("conditionalFormatting");
    expect(rebuilt).not.toContain("dataValidation");
    expect(rebuilt).not.toContain("mergeCell");
    // Formulas become their cached values because the rows are rewritten from
    // parsed cell values rather than copied.
    expect(worksheetCellFormula(rebuilt, "E2")).toBeUndefined();
    expect(worksheetCellValue(rebuilt, "E2")).toBe("20");
  });
});

describe("corpus: named-range split", () => {
  it("pins: a named-range split produces a compact workbook and drops every other structure", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "range",
      pivot: true,
    });

    const result = await splitWorkbookByColumn({
      column: CORPUS_SPLIT_COLUMN,
      input,
      outputDirectory: path.join(directory, "out"),
      range: CORPUS_RANGE_NAME,
    });

    expect(result.metrics).toMatchObject({
      groups: 3,
      inputRows: 6,
      outputFiles: 3,
      skippedRows: 0,
    });
    const alpha = await readWorkbookBytes(
      path.join(directory, "out", "corpus-Alpha.xlsx"),
    );
    const workbookXml = await readPackagePart(alpha, CORPUS_PARTS.workbook);
    expect(workbookXml).not.toContain(CORPUS_RANGE_NAME);
    expect(workbookXml).not.toContain("pivotCache");
    const rebuilt = await readPackagePart(alpha, "xl/worksheets/sheet1.xml");
    expect(worksheetRowNumbers(rebuilt)).toEqual([1, 2, 3, 4]);
    expect(conditionalFormattingSqref(rebuilt)).toBeUndefined();
    expect(dataValidationSqref(rebuilt)).toBeUndefined();
  });

  it("pins: a named range cannot preserve the workbook", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "range",
    });

    await expect(
      splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, "out"),
        preserveWorkbook: true,
        range: CORPUS_RANGE_NAME,
      }),
    ).rejects.toMatchObject({
      code: "XLSX_SPLIT_PRESERVE_REQUIRES_TABLE",
    });
  });
});

describe("corpus: worksheet-selection split", () => {
  it.each(SHAPES)(
    "pins: a %s worksheet split reads the footer block as a blank group",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
        shape,
      });

      const result = await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, "out"),
        headerRow: 3,
        preserveWorkbook: false,
        sheet: CORPUS_SHEET,
      });

      // The worksheet binding has no region boundary, so the totals row and
      // the footer block become records with a blank key and land in their own
      // "blank" output workbook.
      expect(
        result.artifacts.map((artifact) => path.basename(artifact.path)),
      ).toEqual([
        "corpus-Alpha.xlsx",
        "corpus-Beta.xlsx",
        "corpus-Gamma.xlsx",
        "corpus-blank.xlsx",
      ]);
      expect(result.metrics.groups).toBe(4);
    },
  );

  it("pins: includeBlank:false drops the blank-key rows and warns", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "range",
    });

    const result = await splitWorkbookByColumn({
      column: CORPUS_SPLIT_COLUMN,
      includeBlank: false,
      input,
      outputDirectory: path.join(directory, "out"),
      headerRow: 3,
      preserveWorkbook: false,
      sheet: CORPUS_SHEET,
    });

    expect(result.metrics.groups).toBe(3);
    expect(result.metrics.skippedRows).toBeGreaterThan(0);
    expect(result.warnings.join("\n")).toMatch(/Skipped \d+ rows? with blank/u);
  });
});

describe("corpus: byte surface", () => {
  it("pins: splitWorkbookBytes mirrors the table-selection file split", async () => {
    const bytes = await buildCorpusWorkbook({
      shape: "table",
      formulas: "structured",
    });

    const outcome = await splitWorkbookBytes({
      column: CORPUS_SPLIT_COLUMN,
      input: { bytes, name: "corpus.xlsx" },
      table: CORPUS_TABLE_NAME,
    });

    expect(outcome.outputs.map((output) => output.name)).toEqual([
      "corpus-Alpha.xlsx",
      "corpus-Beta.xlsx",
      "corpus-Gamma.xlsx",
    ]);
    const alpha = outcome.outputs[0]!.bytes;
    expect(await readPackagePart(alpha, CORPUS_PARTS.table)).toContain(
      'ref="A3:F7"',
    );
    expect(await hasPackagePart(alpha, CORPUS_PARTS.pivotTable)).toBe(false);
    expect(outcome.result.metrics).toMatchObject({
      groups: 3,
      inputRows: 6,
      outputFiles: 3,
      outputRows: 6,
    });
  });

  it("pins: planSplitWorkbookBytes names outputs without building bytes", async () => {
    const bytes = await buildCorpusWorkbook({
      shape: "range",
    });

    const plan = await planSplitWorkbookBytes({
      column: CORPUS_SPLIT_COLUMN,
      input: { bytes, name: "corpus.xlsx" },
      range: CORPUS_RANGE_NAME,
    });

    expect(plan.outputs.map((output) => output.path)).toEqual([
      "corpus-Alpha.xlsx",
      "corpus-Beta.xlsx",
      "corpus-Gamma.xlsx",
    ]);
    expect(plan.outputs.every((output) => output.exists)).toBe(false);
  });
});
