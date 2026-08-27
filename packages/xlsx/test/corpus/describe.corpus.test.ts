/**
 * Conformance corpus: workbook inspection (`sheets.inspect`).
 *
 * Every cell of `CONTRACT.describe` reads `preserve`, and this file is what
 * holds all of them up. A read-only operation is the one column where a single
 * assertion can prove the whole thing: if the input package is byte-identical
 * after a describe, then no tracked structure was fixed, stripped or refused,
 * because nothing was written at all. The per-structure pins below add the
 * other half - that the structures the description is supposed to *report* are
 * reported, so "preserve" never degrades into "ignored".
 *
 * Fixtures are the paired corpus workbooks, so a capability that worked for
 * only one shape is a failing test rather than a review comment.
 */
import { readFile } from "node:fs/promises";

import { afterEach, describe, expect, it } from "vitest";

import { describeWorkbookBytes } from "../../src/bytes.js";
import { describeWorkbook } from "../../src/index.js";
import {
  buildCorpusWorkbook,
  cleanupCorpusDirectories,
  CORPUS_HEADER_ROW,
  CORPUS_HIDDEN_SHEET,
  CORPUS_LOCAL_NAME,
  CORPUS_RANGE_NAME,
  CORPUS_SHEET,
  CORPUS_SPLIT_COLUMN,
  CORPUS_SUMMARY_SHEET,
  CORPUS_TABLE_NAME,
  CORPUS_VERY_HIDDEN_SHEET,
  createCorpusDirectory,
  packagePartNames,
  writeCorpusWorkbook,
  type CorpusShape,
} from "./fixtures.js";
import { SHAPES } from "./symmetry.js";

afterEach(cleanupCorpusDirectories);

/** Every structure the generator can switch on, in one workbook. */
function everyStructure(
  shape: CorpusShape,
): Parameters<typeof buildCorpusWorkbook>[0] {
  return {
    shape,
    arrayFormula: true,
    comments: true,
    definedNames: true,
    dependents: true,
    footerBlock: true,
    formulas: shape === "table" ? "structured" : "a1",
    hiddenSheets: true,
    numberFormat: true,
    pivot: true,
    sharedFormula: true,
    summarySheet: true,
    uncachedFormula: true,
  };
}

describe("corpus: describe preserves every tracked structure", () => {
  it.each(SHAPES)(
    "invariant: describing a %s workbook leaves every input byte untouched",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const input = await writeCorpusWorkbook(
        directory,
        "corpus.xlsx",
        everyStructure(shape),
      );
      const before = await readFile(input);

      const { result } = await describeWorkbook(input, {
        includeHiddenSheets: true,
      });

      // The whole `describe` contract column in one assertion: an operation
      // that changed no byte cannot have fixed, stripped or refused anything.
      expect(Buffer.compare(await readFile(input), before)).toBe(0);
      // And it produced no output to have done it to.
      expect(result.artifacts).toEqual([]);
      expect(result.operation).toBe("sheets.inspect");
    },
  );

  it.each(SHAPES)(
    "invariant: describing a %s workbook keeps every package part",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const input = await writeCorpusWorkbook(
        directory,
        "corpus.xlsx",
        everyStructure(shape),
      );
      const before = await packagePartNames(await readFile(input));

      await describeWorkbook(input, { includeHiddenSheets: true });

      // Byte identity already implies this; naming the parts makes the pivot
      // cache, comment, drawing and table parts explicit in the corpus.
      expect(await packagePartNames(await readFile(input))).toEqual(before);
    },
  );

  it("invariant: describing a macro workbook leaves every input byte untouched", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsm", {
      shape: "table",
      macro: true,
    });
    const before = await readFile(input);

    await describeWorkbook(input);

    expect(Buffer.compare(await readFile(input), before)).toBe(0);
  });

  it.each(SHAPES)(
    "invariant: describing %s bytes never mutates the input array",
    async (shape) => {
      const bytes = await buildCorpusWorkbook(everyStructure(shape));
      const before = Uint8Array.from(bytes);

      await describeWorkbookBytes(
        { name: "corpus.xlsx", bytes },
        { includeHiddenSheets: true },
      );

      expect(bytes.length).toBe(before.length);
      expect(Buffer.compare(Buffer.from(bytes), Buffer.from(before))).toBe(0);
    },
  );
});

describe("corpus: describe reports the structures it covers", () => {
  it("pins: the description reports defined names without altering them", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "range",
      definedNames: true,
      summarySheet: true,
    });

    const { description, result } = await describeWorkbook(input);

    // Both the workbook-scoped name and the sheet-scoped one are reported.
    expect(description.namedRanges.map((range) => range.name).sort()).toEqual(
      [CORPUS_LOCAL_NAME, CORPUS_RANGE_NAME].sort(),
    );
    expect(
      description.namedRanges.find((range) => range.name === CORPUS_RANGE_NAME)
        ?.sheet,
    ).toBe(CORPUS_SHEET);
    expect(result.metrics.namedRanges).toBe(2);
  });

  it("pins: the description reports an Excel Table's declared headers", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "table",
    });

    const { description, result } = await describeWorkbook(input);

    expect(description.excelTables).toHaveLength(1);
    const table = description.excelTables[0]!;
    expect(table.name).toBe(CORPUS_TABLE_NAME);
    expect(table.sheet).toBe(CORPUS_SHEET);
    expect(table.headers).toContain(CORPUS_SPLIT_COLUMN);
    expect(result.metrics.excelTables).toBe(1);
  });

  it("pins: a totals row is described as worksheet content and never removed", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "table",
      totalsRow: true,
    });
    const before = await readFile(input);

    const { description } = await describeWorkbook(input, {
      headerRow: CORPUS_HEADER_ROW,
    });

    // The bindings disagree about deleting a totals row; describe deletes
    // nothing, so the row is simply counted as one of the sheet's data rows.
    const data = description.sheets.find(
      (sheet) => sheet.name === CORPUS_SHEET,
    );
    expect(data?.dataRowCount).toBeGreaterThan(0);
    expect(Buffer.compare(await readFile(input), before)).toBe(0);
  });

  it("pins: a formula cell is sampled by its cached value, not recalculated", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "range",
      formulas: "a1",
      summarySheet: true,
      uncachedFormula: true,
    });

    const { description } = await describeWorkbook(input, {
      headerRow: CORPUS_HEADER_ROW,
    });

    // Summary!B4 carries a formula with no cached value. The inspection reads
    // stored values only, so it contributes nothing rather than a computed
    // guess - an inspection never invents a value the workbook does not hold.
    const summary = description.sheets.find(
      (sheet) => sheet.name === CORPUS_SUMMARY_SHEET,
    );
    expect(summary).toBeDefined();
    for (const column of summary!.columns) {
      for (const value of column.sampleValues) {
        expect(value).not.toBeNull();
      }
    }

    // Occupancy is the other question, and it has the other answer: the cell
    // exists, so its row counts. Samples stay stored-values-only; a row is not
    // blank merely because nothing cached a result into it.
    expect(summary!.dataRowCount).toBeGreaterThan(0);
    expect(summary!.headerRow).toBeDefined();
  });

  it.each(SHAPES)(
    "pins: hidden and very hidden worksheets are reported by state (%s)",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
        shape,
        hiddenSheets: true,
      });

      const visibleOnly = await describeWorkbook(input);
      expect(
        visibleOnly.description.sheets.map((sheet) => sheet.name),
      ).not.toContain(CORPUS_HIDDEN_SHEET);

      const all = await describeWorkbook(input, { includeHiddenSheets: true });
      const byName = new Map(
        all.description.sheets.map((sheet) => [sheet.name, sheet.visibility]),
      );
      expect(byName.get(CORPUS_SHEET)).toBe("visible");
      expect(byName.get(CORPUS_HIDDEN_SHEET)).toBe("hidden");
      expect(byName.get(CORPUS_VERY_HIDDEN_SHEET)).toBe("very-hidden");
    },
  );

  it.each(SHAPES)(
    "invariant: the file and byte surfaces describe a %s workbook identically",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const options = everyStructure(shape);
      const input = await writeCorpusWorkbook(
        directory,
        "corpus.xlsx",
        options,
      );

      const fromFile = await describeWorkbook(input, {
        includeHiddenSheets: true,
      });
      const fromBytes = await describeWorkbookBytes(
        { name: "corpus.xlsx", bytes: await buildCorpusWorkbook(options) },
        { includeHiddenSheets: true },
      );

      expect(fromBytes.description).toEqual(fromFile.description);
      expect(fromBytes.result).toEqual(fromFile.result);
    },
  );
});
