/**
 * Conformance corpus: the Tier-1 gaps ARCHITECTURE.md's migration plan names
 * for Phase 0.
 *
 * Each gap appears twice: a passing test pinning what the library does today,
 * and an adjacent `it.fails` test asserting the behaviour the architecture
 * requires. The `it.fails` form keeps CI green while the gap is documented in
 * executable form; Phase 2 deletes the `.fails` and the pin flips to the new
 * expectation. Pairing them also guarantees the expected failure fails for the
 * documented reason rather than because the fixture broke.
 *
 * Every gap this file was opened for is now closed: pivot caches crossing
 * outputs, stale cached aggregates in values mode and the stale calculation
 * chain by the Tier-1 utilities, and the dependent references by Phase 1's
 * move onto the layered engine. There are no `it.fails` tests left here; each
 * `.fails` twin became a `Tier-1 fix: ...` test and its pin records the new
 * output.
 */
import path from "node:path";

import { afterEach, describe, expect, it } from "vitest";

import {
  splitWorkbookByColumn,
  type SplitWorkbookByColumnResult,
} from "../../src/index.js";
import {
  calcChainReferences,
  cleanupCorpusDirectories,
  conditionalFormattingSqref,
  CORPUS_PARTS,
  CORPUS_SPLIT_COLUMN,
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

interface CorpusSplitOptions {
  shape: "table" | "range";
  formulas?: "a1" | "structured";
  pivot?: boolean;
  sharedFormula?: boolean;
  values?: boolean;
}

afterEach(cleanupCorpusDirectories);

/** Split the corpus by Group and return the Alpha output with its report. */
async function alphaSplit(options: CorpusSplitOptions): Promise<{
  bytes: Uint8Array;
  result: SplitWorkbookByColumnResult;
}> {
  const directory = await createCorpusDirectory();
  const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
    formulas: options.formulas ?? "a1",
    pivot: options.pivot ?? false,
    shape: options.shape,
    sharedFormula: options.sharedFormula ?? false,
  });
  const outputDirectory = path.join(directory, "out");
  const result = await splitWorkbookByColumn({
    column: CORPUS_SPLIT_COLUMN,
    input,
    outputDirectory,
    values: options.values ?? false,
  });
  return {
    bytes: await readWorkbookBytes(path.join(outputDirectory, "Alpha.xlsx")),
    result,
  };
}

/** Split the corpus by Group and return the Alpha output's bytes. */
async function alphaOutput(options: CorpusSplitOptions): Promise<Uint8Array> {
  return (await alphaSplit(options)).bytes;
}

describe("corpus: Tier-1 fix - pivot caches cross split outputs", () => {
  it("pins: a split output carries no pivot part, relationship or content type", async () => {
    const alpha = await alphaOutput({ shape: "range", pivot: true });

    // A pivot cache cannot be filtered alongside the rows it caches yet, so
    // the whole pivot leaves with them: parts, both relationship entries, the
    // content-type overrides and the workbook's pivotCaches registry.
    expect(
      (await packagePartNames(alpha)).filter((part) =>
        part.startsWith("xl/pivot"),
      ),
    ).toEqual([]);
    expect(await readPackagePart(alpha, CORPUS_PARTS.contentTypes)).not.toMatch(
      /pivot/u,
    );
    expect(await readPackagePart(alpha, CORPUS_PARTS.workbook)).not.toMatch(
      /pivotCache/u,
    );
    expect(
      await readPackagePart(alpha, "xl/_rels/workbook.xml.rels"),
    ).not.toMatch(/pivotCacheDefinition/u);
    expect(
      await readPackagePart(alpha, "xl/worksheets/_rels/sheet2.xml.rels"),
    ).not.toMatch(/pivotTable/u);
  });

  it("Tier-1 fix: pivot-cache records of another group never reach a group's split output", async () => {
    const alpha = await alphaSplit({ shape: "range", pivot: true });

    // Confidentiality: the Alpha recipient must not be able to read Beta's and
    // Gamma's rows out of a pivot cache that travelled with the file.
    for (const part of [
      CORPUS_PARTS.pivotCacheRecords,
      CORPUS_PARTS.pivotCacheDefinition,
      CORPUS_PARTS.pivotTable,
    ]) {
      expect(await hasPackagePart(alpha.bytes, part)).toBe(false);
    }
    expect(alpha.result.warnings.join("\n")).toMatch(
      /their caches contained rows from other groups/u,
    );
    // One pivot table removed from each of the three outputs.
    expect(alpha.result.metrics.pivotTablesRemoved).toBe(3);
  });
});

describe("corpus: Tier-1 fix - stale cached aggregates in values mode", () => {
  it("pins: a values-only split blanks a cross-group aggregate and keeps a single-row reference", async () => {
    const summary = await readPackagePart(
      await alphaOutput({ shape: "range", values: true }),
      CORPUS_PARTS.summarySheet,
    );

    // Summary!B2 is SUM(Data!D4:D9); rows 5, 7 and 8 left with the other
    // groups, so its cached 210 no longer describes anything in this file.
    expect(worksheetCellFormula(summary, "B2")).toBeUndefined();
    expect(worksheetCellValue(summary, "B2")).toBeUndefined();
    // Summary!B3 is Data!D4, a row Alpha keeps, so its cached value stands.
    expect(worksheetCellFormula(summary, "B3")).toBeUndefined();
    expect(worksheetCellValue(summary, "B3")).toBe("10");
  });

  it("Tier-1 fix: a values-only split does not bake an aggregate computed over every group's rows", async () => {
    const alpha = await alphaSplit({ shape: "range", values: true });
    const summary = await readPackagePart(
      alpha.bytes,
      CORPUS_PARTS.summarySheet,
    );

    // A total presented as "Total across every record" must not read as this
    // group's total when it was computed over every group's rows. Blanking is
    // the conservative answer: the cell is visibly empty and reported, rather
    // than quietly wrong.
    expect(worksheetCellValue(summary, "B2")).toBeUndefined();
    expect(alpha.result.warnings.join("\n")).toMatch(/Summary!B2/u);
    expect(alpha.result.warnings.join("\n")).toMatch(
      /rows that are not part of this group/u,
    );
    // Summary!B2 clears in all three outputs; Summary!B3 reads Data!D4, which
    // only Alpha keeps, so it clears in the Beta and Gamma outputs too.
    expect(alpha.result.metrics.formulaCellsBlankedForRemovedRows).toBe(5);
  });
});

describe("corpus: Tier-1 fix - stale calculation chain", () => {
  it("pins: a non-values split prunes and renumbers the calculation chain", async () => {
    const calcChain = await readPackagePart(
      await alphaOutput({ shape: "range" }),
      CORPUS_PARTS.calcChain,
    );

    // E4 stays put, E6 and E9 follow their rows to 5 and 6, and the entries
    // for the deleted rows 5, 7, 8 and the footer's B12 are gone. The Summary
    // entries are untouched because that sheet lost no rows.
    expect(calcChainReferences(calcChain)).toEqual([
      "E4",
      "E5",
      "E6",
      "B2",
      "B3",
    ]);
  });

  it("Tier-1 fix: the calculation chain does not reference cells a split deleted", async () => {
    const alpha = await alphaSplit({ shape: "range" });
    // The chain is kept whenever entries survive, and removed outright when
    // none do; either way nothing in it may name a deleted cell.
    const references = (await hasPackagePart(
      alpha.bytes,
      CORPUS_PARTS.calcChain,
    ))
      ? calcChainReferences(
          await readPackagePart(alpha.bytes, CORPUS_PARTS.calcChain),
        )
      : [];

    // Alpha keeps three data rows, so only E4, E5 and E6 remain calculable on
    // the Data worksheet.
    expect(references.filter((reference) => reference.startsWith("E"))).toEqual(
      ["E4", "E5", "E6"],
    );
    expect(references).not.toContain("B12");
    // Four entries for Alpha, five for Beta and six for Gamma.
    expect(alpha.result.metrics.calcChainEntriesRemoved).toBe(15);
  });
});

describe("corpus: Tier-1 fix - dependent references after row compaction", () => {
  /**
   * Phase 1 closed this gap, on both bindings.
   *
   * The previous engine renumbered the `r` attribute of each surviving `<row>`
   * and of every `<c>` inside it, and nothing else: `mergeCells`,
   * `conditionalFormatting/@sqref`, `dataValidation/@sqref`, `hyperlink/@ref`,
   * `dimension/@ref`, shared-formula `@ref` and the text inside `<f>` all kept
   * their original row numbers while the cells they described moved up.
   *
   * The split now edits through L1, whose row-relocation pass rewrites every
   * one of those in the same traversal, so the bug class is unrepresentable
   * rather than merely fixed: there is no way to move a row that skips it.
   */

  it("pins: dependent references follow the rows they describe", async () => {
    const dataXml = await readPackagePart(
      await alphaOutput({ shape: "range", sharedFormula: true }),
      CORPUS_PARTS.dataSheet,
    );

    // Row 6 (an Alpha record) moved to row 5; row 9 moved to row 6.
    expect(worksheetCellValue(dataXml, "D5")).toBe("30");
    expect(worksheetCellValue(dataXml, "D6")).toBe("60");
    // Everything that referred to those rows refers to their new numbers.
    expect(mergedCellReferences(dataXml)).toEqual(["A1:F1", "H5:I5"]);
    expect(conditionalFormattingSqref(dataXml)).toBe("D4:D6");
    expect(dataValidationSqref(dataXml)).toBe("A4:A6");
    expect(hyperlinkReferences(dataXml)).toEqual(["A5"]);
    expect(worksheetCellFormula(dataXml, "E5")).toBe("D5*2");
    expect(dataXml).toContain('ref="F4:F6"');
  });

  it("Tier-1 fix: merged ranges follow the rows they cover", async () => {
    const dataXml = await readPackagePart(
      await alphaOutput({ shape: "range" }),
      CORPUS_PARTS.dataSheet,
    );

    // H6:I6 covered the record that now lives on row 5.
    expect(mergedCellReferences(dataXml)).toEqual(["A1:F1", "H5:I5"]);
  });

  it("Tier-1 fix: conditional-formatting sqref shrinks with the rows it covers", async () => {
    const dataXml = await readPackagePart(
      await alphaOutput({ shape: "range" }),
      CORPUS_PARTS.dataSheet,
    );

    expect(conditionalFormattingSqref(dataXml)).toBe("D4:D6");
  });

  it("Tier-1 fix: data-validation sqref shrinks with the rows it covers", async () => {
    const dataXml = await readPackagePart(
      await alphaOutput({ shape: "range" }),
      CORPUS_PARTS.dataSheet,
    );

    expect(dataValidationSqref(dataXml)).toBe("A4:A6");
  });

  it("Tier-1 fix: hyperlinks follow the row they decorate", async () => {
    const dataXml = await readPackagePart(
      await alphaOutput({ shape: "range" }),
      CORPUS_PARTS.dataSheet,
    );

    expect(hyperlinkReferences(dataXml)).toEqual(["A5"]);
  });

  it("Tier-1 fix: A1 formulas on a plain worksheet are rewritten when their row moves", async () => {
    const dataXml = await readPackagePart(
      await alphaOutput({ shape: "range" }),
      CORPUS_PARTS.dataSheet,
    );

    // E5 doubles the amount on its own row. Before compaction that amount was
    // D6; it is D5 now, and the formula says so, so the cell still doubles the
    // record it was written for instead of a different one.
    expect(worksheetCellFormula(dataXml, "E5")).toBe("D5*2");
    expect(worksheetCellFormula(dataXml, "E6")).toBe("D6*2");
  });

  it("Tier-1 fix: shared-formula ranges shrink with the rows they span", async () => {
    const dataXml = await readPackagePart(
      await alphaOutput({ shape: "range", sharedFormula: true }),
      CORPUS_PARTS.dataSheet,
    );

    expect(dataXml).toContain('ref="F4:F6"');
    expect(dataXml).not.toContain('ref="F4:F9"');
  });

  it("Tier-1 fix: the Excel Table binding fixes dependent references too", async () => {
    // The compacting table path resizes the table part and its autoFilter; the
    // guarantee is that everything else on the worksheet moves with it.
    const dataXml = await readPackagePart(
      await alphaOutput({ shape: "table", formulas: "structured" }),
      CORPUS_PARTS.dataSheet,
    );

    expect(conditionalFormattingSqref(dataXml)).toBe("D4:D6");
    expect(dataValidationSqref(dataXml)).toBe("A4:A6");
    expect(mergedCellReferences(dataXml)).toEqual(["A1:F1", "H5:I5"]);
    expect(hyperlinkReferences(dataXml)).toEqual(["A5"]);
  });

  it("Tier-1 fix: a cell comment and its VML shape follow the row they annotate", async () => {
    const alpha = await alphaOutput({ shape: "range" });

    // The note is anchored twice: an A1 reference in the comments part and a
    // zero-based row in the legacy VML drawing that draws its box. Data!B6 is
    // an Alpha record, so both have to land on row 5 (VML row 4).
    expect(await readPackagePart(alpha, CORPUS_PARTS.comments)).toContain(
      'ref="B5"',
    );
    expect(
      await readPackagePart(alpha, "xl/drawings/vmlDrawing1.vml"),
    ).toContain("<x:Row>4</x:Row>");
  });

  it("Tier-1 fix: the declared dimension shrinks to the surviving extent", async () => {
    // DECIDED IN PHASE 1: a stale `dimension/@ref` is a dependent reference
    // like any other, so the invariant pass rewrites it. The corpus sheet
    // declares A1:I12; Alpha keeps rows 1, 3, 4, 5 and 6 (H6:I6 having moved
    // to H5:I5), so the surviving extent is A1:I6. The table binding keeps its
    // footer block on row 12 and therefore keeps the declared extent.
    expect(
      await readPackagePart(
        await alphaOutput({ shape: "range" }),
        CORPUS_PARTS.dataSheet,
      ),
    ).toContain('<dimension ref="A1:I6"/>');
    expect(
      await readPackagePart(
        await alphaOutput({ shape: "table", formulas: "structured" }),
        CORPUS_PARTS.dataSheet,
      ),
    ).toContain('<dimension ref="A1:I12"/>');
  });
});
