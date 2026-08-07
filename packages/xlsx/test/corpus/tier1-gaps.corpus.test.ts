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
 * The first three gaps -- pivot caches crossing outputs, stale cached
 * aggregates in values mode, and the stale calculation chain -- are now closed
 * by the Tier-1 utilities in `src/tier1/`, so their `.fails` twins have become
 * `Tier-1 fix: ...` tests and their pins record the new output. The
 * dependent-reference gaps below are still open.
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

describe("corpus: Tier-1 gap - dependent references after row compaction", () => {
  /**
   * INVESTIGATION RESULT: broken, on both bindings.
   *
   * 128a310 taught `removeWorksheetRows` to renumber the `r` attribute of each
   * surviving `<row>` and of every `<c>` inside it. Nothing else in the
   * worksheet part is adjusted: `mergeCells`, `conditionalFormatting/@sqref`,
   * `dataValidation/@sqref`, `hyperlink/@ref`, shared-formula `@ref` and the
   * text inside `<f>` all keep their original row numbers while the cells they
   * describe move up. `filterPlainWorksheets` is the only editor on this path,
   * and it rewrites nothing outside `<sheetData>`.
   *
   * The Excel Table path is no safer. `preserve-table-split.ts` guards A1
   * formulas with `assertRelocatableFormula` (the XLSX_SPLIT_PRESERVE_FORMULA
   * refusal), but that guard only covers formulas inside relocated cells, and
   * `workbook-column-split.ts` sidesteps it entirely: `tableCanBeCompacted`
   * routes any sheet holding an A1 formula onto the plain path instead of
   * refusing. Merged ranges and sqrefs are never checked on either path.
   */

  it("pins: dependent references currently keep their original row numbers", async () => {
    const dataXml = await readPackagePart(
      await alphaOutput({ shape: "range", sharedFormula: true }),
      CORPUS_PARTS.dataSheet,
    );

    // Row 6 (an Alpha record) moved to row 5; row 9 moved to row 6.
    expect(worksheetCellValue(dataXml, "D5")).toBe("30");
    expect(worksheetCellValue(dataXml, "D6")).toBe("60");
    // Everything that referred to those rows still refers to the old numbers.
    expect(mergedCellReferences(dataXml)).toEqual(["A1:F1", "H6:I6"]);
    expect(conditionalFormattingSqref(dataXml)).toBe("D4:D9");
    expect(dataValidationSqref(dataXml)).toBe("A4:A9");
    expect(hyperlinkReferences(dataXml)).toEqual(["A6"]);
    expect(worksheetCellFormula(dataXml, "E5")).toBe("D6*2");
    expect(dataXml).toContain('ref="F4:F9"');
  });

  it.fails(
    "Tier-1 gap: merged ranges must follow the rows they cover",
    async () => {
      const dataXml = await readPackagePart(
        await alphaOutput({ shape: "range" }),
        CORPUS_PARTS.dataSheet,
      );

      // H6:I6 covered the record that now lives on row 5.
      expect(mergedCellReferences(dataXml)).toEqual(["A1:F1", "H5:I5"]);
    },
  );

  it.fails(
    "Tier-1 gap: conditional-formatting sqref must shrink with the rows it covers",
    async () => {
      const dataXml = await readPackagePart(
        await alphaOutput({ shape: "range" }),
        CORPUS_PARTS.dataSheet,
      );

      expect(conditionalFormattingSqref(dataXml)).toBe("D4:D6");
    },
  );

  it.fails(
    "Tier-1 gap: data-validation sqref must shrink with the rows it covers",
    async () => {
      const dataXml = await readPackagePart(
        await alphaOutput({ shape: "range" }),
        CORPUS_PARTS.dataSheet,
      );

      expect(dataValidationSqref(dataXml)).toBe("A4:A6");
    },
  );

  it.fails(
    "Tier-1 gap: hyperlinks must follow the row they decorate",
    async () => {
      const dataXml = await readPackagePart(
        await alphaOutput({ shape: "range" }),
        CORPUS_PARTS.dataSheet,
      );

      expect(hyperlinkReferences(dataXml)).toEqual(["A5"]);
    },
  );

  it.fails(
    "Tier-1 gap: A1 formulas on a plain worksheet must be rewritten when their row moves",
    async () => {
      const dataXml = await readPackagePart(
        await alphaOutput({ shape: "range" }),
        CORPUS_PARTS.dataSheet,
      );

      // E5 doubles the amount on its own row. After compaction that amount is
      // D5, but the formula still reads D6 and therefore doubles a different
      // record; had row 9 been dropped instead it would read a deleted cell.
      expect(worksheetCellFormula(dataXml, "E5")).toBe("D5*2");
      expect(worksheetCellFormula(dataXml, "E6")).toBe("D6*2");
    },
  );

  it.fails(
    "Tier-1 gap: shared-formula ranges must shrink with the rows they span",
    async () => {
      const dataXml = await readPackagePart(
        await alphaOutput({ shape: "range", sharedFormula: true }),
        CORPUS_PARTS.dataSheet,
      );

      expect(dataXml).toContain('ref="F4:F6"');
      expect(dataXml).not.toContain('ref="F4:F9"');
    },
  );

  it.fails(
    "Tier-1 gap: the Excel Table binding must fix dependent references too",
    async () => {
      // The compacting table path resizes the table part and its autoFilter,
      // which proves it knows the new geometry, yet it leaves every other
      // dependent reference on the worksheet alone.
      const dataXml = await readPackagePart(
        await alphaOutput({ shape: "table", formulas: "structured" }),
        CORPUS_PARTS.dataSheet,
      );

      expect(conditionalFormattingSqref(dataXml)).toBe("D4:D6");
      expect(dataValidationSqref(dataXml)).toBe("A4:A6");
      expect(mergedCellReferences(dataXml)).toEqual(["A1:F1", "H5:I5"]);
      expect(hyperlinkReferences(dataXml)).toEqual(["A5"]);
    },
  );
});
