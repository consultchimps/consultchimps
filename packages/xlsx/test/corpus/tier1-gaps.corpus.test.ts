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
 */
import path from "node:path";

import { afterEach, describe, expect, it } from "vitest";

import { splitWorkbookByColumn } from "../../src/index.js";
import {
  calcChainReferences,
  cleanupCorpusDirectories,
  conditionalFormattingSqref,
  CORPUS_PARTS,
  CORPUS_SPLIT_COLUMN,
  createCorpusDirectory,
  dataValidationSqref,
  hyperlinkReferences,
  mergedCellReferences,
  readPackagePart,
  readWorkbookBytes,
  worksheetCellFormula,
  worksheetCellValue,
  writeCorpusWorkbook,
} from "./fixtures.js";

/** Alpha holds records 1, 3 and 6, so its per-group amount total is 100. */
const ALPHA_TOTAL = 100;

afterEach(cleanupCorpusDirectories);

/** Split the corpus by Group and return the Alpha output's bytes. */
async function alphaOutput(options: {
  shape: "table" | "range";
  formulas?: "a1" | "structured";
  pivot?: boolean;
  sharedFormula?: boolean;
  values?: boolean;
}): Promise<Uint8Array> {
  const directory = await createCorpusDirectory();
  const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
    formulas: options.formulas ?? "a1",
    pivot: options.pivot ?? false,
    shape: options.shape,
    sharedFormula: options.sharedFormula ?? false,
  });
  const outputDirectory = path.join(directory, "out");
  await splitWorkbookByColumn({
    column: CORPUS_SPLIT_COLUMN,
    input,
    outputDirectory,
    values: options.values ?? false,
  });
  return readWorkbookBytes(path.join(outputDirectory, "Alpha.xlsx"));
}

describe("corpus: Tier-1 gap - pivot caches cross split outputs", () => {
  it("pins: a split output currently carries the complete pivot cache", async () => {
    const records = await readPackagePart(
      await alphaOutput({ shape: "range", pivot: true }),
      CORPUS_PARTS.pivotCacheRecords,
    );

    expect(records).toContain('<s v="Beta"/>');
    expect(records).toContain('<s v="Client D"/>');
    expect(records).toContain('count="6"');
  });

  it.fails(
    "Tier-1 gap: pivot-cache records of another group must not reach a group's split output",
    async () => {
      const alpha = await alphaOutput({ shape: "range", pivot: true });
      const records = await readPackagePart(
        alpha,
        CORPUS_PARTS.pivotCacheRecords,
      );
      const definition = await readPackagePart(
        alpha,
        CORPUS_PARTS.pivotCacheDefinition,
      );

      // Confidentiality: the Alpha recipient must not be able to read Beta's
      // and Gamma's rows out of the pivot cache that travelled with the file.
      expect(records).not.toContain("Beta");
      expect(records).not.toContain("Gamma");
      expect(records).not.toContain("Client D");
      expect(definition).not.toContain("Beta");
    },
  );
});

describe("corpus: Tier-1 gap - stale cached aggregates in values mode", () => {
  it("pins: a values-only split currently bakes the all-rows total into every output", async () => {
    const summary = await readPackagePart(
      await alphaOutput({ shape: "range", values: true }),
      CORPUS_PARTS.summarySheet,
    );

    expect(worksheetCellFormula(summary, "B2")).toBeUndefined();
    // Summary!B2 is SUM(Data!D4:D9) with a cached 210 covering every group.
    expect(worksheetCellValue(summary, "B2")).toBe("210");
  });

  it.fails(
    "Tier-1 gap: a values-only split must not bake an aggregate computed over every group's rows",
    async () => {
      const summary = await readPackagePart(
        await alphaOutput({ shape: "range", values: true }),
        CORPUS_PARTS.summarySheet,
      );

      // The Alpha workbook contains only Alpha's rows, so a total presented as
      // "Total across every record" must equal Alpha's own total.
      expect(Number(worksheetCellValue(summary, "B2"))).toBe(ALPHA_TOTAL);
    },
  );
});

describe("corpus: Tier-1 gap - stale calculation chain", () => {
  it("pins: a non-values split currently leaves the calculation chain untouched", async () => {
    const calcChain = await readPackagePart(
      await alphaOutput({ shape: "range" }),
      CORPUS_PARTS.calcChain,
    );

    const references = calcChainReferences(calcChain);
    // Rows 7, 8 and 9 were deleted, and rows 10 and 12 no longer exist either.
    expect(references).toContain("E7");
    expect(references).toContain("E8");
    expect(references).toContain("E9");
    expect(references).toContain("B12");
  });

  it.fails(
    "Tier-1 gap: the calculation chain must not reference cells a split deleted",
    async () => {
      const calcChain = await readPackagePart(
        await alphaOutput({ shape: "range" }),
        CORPUS_PARTS.calcChain,
      );

      // Alpha keeps three data rows, so only E4, E5 and E6 remain calculable
      // on the Data worksheet.
      expect(
        calcChainReferences(calcChain).filter((reference) =>
          reference.startsWith("E"),
        ),
      ).toEqual(["E4", "E5", "E6"]);
    },
  );
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
