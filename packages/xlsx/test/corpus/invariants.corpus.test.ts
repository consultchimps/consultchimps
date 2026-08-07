/**
 * Conformance corpus: the invariants the split engine already guarantees.
 *
 * Every fix that landed as a local patch is restated here as a named,
 * corpus-wide invariant, so the layered rewrite has to keep it rather than
 * rediscover it. The names carry the originating commit where there is one.
 */
import { readFile } from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";
import { afterEach, describe, expect, it } from "vitest";

import { splitWorkbookByColumn } from "../../src/index.js";
import { convertWorkbookToValuesWithReport } from "../../src/values-only.js";
import {
  buildCorpusWorkbook,
  cleanupCorpusDirectories,
  CORPUS_PARTS,
  CORPUS_SPLIT_COLUMN,
  CORPUS_TABLE_NAME,
  createCorpusDirectory,
  hasPackagePart,
  packagePartNames,
  readPackagePart,
  readWorkbookBytes,
  worksheetCellFormula,
  worksheetCellValue,
  writeCorpusWorkbook,
  type CorpusShape,
} from "./fixtures.js";

const SHAPES: readonly CorpusShape[] = ["table", "range"];
// Phase 1: the comments part and its VML drawing left this list because a
// split now legitimately edits them - a note anchored on a record follows that
// record's row - so "unchanged" would no longer be the correct guarantee.
const UNFILTERED_PARTS = [
  CORPUS_PARTS.summarySheet,
  CORPUS_PARTS.veryHiddenSheet,
  CORPUS_PARTS.sharedStrings,
  "xl/styles.xml",
];
// A values-only conversion rewrites formulas and nothing else, so the comment
// anchors are still untouched on that path.
const VALUES_UNCHANGED_PARTS = [
  ...UNFILTERED_PARTS,
  CORPUS_PARTS.comments,
  "xl/drawings/vmlDrawing1.vml",
];

afterEach(cleanupCorpusDirectories);

describe("corpus: preservation invariants", () => {
  it.each(SHAPES)(
    "invariant: parts the %s split does not filter pass through byte-identical",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
        shape,
        pivot: true,
      });
      const sourceBytes = await readWorkbookBytes(input);
      await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, "out"),
      });
      const alpha = await readWorkbookBytes(
        path.join(directory, "out", "Alpha.xlsx"),
      );

      for (const part of UNFILTERED_PARTS) {
        expect(await readPackagePart(alpha, part)).toBe(
          await readPackagePart(sourceBytes, part),
        );
      }
      // The pivot parts used to belong in this list. They are now removed
      // rather than preserved, because a pivot cache holds a private copy of
      // every group's rows; tier1-gaps.corpus.test.ts owns that expectation.
      for (const part of [
        CORPUS_PARTS.pivotTable,
        CORPUS_PARTS.pivotCacheDefinition,
        CORPUS_PARTS.pivotCacheRecords,
      ]) {
        expect(await hasPackagePart(alpha, part)).toBe(false);
      }
    },
  );

  it.each(SHAPES)(
    "invariant: a %s split never modifies its input workbook",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
        shape,
      });
      const before = await readFile(input);

      await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, "out"),
      });

      expect(await readFile(input)).toEqual(before);
    },
  );

  it.each(SHAPES)(
    "invariant: splitting the same %s workbook twice produces identical part contents",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
        shape,
      });

      await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, "first"),
      });
      await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, "second"),
      });

      const first = await readWorkbookBytes(
        path.join(directory, "first", "Alpha.xlsx"),
      );
      const second = await readWorkbookBytes(
        path.join(directory, "second", "Alpha.xlsx"),
      );
      const parts = await packagePartNames(first);
      expect(await packagePartNames(second)).toEqual(parts);
      for (const part of parts.filter((name) => name.endsWith(".xml"))) {
        expect(await readPackagePart(second, part)).toBe(
          await readPackagePart(first, part),
        );
      }
    },
  );

  it("invariant: a preserved table split is byte-reproducible", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      formulas: "structured",
      shape: "table",
    });

    for (const run of ["first", "second"]) {
      await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, run),
        table: CORPUS_TABLE_NAME,
      });
    }

    // preserve-table-split.ts writes through package-zip.ts, which stamps a
    // fixed date on every rewritten part and never creates folder entries.
    expect(
      await readFile(path.join(directory, "first", "corpus-Alpha.xlsx")),
    ).toEqual(
      await readFile(path.join(directory, "second", "corpus-Alpha.xlsx")),
    );
  });

  it("invariant: the all-worksheet split writes through the deterministic package writer", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "range",
    });
    for (const run of ["first", "second"]) {
      await splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: path.join(directory, run),
      });
    }

    // Phase 1 replaced the pin that recorded the opposite. The old path wrote
    // replaced parts with JSZip's defaults, so a rewritten worksheet took the
    // current wall-clock time and folder entries appeared that the source
    // package never had; reproducibility held only within one DOS timestamp
    // tick. Every write now goes through L0: fixed stamps, no folder entries.
    const first = await readFile(path.join(directory, "first", "Alpha.xlsx"));
    const archive = await JSZip.loadAsync(first);
    for (const part of [CORPUS_PARTS.summarySheet, CORPUS_PARTS.dataSheet]) {
      expect(archive.file(part)!.date.getUTCFullYear()).toBe(1980);
    }
    expect(Object.values(archive.files).some((entry) => entry.dir)).toBe(false);
    expect(first).toEqual(
      await readFile(path.join(directory, "second", "Alpha.xlsx")),
    );
  });

  it.each(SHAPES)(
    "invariant: values-only conversion of a %s workbook rewrites only formula and calcChain metadata",
    async (shape) => {
      const source = await buildCorpusWorkbook({ shape, pivot: true });
      const conversion = await convertWorkbookToValuesWithReport(source);
      const converted = new Uint8Array(conversion.bytes);

      expect(conversion.formulasConverted).toBeGreaterThan(0);
      const sourceParts = await packagePartNames(source);
      const convertedParts = await packagePartNames(converted);
      // The calculation chain is the only part the conversion removes.
      expect(convertedParts).toEqual(
        sourceParts.filter((part) => part !== CORPUS_PARTS.calcChain),
      );
      for (const part of VALUES_UNCHANGED_PARTS.filter(
        (candidate) => candidate !== CORPUS_PARTS.summarySheet,
      )) {
        expect(await readPackagePart(converted, part)).toBe(
          await readPackagePart(source, part),
        );
      }
      const dataXml = await readPackagePart(converted, CORPUS_PARTS.dataSheet);
      expect(worksheetCellFormula(dataXml, "E4")).toBeUndefined();
      expect(worksheetCellValue(dataXml, "E4")).toBe("20");
      // Styles, widths, merges and dependent structures are untouched.
      expect(dataXml).toContain('<mergeCell ref="H6:I6"/>');
      expect(dataXml).toContain('<c r="A1" s="1" t="s">');
    },
  );

  it("invariant: values-only conversion strips table calculated-column formulas", async () => {
    const source = await buildCorpusWorkbook({
      shape: "table",
      formulas: "structured",
    });
    const conversion = await convertWorkbookToValuesWithReport(source);

    const tableXml = await readPackagePart(
      new Uint8Array(conversion.bytes),
      CORPUS_PARTS.table,
    );
    expect(tableXml).not.toContain("calculatedColumnFormula");
    expect(tableXml).toContain(`displayName="${CORPUS_TABLE_NAME}"`);
  });
});

describe("corpus: output-safety invariants", () => {
  it("invariant: existing outputs are refused unless overwrite is set", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "range",
    });
    const outputDirectory = path.join(directory, "out");
    await splitWorkbookByColumn({
      column: CORPUS_SPLIT_COLUMN,
      input,
      outputDirectory,
    });

    await expect(
      splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory,
      }),
    ).rejects.toThrowError(/Output already exists/u);
    await expect(
      splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory,
        overwrite: true,
      }),
    ).resolves.toMatchObject({ metrics: { outputFiles: 3 } });
  });

  it("invariant: an output path inside the input is refused", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "Alpha.xlsx", {
      shape: "range",
    });

    await expect(
      splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory: directory,
      }),
    ).rejects.toThrowError();
  });

  it("invariant: an aborted split writes nothing", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "range",
    });
    const outputDirectory = path.join(directory, "out");

    await expect(
      splitWorkbookByColumn({
        column: CORPUS_SPLIT_COLUMN,
        input,
        outputDirectory,
        signal: AbortSignal.abort(),
      }),
    ).rejects.toThrowError();
    await expect(
      readFile(path.join(outputDirectory, "Alpha.xlsx")),
    ).rejects.toMatchObject({ code: "ENOENT" });
  });

  it("pins: output names take a prefix from the input file except in all-worksheet mode", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "corpus.xlsx", {
      shape: "table",
      formulas: "structured",
    });

    const allWorksheets = await splitWorkbookByColumn({
      column: CORPUS_SPLIT_COLUMN,
      input,
      outputDirectory: path.join(directory, "all"),
    });
    const tableSelection = await splitWorkbookByColumn({
      column: CORPUS_SPLIT_COLUMN,
      input,
      outputDirectory: path.join(directory, "table"),
      table: CORPUS_TABLE_NAME,
    });
    const prefixed = await splitWorkbookByColumn({
      column: CORPUS_SPLIT_COLUMN,
      filenamePrefix: "region",
      input,
      outputDirectory: path.join(directory, "prefixed"),
    });

    expect(
      allWorksheets.artifacts.map((artifact) => path.basename(artifact.path)),
    ).toEqual(["Alpha.xlsx", "Beta.xlsx", "Gamma.xlsx"]);
    expect(
      tableSelection.artifacts.map((artifact) => path.basename(artifact.path)),
    ).toEqual(["corpus-Alpha.xlsx", "corpus-Beta.xlsx", "corpus-Gamma.xlsx"]);
    expect(
      prefixed.artifacts.map((artifact) => path.basename(artifact.path)),
    ).toEqual(["region-Alpha.xlsx", "region-Beta.xlsx", "region-Gamma.xlsx"]);
  });
});
