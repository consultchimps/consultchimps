/**
 * Conformance corpus: merge and consolidate.
 *
 * Phase 1b rebuilt the merge as a part-level transplant on the L0 package
 * (`src/merge/`). Every pin in this file that read "merge currently drops X"
 * has been flipped DELIBERATELY, one at a time, with the justification on the
 * assertion: the transplant copies worksheet parts with their dependents and
 * rewrites only the indexes and names a copy invalidates, so a structure now
 * survives unless carrying it would be wrong. What is still removed - pivot
 * caches, external links, the calculation chain, a macro project that cannot
 * travel - is pinned here as removal WITH its warning, and declared in
 * `src/contract.ts` under `merge`.
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
  CORPUS_NUMBER_FORMAT_CODE,
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
  styleNumberFormatCode,
  worksheetCellFormula,
  worksheetCellStyle,
  worksheetCellValue,
  writeCorpusWorkbook,
} from "./fixtures.js";
import { SHAPES } from "./symmetry.js";

/** The first input's Data worksheet keeps its part path: it seeds the output. */
const MERGED_DATA_PART = "xl/worksheets/sheet1.xml";
/** The second input's Data worksheet, transplanted after the seed's four. */
const TRANSPLANTED_DATA_PART = "xl/worksheets/sheet5.xml";
/** The second input's Summary worksheet. */
const TRANSPLANTED_SUMMARY_PART = "xl/worksheets/sheet6.xml";
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

  // FLIPPED (was "pins: merge currently drops merged ranges ... only on the
  // first input"): the transplant copies the second input's worksheet part, so
  // the dependent structures arrive on every copy, not only on the seed.
  it("invariant: a transplanted worksheet keeps its own dependents and comment part", async () => {
    const merged = await mergedCorpus();
    const transplanted = await readPackagePart(
      merged.bytes,
      TRANSPLANTED_DATA_PART,
    );

    expect(mergedCellReferences(transplanted)).toEqual(["A1:F1", "H6:I6"]);
    expect(hyperlinkReferences(transplanted)).toEqual(["A6"]);
    // The comment part is copied under a free path, and the worksheet's
    // relationship ids are renumbered to point at the copy.
    expect(await hasPackagePart(merged.bytes, "xl/comments2.xml")).toBe(true);
    expect(await readPackagePart(merged.bytes, "xl/comments2.xml")).toContain(
      CORPUS_COMMENT_TEXT,
    );
    expect(
      await hasPackagePart(merged.bytes, "xl/drawings/vmlDrawing2.vml"),
    ).toBe(true);
    const relationships = await readPackagePart(
      merged.bytes,
      "xl/worksheets/_rels/sheet5.xml.rels",
    );
    expect(relationships).toContain("../comments2.xml");
    expect(relationships).toContain("../drawings/vmlDrawing2.vml");
  });

  // FLIPPED (was "pins: merge currently drops workbook-level and sheet-level
  // defined names"): names are carried; a workbook-scoped collision is renamed
  // rather than silently dropped, and a sheet-scoped name follows its sheet.
  it("invariant: merge carries defined names and renames a workbook-scoped collision", async () => {
    const merged = await mergedCorpus();
    const workbookXml = await readPackagePart(
      merged.bytes,
      CORPUS_PARTS.workbook,
    );

    expect(workbookXml).toContain(
      `<definedName name="${CORPUS_RANGE_NAME}">Data!$A$3:$F$9</definedName>`,
    );
    // First wins; the later workbook's name takes a numeric suffix, and its
    // reference follows the worksheet that was renamed with it.
    expect(workbookXml).toContain(
      `<definedName name="${CORPUS_RANGE_NAME}_2">'Data (2)'!$A$3:$F$9</definedName>`,
    );
    expect(merged.warnings.join("\n")).toContain(
      `Renamed 1 defined name an earlier workbook already claimed: "${CORPUS_RANGE_NAME}" to "${CORPUS_RANGE_NAME}_2".`,
    );
    // Sheet-scoped names never collide: each one is scoped to its own sheet,
    // so both survive under the same name at different sheet indexes.
    expect(workbookXml).toContain(
      `<definedName name="${CORPUS_LOCAL_NAME}" localSheetId="0">Data!$A$12</definedName>`,
    );
    expect(workbookXml).toContain(
      `<definedName name="${CORPUS_LOCAL_NAME}" localSheetId="4">'Data (2)'!$A$12</definedName>`,
    );
  });

  // FLIPPED (was "pins: merge currently drops Excel Table parts"): table parts
  // are transplanted, and their workbook-unique names are re-established.
  it.each(SHAPES)(
    "invariant: merge carries Excel Table parts and renames a %s-shape name collision",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const first = await writeCorpusWorkbook(directory, "first.xlsx", {
        shape: "table",
      });
      const second = await writeCorpusWorkbook(directory, "second.xlsx", {
        shape: "table",
        formulas: shape === "table" ? "structured" : "a1",
      });
      const output = path.join(directory, "merged.xlsx");

      const result = await mergeWorkbooks([first, second], output);
      const bytes = await readWorkbookBytes(output);
      const parts = await packagePartNames(bytes);

      expect(parts).toContain("xl/tables/table1.xml");
      expect(parts).toContain("xl/tables/table2.xml");
      expect(await readPackagePart(bytes, MERGED_DATA_PART)).toContain(
        "tableParts",
      );
      const transplanted = await readPackagePart(bytes, TRANSPLANTED_DATA_PART);
      expect(transplanted).toContain("tableParts");

      // Table names are workbook-unique, so the second copy is renamed - with
      // a fresh table id, because those must be unique too.
      const copy = await readPackagePart(bytes, "xl/tables/table2.xml");
      expect(copy).toContain(`name="${CORPUS_TABLE_NAME}2"`);
      expect(copy).toContain(`displayName="${CORPUS_TABLE_NAME}2"`);
      expect(copy).toContain('id="2"');
      // The totals row is worksheet content plus one table attribute, and a
      // merge moves neither, so it arrives whole.
      expect(copy).toContain('totalsRowCount="1"');
      expect(worksheetCellValue(transplanted, "D10")).toBe("210");
      expect(result.warnings.join("\n")).toContain(
        `Renamed 1 Excel Table to keep table names unique in the merged workbook: "${CORPUS_TABLE_NAME}" to "${CORPUS_TABLE_NAME}2".`,
      );
      if (shape === "table") {
        // Structured references name the table, so they are rewritten with it.
        expect(worksheetCellFormula(transplanted, "E4")).toBe(
          `${CORPUS_TABLE_NAME}2[[#This Row],[Amount]]*2`,
        );
        expect(copy).toContain(`${CORPUS_TABLE_NAME}2[[#This Row],[Amount]]*2`);
      }
    },
  );

  // FLIPPED (was "pins: merge currently drops conditional formatting and data
  // validation"): both are worksheet-local, so a part-level copy keeps them.
  it.each(SHAPES)(
    "invariant: merge carries conditional formatting and data validation on a %s-shape input",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const first = await writeCorpusWorkbook(directory, "first.xlsx", {
        shape,
      });
      const second = await writeCorpusWorkbook(directory, "second.xlsx", {
        shape,
      });
      const output = path.join(directory, "merged.xlsx");

      await mergeWorkbooks([first, second], output);
      const bytes = await readWorkbookBytes(output);

      for (const part of [MERGED_DATA_PART, TRANSPLANTED_DATA_PART]) {
        const worksheet = await readPackagePart(bytes, part);
        expect(conditionalFormattingSqref(worksheet)).toBe("D4:D9");
        expect(dataValidationSqref(worksheet)).toBe("A4:A9");
        // The rule's differential format is remapped into the merged styles
        // part; both inputs declare the same one, so it dedupes to index 0.
        expect(worksheet).toContain('dxfId="0"');
      }
      expect(await readPackagePart(bytes, "xl/styles.xml")).toContain(
        "<dxfs count=",
      );
    },
  );

  // FLIPPED (was "pins: merge currently drops the calculation chain and the
  // shared string table"): one merged table with remapped indexes replaces the
  // old engine's inline re-emission. The chain stays dropped, deliberately.
  it("invariant: merge keeps one shared string table and remaps every transplanted index", async () => {
    const merged = await mergedCorpus();
    const parts = await packagePartNames(merged.bytes);

    expect(parts).toContain(CORPUS_PARTS.sharedStrings);
    const table = await readPackagePart(
      merged.bytes,
      CORPUS_PARTS.sharedStrings,
    );
    // Both inputs intern the same strings, so the merged table is the size of
    // one of them: identical items collapse rather than being appended.
    const items = [...table.matchAll(/<si>/gu)].length;
    expect(items).toBe(27);
    expect(table).toContain("<si><t>Client A</t></si>");

    // Cells still point INTO the table rather than carrying inline copies.
    const transplanted = await readPackagePart(
      merged.bytes,
      TRANSPLANTED_DATA_PART,
    );
    expect(transplanted).toContain('<c r="B4" t="s">');
    expect(transplanted).not.toContain("Client A");
    const index = Number(worksheetCellValue(transplanted, "B4"));
    expect(items).toBeGreaterThan(index);
    const item = [...table.matchAll(/<si>(?:(?!<\/si>)[\s\S])*<\/si>/gu)][
      index
    ]?.[0];
    expect(item).toBe("<si><t>Client A</t></si>");

    // The calculation chain is a derived index keyed by sheet id, and every
    // transplanted worksheet takes a new one. Dropping it is the fix: the
    // merged workbook asks Excel to recalculate instead.
    expect(parts).not.toContain(CORPUS_PARTS.calcChain);
    expect(
      await readPackagePart(merged.bytes, CORPUS_PARTS.workbook),
    ).toContain('fullCalcOnLoad="1"');
  });

  // NEW: the shared string table only dedupes when the merged indexes are
  // rewritten, so this pins the remap itself rather than the table's size.
  it("invariant: a transplanted string index resolves to the same text it did in its source", async () => {
    const directory = await createCorpusDirectory();
    // The second workbook's table starts with strings the first never uses, so
    // an unremapped index would resolve to the wrong text rather than to none.
    const first = await writeCorpusWorkbook(directory, "first.xlsx", {
      shape: "range",
      comments: false,
      dependents: false,
      hiddenSheets: false,
      summarySheet: false,
      footerBlock: false,
      formulas: "none",
      definedNames: false,
    });
    const second = await writeCorpusWorkbook(directory, "second.xlsx", {
      shape: "range",
    });
    const output = path.join(directory, "merged.xlsx");

    await mergeWorkbooks([first, second], output);
    const bytes = await readWorkbookBytes(output);
    const workbook = XLSX.read(await readFile(output), { type: "buffer" });

    expect(workbook.Sheets["Data (2)"]?.B4?.v).toBe("Client A");
    expect(workbook.Sheets["Data (2)"]?.A12?.v).toBe("Footer note");
    expect(workbook.Sheets["VeryHidden"]?.A1?.v).toBe("Archive note");
    // "Archive note" is only in the second workbook, so its index had to move.
    const seedTable = await readPackagePart(
      await buildCorpusWorkbook({
        shape: "range",
        comments: false,
        dependents: false,
        hiddenSheets: false,
        summarySheet: false,
        footerBlock: false,
        formulas: "none",
        definedNames: false,
      }),
      CORPUS_PARTS.sharedStrings,
    );
    expect(seedTable).not.toContain("Archive note");
    expect(await readPackagePart(bytes, CORPUS_PARTS.sharedStrings)).toContain(
      "Archive note",
    );
  });

  // NEW: styles are the other per-workbook index space, and the one the old
  // engine mangled most visibly - a number format is a workbook-level object a
  // cell only points at.
  it.each(SHAPES)(
    "invariant: merge copies a %s-shape input's number format with the cells that use it",
    async (shape) => {
      const directory = await createCorpusDirectory();
      // Only the SECOND workbook defines the format, so its style index cannot
      // survive unremapped: the seed has no entry at that index at all.
      const first = await writeCorpusWorkbook(directory, "first.xlsx", {
        shape,
      });
      const second = await writeCorpusWorkbook(directory, "second.xlsx", {
        shape,
        numberFormat: true,
      });
      const output = path.join(directory, "merged.xlsx");

      await mergeWorkbooks([first, second], output);
      const bytes = await readWorkbookBytes(output);
      const styles = await readPackagePart(bytes, "xl/styles.xml");
      const transplanted = await readPackagePart(bytes, TRANSPLANTED_DATA_PART);

      expect(styles).toContain("numFmt");
      const styleIndex = worksheetCellStyle(transplanted, "D4");
      expect(styleIndex).toBeGreaterThan(0);
      expect(styleNumberFormatCode(styles, styleIndex)).toBe(
        CORPUS_NUMBER_FORMAT_CODE,
      );
      // The seed's own cells keep the indexes they had, so its formatting is
      // untouched by what a later workbook contributed.
      const seed = await readPackagePart(bytes, MERGED_DATA_PART);
      expect(worksheetCellStyle(seed, "A3")).toBe(1);
      expect(worksheetCellStyle(seed, "D4")).toBe(0);
      // Reading through a spreadsheet library agrees with the parts.
      const workbook = XLSX.read(await readFile(output), {
        cellStyles: true,
        type: "buffer",
      });
      expect(workbook.Sheets["Data (2)"]?.D4?.z).toBe(
        CORPUS_NUMBER_FORMAT_CODE,
      );
    },
  );

  // NEW: two workbooks that define the SAME format must not grow the table.
  it("invariant: identical style entries from two inputs collapse into one", async () => {
    const directory = await createCorpusDirectory();
    const first = await writeCorpusWorkbook(directory, "first.xlsx", {
      shape: "range",
      numberFormat: true,
    });
    const second = await writeCorpusWorkbook(directory, "second.xlsx", {
      shape: "range",
      numberFormat: true,
    });
    const output = path.join(directory, "merged.xlsx");

    await mergeWorkbooks([first, second], output);
    const styles = await readPackagePart(
      await readWorkbookBytes(output),
      "xl/styles.xml",
    );

    expect([...styles.matchAll(/<numFmt\b/gu)]).toHaveLength(1);
    expect(/<cellXfs count="(\d+)"/u.exec(styles)?.[1]).toBe("3");
  });

  // STILL REMOVED, now with the warning the contract promises. A pivot cache
  // is a private copy of its source rows and its registry is workbook-scoped,
  // so it cannot be merged - the same Tier-1 policy the split applies.
  it("invariant: merge strips pivot tables and their caches, and says so", async () => {
    const merged = await mergedCorpus();
    const parts = await packagePartNames(merged.bytes);

    expect(parts.some((part) => part.startsWith("xl/pivotTables/"))).toBe(
      false,
    );
    expect(parts.some((part) => part.startsWith("xl/pivotCache/"))).toBe(false);
    expect(
      await readPackagePart(merged.bytes, CORPUS_PARTS.workbook),
    ).not.toContain("pivotCache");
    expect(merged.warnings.join("\n")).toMatch(/Removed 1 pivot table:/u);
  });

  // STILL REMOVED: a merge cannot renumber two workbooks' external references,
  // which formulas cite by position. Only the first input's macro project can
  // travel, and only into an output named as a macro-enabled workbook.
  it("invariant: a macro project travels only into an .xlsm output", async () => {
    const directory = await createCorpusDirectory();
    const macro = await writeCorpusWorkbook(directory, "macro.xlsm", {
      shape: "range",
      macro: true,
    });
    const plain = await writeCorpusWorkbook(directory, "plain.xlsx", {
      shape: "range",
    });

    const kept = path.join(directory, "merged.xlsm");
    const keptResult = await mergeWorkbooks([macro, plain], kept);
    const keptBytes = await readWorkbookBytes(kept);
    expect(keptResult.warnings.join("\n")).not.toContain("macro project");
    expect(await hasPackagePart(keptBytes, CORPUS_PARTS.vbaProject)).toBe(true);
    expect(
      await readPackagePart(keptBytes, CORPUS_PARTS.contentTypes),
    ).toContain("macroEnabled.main+xml");
    expect(keptResult.artifacts[0]?.mediaType).toBe(
      "application/vnd.ms-excel.sheet.macroEnabled.12",
    );

    const dropped = path.join(directory, "merged.xlsx");
    const droppedResult = await mergeWorkbooks([macro, plain], dropped);
    const droppedBytes = await readWorkbookBytes(dropped);
    expect(droppedResult.warnings.join("\n")).toContain(
      "Removed the macro project: name the merged workbook with an .xlsm extension to keep it.",
    );
    expect(await hasPackagePart(droppedBytes, CORPUS_PARTS.vbaProject)).toBe(
      false,
    );
    expect(
      await readPackagePart(droppedBytes, CORPUS_PARTS.contentTypes),
    ).not.toContain("macroEnabled.main+xml");
  });

  it("invariant: two macro projects cannot be combined, so neither travels", async () => {
    const directory = await createCorpusDirectory();
    const first = await writeCorpusWorkbook(directory, "first.xlsm", {
      shape: "range",
      macro: true,
    });
    const second = await writeCorpusWorkbook(directory, "second.xlsm", {
      shape: "range",
      macro: true,
    });
    const output = path.join(directory, "merged.xlsm");

    const result = await mergeWorkbooks([first, second], output);

    expect(result.warnings.join("\n")).toContain(
      "Removed the macro project: 2 inputs carried one and two VBA projects cannot be combined.",
    );
    expect(
      await hasPackagePart(
        await readWorkbookBytes(output),
        CORPUS_PARTS.vbaProject,
      ),
    ).toBe(false);
  });

  it("invariant: a values-only merge bakes cached values and keeps the formats around them", async () => {
    const directory = await createCorpusDirectory();
    const input = await writeCorpusWorkbook(directory, "first.xlsx", {
      shape: "range",
      numberFormat: true,
    });
    const output = path.join(directory, "merged.xlsx");

    await mergeWorkbooks([input], output, { values: true });

    const bytes = await readWorkbookBytes(output);
    const dataXml = await readPackagePart(bytes, MERGED_DATA_PART);
    expect(worksheetCellFormula(dataXml, "E4")).toBeUndefined();
    expect(worksheetCellValue(dataXml, "E4")).toBe("20");
    // Values mode converts the OUTPUT, so everything the transplant preserved
    // is still there afterwards - only formulas are gone.
    expect(conditionalFormattingSqref(dataXml)).toBe("D4:D9");
    expect(
      styleNumberFormatCode(
        await readPackagePart(bytes, "xl/styles.xml"),
        worksheetCellStyle(dataXml, "D4"),
      ),
    ).toBe(CORPUS_NUMBER_FORMAT_CODE);
  });

  it("invariant: identical inputs merge to identical bytes", async () => {
    const inputs = [
      {
        bytes: await buildCorpusWorkbook({ shape: "table" }),
        name: "first.xlsx",
      },
      {
        bytes: await buildCorpusWorkbook({ shape: "range" }),
        name: "second.xlsx",
      },
    ];

    const first = await mergeWorkbooksBytes({ inputs });
    const second = await mergeWorkbooksBytes({ inputs });

    expect(
      Buffer.compare(
        Buffer.from(first.outputs[0]!.bytes),
        Buffer.from(second.outputs[0]!.bytes),
      ),
    ).toBe(0);
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

  // NEW: a transplanted worksheet's cross-sheet formulas must follow the sheet
  // that was renamed under them, or they silently read another workbook's data.
  it("invariant: cross-sheet formulas follow a worksheet renamed by a collision", async () => {
    const merged = await mergedCorpus();
    const summary = await readPackagePart(
      merged.bytes,
      TRANSPLANTED_SUMMARY_PART,
    );

    expect(worksheetCellFormula(summary, "B2")).toBe("SUM('Data (2)'!D4:D9)");
    expect(worksheetCellFormula(summary, "B3")).toBe("'Data (2)'!D4");
    // The first workbook's Summary kept its name, so its formula is untouched.
    const seedSummary = await readPackagePart(
      merged.bytes,
      CORPUS_PARTS.summarySheet,
    );
    expect(worksheetCellFormula(seedSummary, "B2")).toBe("SUM(Data!D4:D9)");
  });

  // NEW: the formula flavours a spreadsheet library cannot round-trip, which
  // is exactly why the old engine lost them. A part-level copy keeps the
  // element text, so `si`, `ref` and a missing cached value all survive.
  it.each(SHAPES)(
    "invariant: merge carries shared, array and uncached formulas verbatim on a %s-shape input",
    async (shape) => {
      const directory = await createCorpusDirectory();
      const options = {
        arrayFormula: true,
        shape,
        sharedFormula: true,
        uncachedFormula: true,
      } as const;
      const first = await writeCorpusWorkbook(directory, "first.xlsx", options);
      const second = await writeCorpusWorkbook(
        directory,
        "second.xlsx",
        options,
      );
      const output = path.join(directory, "merged.xlsx");

      await mergeWorkbooks([first, second], output);
      const bytes = await readWorkbookBytes(output);
      const transplanted = await readPackagePart(bytes, TRANSPLANTED_DATA_PART);

      expect(transplanted).toContain('t="shared" ref="F4:F9" si="0"');
      expect(transplanted).toContain('t="array" ref="G4:G4"');
      // A formula with no cached value keeps having none: the merge never
      // invents a result it did not read.
      const summary = await readPackagePart(bytes, TRANSPLANTED_SUMMARY_PART);
      expect(worksheetCellFormula(summary, "B4")).toBe("SUM('Data (2)'!F4:F9)");
      expect(worksheetCellValue(summary, "B4")).toBeUndefined();
    },
  );
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
      // This pin runs without a column mapping, so both mapping counts are
      // zero. The contract's consolidate column is unaffected: a mapping
      // renames and folds columns of the values-only table this operation
      // already writes, and touches no tracked workbook structure.
      suggestedColumns: 0,
      unmappedColumns: 0,
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
