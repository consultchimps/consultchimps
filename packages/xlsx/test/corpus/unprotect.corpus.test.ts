/**
 * Conformance corpus: workbook unprotect (`sheets.unprotect`).
 *
 * Every cell of `CONTRACT.unprotect` reads `preserve`, and this file holds them
 * up. Unprotect removes only the worksheet (`sheetProtection`) and workbook
 * (`workbookProtection`) protection elements, so the proof is a per-part
 * comparison: build a workbook that carries every structure the generator can
 * emit, add protection to it, unprotect it, and assert that each output part
 * equals its input part with only the protection element removed. A structure
 * that changed would fail that comparison, so no cell here can be true while
 * another is false.
 *
 * Fixtures are the paired corpus workbooks, so a capability that worked for only
 * one shape is a failing test rather than a review comment.
 */
import JSZip from "jszip";
import { afterEach, describe, expect, it } from "vitest";

import { unprotectWorkbookBytes } from "../../src/bytes.js";
import {
  buildCorpusWorkbook,
  cleanupCorpusDirectories,
  CORPUS_PARTS,
  packagePartNames,
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

const WORKSHEET_PART = /^xl\/worksheets\/[^/]+\.xml$/u;
const SHEET_PROTECTION =
  '<sheetProtection sheet="1" objects="1" scenarios="1"/>';
const WORKBOOK_PROTECTION =
  '<workbookProtection lockStructure="1" workbookPassword="ABCD"/>';

interface ProtectedWorkbook {
  bytes: Uint8Array;
  worksheetParts: string[];
}

/**
 * Add worksheet and workbook protection to a corpus workbook. Every worksheet
 * part gains a self-closing `<sheetProtection/>`, and `xl/workbook.xml` gains a
 * `<workbookProtection/>`, exactly the empty elements unprotect removes.
 */
async function protectWorkbook(bytes: Uint8Array): Promise<ProtectedWorkbook> {
  const archive = await JSZip.loadAsync(bytes);
  const worksheetParts = Object.values(archive.files)
    .filter((entry) => !entry.dir && WORKSHEET_PART.test(entry.name))
    .map((entry) => entry.name)
    .sort();
  for (const part of worksheetParts) {
    const xml = await archive.file(part)!.async("text");
    archive.file(
      part,
      xml.replace("</worksheet>", `${SHEET_PROTECTION}</worksheet>`),
    );
  }
  const workbookXml = await archive.file(CORPUS_PARTS.workbook)!.async("text");
  archive.file(
    CORPUS_PARTS.workbook,
    workbookXml.replace("</workbook>", `${WORKBOOK_PROTECTION}</workbook>`),
  );
  return {
    bytes: await archive.generateAsync({
      compression: "DEFLATE",
      type: "uint8array",
    }),
    worksheetParts,
  };
}

/** The decoded contents of every non-directory part, keyed by part path. */
async function partContents(
  bytes: Uint8Array,
): Promise<Map<string, Uint8Array>> {
  const archive = await JSZip.loadAsync(bytes);
  const contents = new Map<string, Uint8Array>();
  for (const entry of Object.values(archive.files)) {
    if (entry.dir) {
      continue;
    }
    contents.set(entry.name, await entry.async("uint8array"));
  }
  return contents;
}

const decoder = new TextDecoder();

function decode(bytes: Uint8Array | undefined): string {
  return bytes === undefined ? "" : decoder.decode(bytes);
}

describe("corpus: unprotect preserves every tracked structure", () => {
  it.each(SHAPES)(
    "pins: every output part equals its input part with only the protection removed (%s)",
    async (shape) => {
      const source = await buildCorpusWorkbook(everyStructure(shape));
      const { bytes: protectedBytes, worksheetParts } =
        await protectWorkbook(source);
      const before = await partContents(protectedBytes);

      const outcome = await unprotectWorkbookBytes({
        input: { name: "corpus.xlsx", bytes: protectedBytes },
      });

      expect(outcome.result.metrics).toEqual({
        sheetProtectionsRemoved: worksheetParts.length,
        workbookProtectionsRemoved: 1,
      });

      const output = outcome.outputs[0]!.bytes;
      const after = await partContents(output);

      // No part is added or removed: unprotect only edits existing parts.
      expect([...after.keys()].sort()).toEqual([...before.keys()].sort());

      for (const [part, inputBytes] of before) {
        if (part === CORPUS_PARTS.workbook) {
          expect(decode(after.get(part))).toBe(
            decode(inputBytes).replace(WORKBOOK_PROTECTION, ""),
          );
        } else if (WORKSHEET_PART.test(part)) {
          expect(decode(after.get(part))).toBe(
            decode(inputBytes).replace(SHEET_PROTECTION, ""),
          );
        } else {
          // Every other tracked structure lives in a part unprotect never
          // touches: the table part, comment and drawing parts, pivot cache and
          // table, styles, shared strings, and the calculation chain all travel
          // byte-identical.
          expect(after.get(part)).toEqual(inputBytes);
        }
      }
    },
  );

  it.each(SHAPES)(
    "pins: the workbook and worksheets keep their structures after unprotect (%s)",
    async (shape) => {
      const source = await buildCorpusWorkbook(everyStructure(shape));
      const { bytes: protectedBytes } = await protectWorkbook(source);

      const outcome = await unprotectWorkbookBytes({
        input: { name: "corpus.xlsx", bytes: protectedBytes },
      });
      const after = await partContents(outcome.outputs[0]!.bytes);

      const workbook = decode(after.get(CORPUS_PARTS.workbook));
      expect(workbook).not.toContain("workbookProtection");
      // Defined names and the pivot cache registry stay in the workbook part.
      expect(workbook).toContain("definedName");
      expect(workbook).toContain("pivotCache");

      const dataSheet = decode(after.get(CORPUS_PARTS.dataSheet));
      expect(dataSheet).not.toContain("sheetProtection");
      // Merged cells, conditional formatting, data validation, and hyperlinks
      // all survive on the data worksheet.
      expect(dataSheet).toContain("mergeCell");
      expect(dataSheet).toContain("conditionalFormatting");
      expect(dataSheet).toContain("dataValidation");
      expect(dataSheet).toContain("hyperlink");
      if (shape === "table") {
        expect(dataSheet).toContain("tableParts");
      }
    },
  );

  it("pins: a macro workbook keeps its VBA project and macro content type", async () => {
    const source = await buildCorpusWorkbook({ shape: "table", macro: true });
    const { bytes: protectedBytes, worksheetParts } =
      await protectWorkbook(source);
    const before = await partContents(protectedBytes);

    const outcome = await unprotectWorkbookBytes({
      input: { name: "corpus.xlsm", bytes: protectedBytes },
    });

    expect(outcome.result.metrics).toEqual({
      sheetProtectionsRemoved: worksheetParts.length,
      workbookProtectionsRemoved: 1,
    });

    const after = await partContents(outcome.outputs[0]!.bytes);
    // The opaque VBA project travels byte-identical, and the package keeps its
    // macro-enabled content type, so the output is still a macro workbook.
    expect(after.get(CORPUS_PARTS.vbaProject)).toEqual(
      before.get(CORPUS_PARTS.vbaProject),
    );
    expect(decode(after.get(CORPUS_PARTS.contentTypes))).toContain(
      "macroEnabled.main+xml",
    );
  });

  it.each(SHAPES)(
    "invariant: unprotecting a %s workbook adds and removes no package part",
    async (shape) => {
      const source = await buildCorpusWorkbook(everyStructure(shape));
      const { bytes: protectedBytes } = await protectWorkbook(source);
      const before = await packagePartNames(protectedBytes);

      const outcome = await unprotectWorkbookBytes({
        input: { name: "corpus.xlsx", bytes: protectedBytes },
      });

      expect(await packagePartNames(outcome.outputs[0]!.bytes)).toEqual(before);
    },
  );
});
