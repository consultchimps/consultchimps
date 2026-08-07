/**
 * L0 unit tests: deterministic package load and save, part and relationship
 * access, and the path arithmetic OOXML relationships are resolved with.
 */
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import {
  joinPackagePath,
  normalizePackagePath,
  packagePartDirectory,
  packagePartName,
  relationshipsPartPath,
  resolveRelationshipTarget,
  WorkbookPackage,
} from "../src/package/index.js";
import {
  buildCorpusWorkbook,
  CORPUS_PARTS,
  CORPUS_VBA_BYTES,
} from "./corpus/fixtures.js";

async function corpusPackage(): Promise<{
  bytes: Uint8Array;
  workbookPackage: WorkbookPackage;
}> {
  const bytes = await buildCorpusWorkbook({ shape: "table", pivot: true });
  return { bytes, workbookPackage: await WorkbookPackage.load(bytes) };
}

describe("package: parts", () => {
  it("keeps every part, in source order, and no folder entries", async () => {
    const { bytes, workbookPackage } = await corpusPackage();
    const sourceArchive = await JSZip.loadAsync(bytes);
    const sourceNames = Object.values(sourceArchive.files)
      .filter((entry) => !entry.dir)
      .map((entry) => entry.name);

    expect(workbookPackage.partNames()).toEqual(sourceNames);

    const saved = await JSZip.loadAsync(await workbookPackage.save());
    expect(Object.values(saved.files).some((entry) => entry.dir)).toBe(false);
    expect(Object.values(saved.files).map((entry) => entry.name)).toEqual(
      sourceNames,
    );
  });

  it("reads text and binary parts and reports missing ones", async () => {
    const { workbookPackage } = await corpusPackage();

    expect(workbookPackage.has(CORPUS_PARTS.workbook)).toBe(true);
    expect(workbookPackage.has("xl/absent.xml")).toBe(false);
    expect(workbookPackage.readText("xl/absent.xml")).toBeUndefined();
    expect(workbookPackage.readBytes("xl/absent.xml")).toBeUndefined();
    expect(workbookPackage.requireText(CORPUS_PARTS.workbook)).toContain(
      "<workbook",
    );
    expect(() => workbookPackage.requireText("xl/absent.xml")).toThrowError(
      /part is missing/u,
    );
  });

  it("round-trips binary parts unchanged", async () => {
    const bytes = await buildCorpusWorkbook({ shape: "range", macro: true });
    const workbookPackage = await WorkbookPackage.load(bytes);
    workbookPackage.writeText(CORPUS_PARTS.dataSheet, "<worksheet/>");

    const saved = await JSZip.loadAsync(await workbookPackage.save());
    const vba = await saved.file(CORPUS_PARTS.vbaProject)!.async("uint8array");
    expect([...vba]).toEqual([...CORPUS_VBA_BYTES]);
  });

  it("writes, removes and lists parts by pattern", async () => {
    const { workbookPackage } = await corpusPackage();

    workbookPackage.writeText("xl/added.xml", "<added/>");
    expect(workbookPackage.readText("xl/added.xml")).toBe("<added/>");
    workbookPackage.remove("xl/added.xml");
    expect(workbookPackage.has("xl/added.xml")).toBe(false);

    expect(
      workbookPackage.partsMatching(/^xl\/worksheets\/[^/]+\.xml$/u),
    ).toHaveLength(4);
  });
});

describe("package: deterministic serialization", () => {
  it("produces identical bytes for identical edits", async () => {
    const bytes = await buildCorpusWorkbook({ shape: "range" });

    const runs: Uint8Array[] = [];
    for (let run = 0; run < 2; run += 1) {
      const workbookPackage = await WorkbookPackage.load(bytes);
      workbookPackage.writeText(CORPUS_PARTS.dataSheet, "<worksheet/>");
      runs.push(await workbookPackage.save());
    }

    expect([...runs[0]!]).toEqual([...runs[1]!]);
  });

  it("stamps a fixed date on written parts and leaves untouched ones alone", async () => {
    const { bytes, workbookPackage } = await corpusPackage();
    const sourceDate = (await JSZip.loadAsync(bytes)).file(
      CORPUS_PARTS.summarySheet,
    )!.date;
    workbookPackage.writeText(CORPUS_PARTS.dataSheet, "<worksheet/>");

    const saved = await JSZip.loadAsync(await workbookPackage.save());
    expect(saved.file(CORPUS_PARTS.dataSheet)!.date.getUTCFullYear()).toBe(
      1980,
    );
    expect(saved.file(CORPUS_PARTS.summarySheet)!.date).toEqual(sourceDate);
  });

  it("passes untouched parts through byte-identical", async () => {
    const { bytes, workbookPackage } = await corpusPackage();
    workbookPackage.writeText(CORPUS_PARTS.dataSheet, "<worksheet/>");
    const saved = await JSZip.loadAsync(await workbookPackage.save());
    const source = await JSZip.loadAsync(bytes);

    for (const part of [
      CORPUS_PARTS.summarySheet,
      CORPUS_PARTS.sharedStrings,
      CORPUS_PARTS.pivotCacheRecords,
    ]) {
      expect([...(await saved.file(part)!.async("uint8array"))]).toEqual([
        ...(await source.file(part)!.async("uint8array")),
      ]);
    }
  });
});

describe("package: relationships and content types", () => {
  it("reads a part's relationships and resolves their targets", async () => {
    const { workbookPackage } = await corpusPackage();
    const relationships = workbookPackage.relationshipsOf("xl/workbook.xml");

    expect(relationships.length).toBeGreaterThan(0);
    const worksheet = relationships.find((relationship) =>
      relationship.type.endsWith("/worksheet"),
    )!;
    expect(
      workbookPackage.resolvePart("xl/workbook.xml", worksheet.target),
    ).toMatch(/^xl\/worksheets\//u);
    // A part with no relationships part reports none rather than failing.
    expect(workbookPackage.relationshipsOf(CORPUS_PARTS.sharedStrings)).toEqual(
      [],
    );
  });

  it("removes a part together with its relationship and content type", async () => {
    const { workbookPackage } = await corpusPackage();
    workbookPackage.removePartAndReferences(
      CORPUS_PARTS.calcChain,
      "xl/workbook.xml",
    );

    expect(workbookPackage.has(CORPUS_PARTS.calcChain)).toBe(false);
    expect(
      workbookPackage.readText("xl/_rels/workbook.xml.rels"),
    ).not.toContain("calcChain");
    expect(workbookPackage.readText(CORPUS_PARTS.contentTypes)).not.toContain(
      "calcChain",
    );
  });

  it("drops one relationship by id", async () => {
    const { workbookPackage } = await corpusPackage();
    const relationships =
      await workbookPackage.relationships("xl/workbook.xml");
    const target = relationships.at(-1)!;

    workbookPackage.removeRelationship("xl/workbook.xml", target.id);

    expect(
      (await workbookPackage.relationships("xl/workbook.xml")).map(
        (relationship) => relationship.id,
      ),
    ).not.toContain(target.id);
    // A part that declares no relationships is a no-op, not a failure.
    expect(() =>
      workbookPackage.removeRelationship(CORPUS_PARTS.sharedStrings, "rId1"),
    ).not.toThrow();
  });

  it("marks external relationship targets", async () => {
    const { workbookPackage } = await corpusPackage();
    const hyperlink = (
      await workbookPackage.relationships(CORPUS_PARTS.dataSheet)
    ).find((relationship) => relationship.type.endsWith("/hyperlink"));

    expect(hyperlink?.targetMode).toBe("External");
  });

  it("exposes parts through the seam accessor", async () => {
    const { workbookPackage } = await corpusPackage();
    const part = workbookPackage.part(CORPUS_PARTS.workbook)!;

    expect(part.path).toBe(CORPUS_PARTS.workbook);
    expect(await part.text()).toContain("<workbook");
    expect((await part.bytes()).byteLength).toBeGreaterThan(0);
    expect(workbookPackage.part("xl/absent.xml")).toBeUndefined();
  });

  it("labels the source it was loaded from", async () => {
    const bytes = await buildCorpusWorkbook({ shape: "range" });

    expect((await WorkbookPackage.load(bytes)).sourceLabel).toBe("memory");
    const labelled = await WorkbookPackage.load(bytes, {
      sourceLabel: "corpus.xlsx",
    });
    expect(labelled.sourceLabel).toBe("corpus.xlsx");
    expect(() => labelled.requireText("xl/absent.xml")).toThrowError(
      /corpus\.xlsx/u,
    );
  });

  it("removing an absent part changes nothing", async () => {
    const { workbookPackage } = await corpusPackage();
    const before = workbookPackage.readText(CORPUS_PARTS.contentTypes);

    workbookPackage.removePartAndReferences("xl/absent.xml", "xl/workbook.xml");

    expect(workbookPackage.readText(CORPUS_PARTS.contentTypes)).toBe(before);
  });
});

describe("package: part paths", () => {
  it("splits and joins POSIX part paths", () => {
    expect(packagePartDirectory("xl/worksheets/sheet1.xml")).toBe(
      "xl/worksheets",
    );
    expect(packagePartDirectory("[Content_Types].xml")).toBe("");
    expect(packagePartName("xl/worksheets/sheet1.xml")).toBe("sheet1.xml");
    expect(packagePartName("[Content_Types].xml")).toBe("[Content_Types].xml");
    expect(joinPackagePath("xl", "", "tables", "table1.xml")).toBe(
      "xl/tables/table1.xml",
    );
  });

  it("normalizes dot segments and keeps escapes visible", () => {
    expect(normalizePackagePath("xl/worksheets/../tables/table1.xml")).toBe(
      "xl/tables/table1.xml",
    );
    expect(normalizePackagePath("./xl/./workbook.xml")).toBe("xl/workbook.xml");
    expect(normalizePackagePath("")).toBe(".");
    expect(normalizePackagePath("/xl/../../etc")).toBe("/etc");
    expect(normalizePackagePath("../../outside")).toBe("../../outside");
  });

  it("resolves relationship targets and refuses ones that escape", () => {
    expect(
      resolveRelationshipTarget("xl/worksheets/sheet1.xml", "../tables/t.xml"),
    ).toBe("xl/tables/t.xml");
    expect(resolveRelationshipTarget("xl/workbook.xml", "/xl/styles.xml")).toBe(
      "xl/styles.xml",
    );
    expect(() =>
      resolveRelationshipTarget("xl/workbook.xml", "../../secrets.xml"),
    ).toThrowError(/escapes the workbook package/u);
  });

  it("names the relationships part of a part", () => {
    expect(relationshipsPartPath("xl/workbook.xml")).toBe(
      "xl/_rels/workbook.xml.rels",
    );
    expect(relationshipsPartPath("xl/worksheets/sheet1.xml")).toBe(
      "xl/worksheets/_rels/sheet1.xml.rels",
    );
  });
});
