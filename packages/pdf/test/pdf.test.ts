import { readFile, writeFile, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { PDFDocument } from "pdf-lib";
import { describe, expect, it } from "vitest";

import { mergePdfs, splitPdf } from "../src/index.js";

describe("PDF operations", () => {
  it("splits and merges documents without changing page count", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "chimpcons-pdf-"));

    try {
      const input = path.join(directory, "source.pdf");
      const pagesDirectory = path.join(directory, "pages");
      const merged = path.join(directory, "merged.pdf");
      const document = await PDFDocument.create();
      document.addPage([300, 400]);
      document.addPage([400, 500]);
      document.addPage([500, 600]);
      await writeFile(input, await document.save());

      const splitResult = await splitPdf(input, pagesDirectory);
      expect(splitResult.metrics.pages).toBe(3);
      expect(splitResult.artifacts).toHaveLength(3);

      const mergeResult = await mergePdfs(
        splitResult.artifacts.map((artifact) => artifact.path),
        merged,
      );
      expect(mergeResult.metrics.pages).toBe(3);

      const mergedDocument = await PDFDocument.load(await readFile(merged));
      expect(mergedDocument.getPageCount()).toBe(3);
      expect(mergedDocument.getPage(0).getSize()).toEqual({
        height: 400,
        width: 300,
      });
      expect(mergedDocument.getPage(2).getSize()).toEqual({
        height: 600,
        width: 500,
      });
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });
});
