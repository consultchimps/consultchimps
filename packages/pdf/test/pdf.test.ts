import { readFile, writeFile, mkdtemp, rm, readdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  isConsultChimpsError,
  OPERATION_ABORTED,
  type OperationProgress,
} from "@consultchimps/core";
import { PDFDocument } from "pdf-lib";
import { describe, expect, it } from "vitest";

import {
  mergePdfs,
  planMergePdfs,
  planSplitPdf,
  splitPdf,
} from "../src/index.js";

async function createSamplePdf(filePath: string, pages: number): Promise<void> {
  const document = await PDFDocument.create();
  for (let index = 0; index < pages; index += 1) {
    document.addPage([300 + index * 100, 400 + index * 100]);
  }
  await writeFile(filePath, await document.save());
}

describe("PDF operations", () => {
  it("splits and merges documents without changing page count", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-pdf-"));

    try {
      const input = path.join(directory, "source.pdf");
      const pagesDirectory = path.join(directory, "pages");
      const merged = path.join(directory, "merged.pdf");
      await createSamplePdf(input, 3);

      const splitResult = await splitPdf({
        input,
        outputDirectory: pagesDirectory,
      });
      expect(splitResult.metrics.pages).toBe(3);
      expect(splitResult.artifacts).toHaveLength(3);

      const mergeResult = await mergePdfs({
        inputs: splitResult.artifacts.map((artifact) => artifact.path),
        output: merged,
      });
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

  it("reports deterministic progress while splitting", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-pdf-"));

    try {
      const input = path.join(directory, "source.pdf");
      await createSamplePdf(input, 3);

      const events: OperationProgress[] = [];
      await splitPdf({
        input,
        outputDirectory: path.join(directory, "pages"),
        onProgress: (progress) => events.push(progress),
      });

      expect(events).toHaveLength(3);
      expect(events.map((event) => event.completed)).toEqual([1, 2, 3]);
      expect(
        events.every(
          (event) => event.total === 3 && event.stage === "writing-pages",
        ),
      ).toBe(true);
      expect(events[0]?.detail).toBe("source-page-001.pdf");
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("stops before writing anything when already aborted", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-pdf-"));

    try {
      const input = path.join(directory, "source.pdf");
      const pagesDirectory = path.join(directory, "pages");
      await createSamplePdf(input, 2);

      const controller = new AbortController();
      controller.abort();

      let thrown: unknown;
      try {
        await splitPdf({
          input,
          outputDirectory: pagesDirectory,
          signal: controller.signal,
        });
      } catch (error) {
        thrown = error;
      }

      expect(isConsultChimpsError(thrown)).toBe(true);
      expect((thrown as { code: string }).code).toBe(OPERATION_ABORTED);
      await expect(readdir(pagesDirectory)).rejects.toThrow();
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("plans a split without writing any file", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-pdf-"));

    try {
      const input = path.join(directory, "source.pdf");
      const pagesDirectory = path.join(directory, "pages");
      await createSamplePdf(input, 2);

      const plan = await planSplitPdf({
        input,
        outputDirectory: pagesDirectory,
      });

      expect(plan.operation).toBe("pdf.split");
      expect(plan.inputs).toEqual([path.resolve(input)]);
      expect(plan.outputs).toHaveLength(2);
      expect(plan.outputs.every((output) => output.exists === false)).toBe(
        true,
      );
      expect(plan.warnings).toEqual([]);
      await expect(readdir(pagesDirectory)).rejects.toThrow();
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("plans a merge and flags an existing output as a collision", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-pdf-"));

    try {
      const first = path.join(directory, "first.pdf");
      const second = path.join(directory, "second.pdf");
      const output = path.join(directory, "merged.pdf");
      await createSamplePdf(first, 1);
      await createSamplePdf(second, 1);
      await createSamplePdf(output, 1);

      const plan = await planMergePdfs({
        inputs: [first, second],
        output,
      });

      expect(plan.operation).toBe("pdf.merge");
      expect(plan.outputs).toEqual([
        {
          kind: "file",
          mediaType: "application/pdf",
          path: path.resolve(output),
          exists: true,
        },
      ]);
      expect(plan.warnings).toHaveLength(1);

      const overwritingPlan = await planMergePdfs({
        inputs: [first, second],
        output,
        overwrite: true,
      });
      expect(overwritingPlan.warnings).toEqual([]);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });
});
