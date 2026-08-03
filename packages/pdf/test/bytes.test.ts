import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

import {
  isConsultChimpsError,
  OPERATION_ABORTED,
  type OperationProgress,
} from "@consultchimps/core";
import { PDFDocument } from "pdf-lib";
import { describe, expect, it } from "vitest";

import {
  mergePdfsBytes,
  planSplitPdfBytes,
  splitPdfBytes,
} from "../src/bytes.js";

async function samplePdfBytes(pages: number): Promise<Uint8Array> {
  const document = await PDFDocument.create();
  for (let index = 0; index < pages; index += 1) {
    document.addPage([300 + index * 100, 400 + index * 100]);
  }
  return document.save();
}

describe("byte-level PDF operations", () => {
  it("splits in memory and reports names instead of paths", async () => {
    const bytes = await samplePdfBytes(3);
    const { result, outputs } = await splitPdfBytes({
      input: { name: "quarterly report.pdf", bytes },
    });

    expect(result.metrics).toEqual({
      inputFiles: 1,
      outputFiles: 3,
      pages: 3,
    });
    expect(outputs.map((output) => output.name)).toEqual([
      "quarterly report-page-001.pdf",
      "quarterly report-page-002.pdf",
      "quarterly report-page-003.pdf",
    ]);
    expect(result.artifacts.map((artifact) => artifact.path)).toEqual(
      outputs.map((output) => output.name),
    );

    const firstPage = await PDFDocument.load(outputs[0]!.bytes);
    expect(firstPage.getPageCount()).toBe(1);
    expect(firstPage.getPage(0).getSize()).toEqual({
      height: 400,
      width: 300,
    });
  });

  it("sanitizes unsafe filename fragments", async () => {
    const bytes = await samplePdfBytes(1);
    const { outputs } = await splitPdfBytes({
      input: { name: "cli*ent<>secret.pdf", bytes },
    });
    expect(outputs[0]?.name).toBe("cli-ent-secret-page-001.pdf");

    const reserved = await splitPdfBytes({
      input: { name: "CON.pdf", bytes },
    });
    expect(reserved.outputs[0]?.name).toBe("_CON-page-001.pdf");

    const reservedMerge = await mergePdfsBytes({
      inputs: [{ name: "a.pdf", bytes }],
      outputName: "aux.pdf",
    });
    expect(reservedMerge.outputs[0]?.name).toBe("_aux.pdf");

    const control = await splitPdfBytes({
      input: { name: "client.pdf", bytes },
    });
    expect(control.outputs[0]?.name).toBe("client-page-001.pdf");

    const longStem = `${"a".repeat(300)}.pdf`;
    const bounded = await splitPdfBytes({
      input: { name: longStem, bytes },
    });
    expect(bounded.outputs[0]?.name).toBe(`${"a".repeat(80)}-page-001.pdf`);
    expect(bounded.outputs[0]!.name.length).toBeLessThanOrEqual(255);

    const emojiStem = `${"🙂".repeat(100)}.pdf`;
    const boundedBytes = await splitPdfBytes({
      input: { name: emojiStem, bytes },
    });
    expect(
      Buffer.byteLength(boundedBytes.outputs[0]!.name, "utf8"),
    ).toBeLessThanOrEqual(255);
    expect(boundedBytes.outputs[0]?.name.endsWith("-page-001.pdf")).toBe(true);
  });

  it("plans a split without producing any bytes", async () => {
    const bytes = await samplePdfBytes(2);
    const plan = await planSplitPdfBytes({
      input: { name: "report.pdf", bytes },
    });

    expect(plan.operation).toBe("pdf.split");
    expect(plan.inputs).toEqual(["report.pdf"]);
    expect(plan.outputs.map((output) => output.path)).toEqual([
      "report-page-001.pdf",
      "report-page-002.pdf",
    ]);
    expect(plan.outputs.every((output) => output.exists === false)).toBe(true);
    expect(plan.metrics).toEqual({ inputFiles: 1, outputFiles: 2, pages: 2 });
  });

  it("merges in memory with a derived output name and progress", async () => {
    const first = await samplePdfBytes(2);
    const second = await samplePdfBytes(1);
    const events: OperationProgress[] = [];

    const { result, outputs } = await mergePdfsBytes({
      inputs: [
        { name: "first.pdf", bytes: first },
        { name: "second.pdf", bytes: second },
      ],
      outputName: "client pack.pdf",
      onProgress: (progress) => events.push(progress),
    });

    expect(outputs).toHaveLength(1);
    expect(outputs[0]?.name).toBe("client pack.pdf");
    expect(result.metrics).toEqual({
      inputFiles: 2,
      outputFiles: 1,
      pages: 3,
    });
    expect(events.map((event) => [event.stage, event.completed])).toEqual([
      ["merging-inputs", 1],
      ["merging-inputs", 2],
    ]);

    const merged = await PDFDocument.load(outputs[0]!.bytes);
    expect(merged.getPageCount()).toBe(3);

    const defaulted = await mergePdfsBytes({
      inputs: [{ name: "first.pdf", bytes: first }],
    });
    expect(defaulted.outputs[0]?.name).toBe("combined.pdf");
  });

  it("produces byte-identical outputs for identical inputs", async () => {
    const bytes = await samplePdfBytes(2);
    const first = await splitPdfBytes({ input: { name: "a.pdf", bytes } });
    await new Promise((resolve) => setTimeout(resolve, 1100));
    const second = await splitPdfBytes({ input: { name: "a.pdf", bytes } });

    for (let index = 0; index < first.outputs.length; index += 1) {
      expect(
        Buffer.compare(
          Buffer.from(first.outputs[index]!.bytes),
          Buffer.from(second.outputs[index]!.bytes),
        ),
      ).toBe(0);
    }
  });

  it("honours cancellation and reports stable error codes", async () => {
    const bytes = await samplePdfBytes(1);
    const controller = new AbortController();
    controller.abort();

    let thrown: unknown;
    try {
      await splitPdfBytes({
        input: { name: "a.pdf", bytes },
        signal: controller.signal,
      });
    } catch (error) {
      thrown = error;
    }
    expect(isConsultChimpsError(thrown)).toBe(true);
    expect((thrown as { code: string }).code).toBe(OPERATION_ABORTED);
    expect((thrown as Error).message).toContain(
      "no partial output was produced",
    );
    expect((thrown as Error).message).not.toContain("output files");

    await expect(mergePdfsBytes({ inputs: [] })).rejects.toMatchObject({
      code: "PDF_NO_INPUTS",
    });
    await expect(
      splitPdfBytes({
        input: { name: "junk.pdf", bytes: new Uint8Array([1, 2, 3]) },
      }),
    ).rejects.toMatchObject({ code: "PDF_READ_FAILED" });
  });

  it("keeps the built bytes entry free of node imports", async () => {
    const distDirectory = new URL("../dist/", import.meta.url);
    const visited = new Set<string>();
    const queue = ["bytes.js"];

    while (queue.length > 0) {
      const entry = queue.pop()!;
      if (visited.has(entry)) {
        continue;
      }
      visited.add(entry);
      const source = await readFile(
        fileURLToPath(new URL(entry, distDirectory)),
        "utf8",
      );
      expect(source, `${entry} must not import node builtins`).not.toMatch(
        /["']node:/u,
      );
      for (const match of source.matchAll(/from\s+["']\.\/([^"']+)["']/gu)) {
        queue.push(match[1]!);
      }
    }

    expect(visited.size).toBeGreaterThan(0);
  });
});
