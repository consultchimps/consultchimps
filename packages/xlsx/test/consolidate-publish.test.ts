/**
 * Publishing a consolidated workbook without overwrite on a filesystem that
 * refuses hard links: the fallback must still refuse to replace a destination
 * that appeared while the output was being written.
 */
import { writeFileSync } from "node:fs";
import { mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import type * as FsPromises from "node:fs/promises";

import { afterEach, describe, expect, it, vi } from "vitest";
import * as XLSX from "xlsx";

const linkBehaviour = vi.hoisted(() => ({
  /** Called with the destination before the link fails. */
  beforeFailing: undefined as ((destination: string) => void) | undefined,
}));

vi.mock("node:fs/promises", async (importOriginal) => {
  const actual = await importOriginal<typeof FsPromises>();
  return {
    ...actual,
    link: (_existing: string, destination: string) => {
      linkBehaviour.beforeFailing?.(destination);
      return Promise.reject(
        Object.assign(new Error("hard links are not supported"), {
          code: "EPERM",
        }),
      );
    },
  };
});

const { consolidateWorkbooks } = await import("../src/index.js");

async function inputFile(directory: string): Promise<string> {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Client", "Amount"],
      ["A", 10],
    ]),
    "North",
  );
  const input = path.join(directory, "north.xlsx");
  await writeFile(
    input,
    XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
  );
  return input;
}

afterEach(() => {
  linkBehaviour.beforeFailing = undefined;
});

describe("consolidation publishes without hard links", () => {
  it("writes the output by copying when nothing stands in the way", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));
    try {
      const output = path.join(directory, "consolidated.xlsx");
      await consolidateWorkbooks({
        inputs: [await inputFile(directory)],
        output,
      });
      expect((await readFile(output)).length).toBeGreaterThan(0);
      expect((await readdir(directory)).sort()).toEqual([
        "consolidated.xlsx",
        "north.xlsx",
      ]);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("refuses, and keeps, a destination created while the output was written", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));
    try {
      const output = path.join(directory, "consolidated.xlsx");
      linkBehaviour.beforeFailing = (destination) => {
        writeFileSync(destination, "someone else's file");
      };
      await expect(
        consolidateWorkbooks({ inputs: [await inputFile(directory)], output }),
      ).rejects.toMatchObject({ code: "FILES_OUTPUT_EXISTS" });
      expect(await readFile(output, "utf8")).toBe("someone else's file");
      expect((await readdir(directory)).sort()).toEqual([
        "consolidated.xlsx",
        "north.xlsx",
      ]);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });
});
