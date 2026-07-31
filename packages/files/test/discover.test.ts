import {
  mkdir,
  mkdtemp,
  readFile,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, describe, expect, it } from "vitest";

import {
  discoverFiles,
  ensureDirectory,
  ensureOutputAvailable,
  ensureParentDirectory,
  FILES_ERRORS,
  pathExists,
  refuseInputOverwrite,
} from "../src/index.js";

const temporaryDirectories: string[] = [];

async function createTemporaryDirectory(): Promise<string> {
  const directory = await mkdtemp(
    path.join(tmpdir(), "consultchimps-files-test-"),
  );
  temporaryDirectories.push(directory);
  return directory;
}

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { force: true, recursive: true })),
  );
});

describe("discoverFiles", () => {
  it("resolves an absolute platform-native glob pattern", async () => {
    const directory = await createTemporaryDirectory();
    const first = path.join(directory, "first.xlsx");
    const second = path.join(directory, "second.xlsx");
    await writeFile(first, "first");
    await writeFile(second, "second");

    await expect(
      discoverFiles([path.join(directory, "*.xlsx")], {
        extensions: [".xlsx"],
      }),
    ).resolves.toEqual([first, second].sort());
  });

  it("expands directories recursively and deduplicates matches", async () => {
    const directory = await createTemporaryDirectory();
    const nested = path.join(directory, "nested");
    await mkdir(nested, { recursive: true });
    const top = path.join(directory, "top.xlsx");
    const deep = path.join(nested, "deep.xlsx");
    const ignored = path.join(nested, "notes.txt");
    await writeFile(top, "top");
    await writeFile(deep, "deep");
    await writeFile(ignored, "ignored");

    const discovered = await discoverFiles([directory, top], {
      extensions: [".xlsx"],
    });
    expect(discovered).toEqual([deep, top].sort());
  });

  it("normalizes extension filters with and without leading dots", async () => {
    const directory = await createTemporaryDirectory();
    const workbook = path.join(directory, "MIXED.XLSX");
    await writeFile(workbook, "workbook");

    await expect(
      discoverFiles([workbook], { extensions: ["xlsx"] }),
    ).resolves.toEqual([workbook]);
    await expect(
      discoverFiles([workbook], { extensions: [".XLSX"] }),
    ).resolves.toEqual([workbook]);
  });

  it("throws stable errors for empty and unmatched inputs", async () => {
    const directory = await createTemporaryDirectory();

    await expect(discoverFiles([])).rejects.toMatchObject({
      code: FILES_ERRORS.FILES_NO_INPUTS,
    });
    await expect(
      discoverFiles([path.join(directory, "missing-*.xlsx")]),
    ).rejects.toMatchObject({ code: FILES_ERRORS.FILES_NOT_FOUND });
    const stray = path.join(directory, "stray.txt");
    await writeFile(stray, "stray");
    await expect(
      discoverFiles([stray], { extensions: [".xlsx"] }),
    ).rejects.toMatchObject({ code: FILES_ERRORS.FILES_NOT_FOUND });
  });
});

describe("path helpers", () => {
  it("reports path existence without throwing", async () => {
    const directory = await createTemporaryDirectory();
    const existing = path.join(directory, "existing.txt");
    await writeFile(existing, "content");

    await expect(pathExists(existing)).resolves.toBe(true);
    await expect(pathExists(directory)).resolves.toBe(true);
    await expect(pathExists(path.join(directory, "missing.txt"))).resolves.toBe(
      false,
    );
  });

  it("creates directories and parent directories recursively", async () => {
    const directory = await createTemporaryDirectory();
    const nested = path.join(directory, "a", "b", "c");
    await expect(ensureDirectory(nested)).resolves.toBe(path.resolve(nested));
    expect((await stat(nested)).isDirectory()).toBe(true);

    const filePath = path.join(directory, "x", "y", "output.xlsx");
    await expect(ensureParentDirectory(filePath)).resolves.toBe(
      path.resolve(filePath),
    );
    expect((await stat(path.dirname(filePath))).isDirectory()).toBe(true);
    await expect(pathExists(filePath)).resolves.toBe(false);
  });

  it("guards existing outputs unless overwrite is enabled", async () => {
    const directory = await createTemporaryDirectory();
    const output = path.join(directory, "output.xlsx");

    await expect(ensureOutputAvailable(output)).resolves.toBe(
      path.resolve(output),
    );
    await writeFile(output, "existing");
    await expect(ensureOutputAvailable(output)).rejects.toMatchObject({
      code: FILES_ERRORS.FILES_OUTPUT_EXISTS,
    });
    await expect(
      ensureOutputAvailable(output, { overwrite: true }),
    ).resolves.toBe(path.resolve(output));
    expect(await readFile(output, "utf8")).toBe("existing");
  });

  it("refuses to use an input path as the output", () => {
    const input = path.resolve("clients.xlsx");
    expect(() => refuseInputOverwrite(input, [input])).toThrow(
      expect.objectContaining({ code: FILES_ERRORS.FILES_INPUT_OVERWRITE }),
    );
    expect(() =>
      refuseInputOverwrite(path.resolve("other.xlsx"), [input]),
    ).not.toThrow();
  });

  it.runIf(process.platform === "win32" || process.platform === "darwin")(
    "refuses an input overwrite when only path casing differs",
    () => {
      expect(() =>
        refuseInputOverwrite(path.resolve("CLIENTS.PPTX"), [
          path.resolve("clients.pptx"),
        ]),
      ).toThrow(
        expect.objectContaining({
          code: FILES_ERRORS.FILES_INPUT_OVERWRITE,
        }),
      );
    },
  );
});
