import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { discoverFiles, refuseInputOverwrite } from "../src/index.js";

describe("discoverFiles", () => {
  it("resolves an absolute platform-native glob pattern", async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "consultchimps-files-test-"),
    );

    try {
      const first = path.join(directory, "first.xlsx");
      const second = path.join(directory, "second.xlsx");
      await writeFile(first, "first");
      await writeFile(second, "second");

      await expect(
        discoverFiles([path.join(directory, "*.xlsx")], {
          extensions: [".xlsx"],
        }),
      ).resolves.toEqual([first, second].sort());
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
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
          code: "FILES_INPUT_OVERWRITE",
        }),
      );
    },
  );
});
