import {
  mkdtemp,
  mkdir,
  readFile,
  rename,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { createDatabase, openPreparedImport } from "../src/node.js";

const directories: string[] = [];
afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

const cases = ([false, true] as const).flatMap((readonly) =>
  (
    ["missing", "missing-parent", "directory", "invalid", "duckdb"] as const
  ).map((kind) => ({ readonly, kind })),
);

test.each(cases)(
  "rejects $kind prepared input with readonly=$readonly using a stable error",
  async ({ readonly, kind }) => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-prepared-open-"));
    directories.push(directory);
    const filePath =
      kind === "missing-parent"
        ? path.join(directory, "absent", "review.ccplan")
        : path.join(directory, "review.ccplan");
    if (kind === "directory") await mkdir(filePath);
    if (kind === "invalid")
      await writeFile(filePath, "Synthetic non-database input");
    if (kind === "duckdb") {
      const { database } = await createDatabase({
        path: filePath,
        format: "duckdb",
      });
      await database.close();
    }
    const before =
      kind === "invalid" || kind === "duckdb"
        ? await readFile(filePath)
        : undefined;
    await expect(
      openPreparedImport({ path: filePath, readonly }),
    ).rejects.toMatchObject({
      code: "DB_INVALID_PREPARED_IMPORT",
      message: expect.stringMatching(/import plan/u),
    });
    if (before !== undefined)
      expect((await readFile(filePath)).equals(before)).toBe(true);
    if (kind === "missing" || kind === "missing-parent") {
      await expect(stat(filePath)).rejects.toMatchObject({ code: "ENOENT" });
    } else {
      const moved = `${filePath}.moved`;
      await rename(filePath, moved);
      await rename(moved, filePath);
    }
  },
);
