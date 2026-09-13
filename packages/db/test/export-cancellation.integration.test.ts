import { mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test, vi } from "vitest";

import { inspectDatabase } from "../src/database.js";
import { NodeDuckDbEngine } from "../src/engines/duckdb/node.js";
import { NodeSqliteEngine } from "../src/engines/sqlite/node.js";
import { NativeFileRegistry } from "../src/native-files.js";
import { createDatabase, exportDatabase } from "../src/node.js";

const directories: string[] = [];

afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

for (const format of ["sqlite", "duckdb"] as const) {
  for (const cancellation of ["copy-start", "copy-complete"] as const) {
    test(`${format}: cancellation at ${cancellation} does not publish an export`, async () => {
      const directory = await mkdtemp(
        path.join(tmpdir(), `cc-export-cancel-${format}-`),
      );
      directories.push(directory);
      const sourcePath = path.join(directory, `source.${format}`);
      const output = path.join(directory, `existing.${format}`);
      const priorOutput = Buffer.from("prior destination contents");
      await writeFile(output, priorOutput);
      const { database } = await createDatabase({
        path: sourcePath,
        format,
        schema: {
          version: 1,
          tables: [
            {
              name: "Records",
              recordId: { prefix: "REC", padding: 6 },
              columns: [{ name: "Value", type: "text" }],
            },
          ],
        },
      });
      const before = await inspectDatabase({ database });
      const controller = new AbortController();
      if (format === "sqlite") {
        const original = NodeSqliteEngine.prototype.backupTo;
        vi.spyOn(NodeSqliteEngine.prototype, "backupTo").mockImplementation(
          async function (this: NodeSqliteEngine, destination, signal) {
            if (cancellation === "copy-start") controller.abort(cancellation);
            await original.call(this, destination, signal);
            if (cancellation === "copy-complete")
              controller.abort(cancellation);
          },
        );
      } else {
        const original = NodeDuckDbEngine.prototype.copyTo;
        vi.spyOn(NodeDuckDbEngine.prototype, "copyTo").mockImplementation(
          async function (this: NodeDuckDbEngine, destination, signal) {
            if (cancellation === "copy-start") controller.abort(cancellation);
            await original.call(this, destination, signal);
            if (cancellation === "copy-complete")
              controller.abort(cancellation);
          },
        );
      }

      try {
        await expect(
          exportDatabase({
            database,
            output,
            overwrite: true,
            signal: controller.signal,
          }),
        ).rejects.toMatchObject({ code: "OPERATION_ABORTED" });
        expect(Buffer.from(await readFile(output)).equals(priorOutput)).toBe(
          true,
        );
        expect(await inspectDatabase({ database })).toEqual(before);
        expect(
          (await readdir(directory)).filter((name) =>
            name.includes(".cc-export-"),
          ),
        ).toEqual([]);
      } finally {
        await database.close();
      }
    });
  }
}

test("cancellation while waiting for publication keeps the staged file private", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-publish-cancel-"));
  directories.push(directory);
  const output = path.join(directory, "output.sqlite");
  const temporary = path.join(directory, ".output.sqlite.cc-export-test");
  const priorOutput = Buffer.from("prior destination contents");
  const stagedOutput = Buffer.from("completed private export");
  await writeFile(output, priorOutput);
  await writeFile(temporary, stagedOutput);
  const registry = new NativeFileRegistry();
  const plan = await registry.planPublication({
    output,
    inputs: [],
    overwrite: true,
  });
  let release!: () => void;
  let entered!: () => void;
  const inside = new Promise<void>((resolve) => {
    entered = resolve;
  });
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  const holding = registry.inspect(async () => {
    entered();
    await blocked;
  });
  await inside;
  const controller = new AbortController();
  const publishing = registry.publish({
    temporary,
    plan,
    signal: controller.signal,
  });
  controller.abort("cancelled while waiting");
  release();

  await expect(publishing).rejects.toMatchObject({
    code: "OPERATION_ABORTED",
  });
  await holding;
  expect(Buffer.from(await readFile(output)).equals(priorOutput)).toBe(true);
  expect(Buffer.from(await readFile(temporary)).equals(stagedOutput)).toBe(
    true,
  );
});
