import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { inspectDatabase } from "../src/database.js";
import { inspectImport } from "../src/import/operations.js";
import {
  PREPARED_BINDING_TABLE,
  PREPARED_CAPTURE_TABLE,
  PREPARED_ROW_TABLE,
  preparedEngineOf,
} from "../src/prepared.js";
import { createDatabase, createPreparedImport } from "../src/node.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

test("pages prepared rows with a source-row keyset cursor", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-import-page-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "workspace.sqlite"),
    format: "sqlite",
  });
  const baseline = await inspectDatabase({ database });
  const prepared = await createPreparedImport({
    path: path.join(directory, "review.data"),
    database,
    recipe: { version: 1, routes: [] },
    baselineRevision: baseline.revision,
  });
  try {
    const engine = preparedEngineOf(prepared);
    for (const [capture, source] of [
      ["CAPTURE-A", "A"],
      ["CAPTURE-B", "B"],
    ] as const) {
      await engine.execute(
        `INSERT INTO ${PREPARED_CAPTURE_TABLE} VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [
          capture,
          null,
          source,
          `${source}.xlsx`,
          "Sheet1",
          "Sheet1",
          "test-1",
          source.repeat(64),
          1n,
          0n,
          2n,
          '[{"name":"Value","type":"text"}]',
        ],
      );
      await engine.execute(
        `INSERT INTO ${PREPARED_BINDING_TABLE} VALUES (?, ?, ?, ?)`,
        [source, "Sheet1", capture, `${source}.xlsx`],
      );
      for (const sourceRow of [1n, 2n]) {
        await engine.execute(
          `INSERT INTO ${PREPARED_ROW_TABLE} VALUES (?, ?, ?)`,
          [
            capture,
            sourceRow,
            JSON.stringify({
              Value: { kind: "string", value: `${source}${sourceRow}` },
            }),
          ],
        );
      }
    }

    const first = await inspectImport({
      prepared,
      page: { limit: 2 },
    });
    expect(first.examples.map((example) => example.values["Value"])).toEqual([
      { kind: "string", value: "A1" },
      { kind: "string", value: "A2" },
    ]);
    expect(first.nextCursor).toBe('["A","Sheet1","2"]');

    const second = await inspectImport({
      prepared,
      page: { limit: 2, cursor: first.nextCursor },
    });
    expect(second.examples.map((example) => example.values["Value"])).toEqual([
      { kind: "string", value: "B1" },
      { kind: "string", value: "B2" },
    ]);
    expect(second.nextCursor).toBeUndefined();

    await expect(
      inspectImport({
        prepared,
        page: {
          limit: 1,
          source: "B",
          selection: "Sheet1",
          cursor: first.nextCursor,
        },
      }),
    ).rejects.toMatchObject({ code: "DB_INVALID_CURSOR" });
  } finally {
    await prepared.close();
    await database.close();
  }
});
