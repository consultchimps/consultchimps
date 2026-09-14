import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { inspectDatabase } from "../src/database.js";
import { NodeSqliteEngine } from "../src/engines/sqlite/node.js";
import { prepareImport } from "../src/import/operations.js";
import type { ImportProfile, ImportSource } from "../src/import/types.js";
import {
  createDatabase,
  createImportBatch,
  openImportBatch,
} from "../src/node.js";
import { PREPARED_ROW_TABLE } from "../src/prepared.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

const profile: ImportProfile = {
  version: 1,
  routes: [
    {
      source: "source",
      selection: "Data",
      columns: [{ source: "Value", target: "Value", type: "text" }],
      destination: {
        kind: "new-table",
        schema: {
          name: "Rows",
          columns: [{ name: "Value", type: "text", nullable: true }],
          recordId: { prefix: "ROW", separator: "-", padding: 6 },
        },
      },
    },
  ],
};

function source(): ImportSource {
  const bytes = new TextEncoder().encode("synthetic prepared source");
  return {
    key: "source",
    readerVersion: "orphan-row-test-1",
    bytes: {
      name: "source.xlsx",
      size: bytes.byteLength,
      async readAt(offset, length) {
        return bytes.slice(offset, offset + length);
      },
    },
    selections: [
      {
        key: "Data",
        label: "Data",
        async open() {
          return {
            columns: ["Value"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: { Value: { kind: "string", value: "kept" } },
                },
              ];
            },
            async close() {},
          };
        },
      },
    ],
  };
}

test.each(["sqlite", "duckdb"] as const)(
  "%s rejects crash-leftover prepared rows without changing the plan",
  async (format) => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-plan-orphan-"));
    directories.push(directory);
    const databasePath = path.join(directory, `workspace.${format}`);
    const planPath = path.join(directory, "review.ccplan");
    const { database } = await createDatabase({ path: databasePath, format });
    const prepared = await createImportBatch({
      path: planPath,
      database,
      profile,
      baselineRevision: (await inspectDatabase({ database })).revision,
    });
    try {
      const outcome = await prepareImport({
        database,
        prepared,
        profile,
        sources: [source()],
      });
      expect(outcome.prepared.state).toBe("ready");
    } finally {
      await prepared.close();
      await database.close();
    }

    const valid = await openImportBatch({ path: planPath, readonly: true });
    await valid.close();

    const damaged = NodeSqliteEngine.open(planPath);
    await damaged.execute(
      `INSERT INTO ${PREPARED_ROW_TABLE} VALUES (?, ?, ?)`,
      [
        "ZZZ-ORPHAN",
        3n,
        JSON.stringify({ Value: { kind: "string", value: "hidden" } }),
      ],
    );
    await damaged.close();
    const before = await readFile(planPath);

    for (const readonly of [false, true]) {
      await expect(
        openImportBatch({ path: planPath, readonly }),
      ).rejects.toMatchObject({
        code: "DB_INVALID_PREPARED_IMPORT",
        message: expect.stringContaining(
          "captured rows without their capture definition",
        ),
        details: { captureId: "ZZZ-ORPHAN" },
      });
      expect((await readFile(planPath)).equals(before)).toBe(true);
    }
  },
);
