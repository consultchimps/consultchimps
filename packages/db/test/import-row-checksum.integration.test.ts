import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { engineOf, inspectDatabase } from "../src/database.js";
import type { EngineTransaction } from "../src/internal/engine.js";
import {
  applyImport,
  prepareImport,
  resolveImport,
} from "../src/import/operations.js";
import type { ImportRecipe, ImportSource } from "../src/import/types.js";
import { CAPTURE_ROW_TABLE, DATABASE_METADATA_TABLE } from "../src/metadata.js";
import {
  PREPARED_BINDING_TABLE,
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

function source(options: {
  readonly key: string;
  readonly selection?: string | undefined;
  readonly value: string;
  readonly content?: string | undefined;
}): ImportSource {
  const selection = options.selection ?? "Data";
  const bytes = new TextEncoder().encode(options.content ?? options.value);
  return {
    key: options.key,
    readerVersion: "checksum-test-1",
    bytes: {
      name: `${options.key}.xlsx`,
      size: bytes.byteLength,
      async readAt(offset, length) {
        return bytes.slice(offset, offset + length);
      },
    },
    selections: [
      {
        key: selection,
        label: selection,
        async open() {
          return {
            columns: ["Value"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: {
                    Value: { kind: "string" as const, value: options.value },
                  },
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

function route(
  sourceKey: string,
  table: string,
): ImportRecipe["routes"][number] {
  return {
    source: sourceKey,
    selection: "Data",
    destination: {
      kind: "new-table",
      schema: {
        name: table,
        columns: [{ name: "Value", type: "text" }],
        recordId: { prefix: table.toUpperCase(), padding: 4 },
      },
    },
    columns: [{ source: "Value", target: "Value", type: "text" }],
  };
}

function emptySource(): ImportSource {
  const bytes = new TextEncoder().encode("empty captured selection");
  return {
    key: "omitted",
    readerVersion: "checksum-test-1",
    bytes: {
      name: "empty.xlsx",
      size: bytes.byteLength,
      async readAt(offset, length) {
        return bytes.slice(offset, offset + length);
      },
    },
    selections: [
      {
        key: "Empty",
        label: "Empty",
        async open() {
          return {
            columns: ["Value"],
            async *batches() {
              yield [];
            },
            async close() {},
          };
        },
      },
    ],
  };
}

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format}: an edited excluded capture rolls back earlier aliased capture work`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-row-seal-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const recipe: ImportRecipe = {
      version: 1,
      routes: [route("alias-a", "Included"), route("alias-b", "Included")],
    };
    const prepared = await createPreparedImport({
      path: path.join(directory, "review.ccplan"),
      database,
      recipe,
      baselineRevision: 0n,
    });
    try {
      const outcome = await prepareImport({
        database,
        prepared,
        recipe,
        sources: [
          source({
            key: "alias-a",
            value: "included",
            content: "identical aliased source",
          }),
          source({
            key: "alias-b",
            value: "included",
            content: "identical aliased source",
          }),
          source({
            key: "omitted",
            selection: "Omitted",
            value: "excluded",
          }),
        ],
      });
      expect(outcome.prepared.state).toBe("needs-review");
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [
          {
            kind: "exclude",
            source: "omitted",
            selection: "Omitted",
            reason: "Outside this import",
          },
        ],
      });
      if (approved.state !== "ready") {
        throw new Error("Excluded plan did not become ready");
      }
      const aliasBindings = await preparedEngineOf(prepared).query(
        `SELECT capture_id FROM ${PREPARED_BINDING_TABLE} WHERE source_key IN (?, ?) ORDER BY source_key`,
        ["alias-a", "alias-b"],
      );
      expect(aliasBindings).toHaveLength(2);
      expect(aliasBindings[0]?.["capture_id"]).toBe(
        aliasBindings[1]?.["capture_id"],
      );
      await preparedEngineOf(prepared).execute(
        `UPDATE ${PREPARED_ROW_TABLE} SET values_json = ? WHERE capture_id = (SELECT capture_id FROM ${PREPARED_BINDING_TABLE} WHERE source_key = ?)`,
        [
          JSON.stringify({
            Value: { kind: "string", value: "changed after review" },
          }),
          "omitted",
        ],
      );

      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: `${format}-excluded-checksum`,
        }),
      ).rejects.toMatchObject({ code: "DB_INVALID_PREPARED_IMPORT" });
      await expect(inspectDatabase({ database })).resolves.toMatchObject({
        revision: 0n,
        tables: [],
        captures: 0n,
        completedImports: 0n,
        deliveries: 0n,
      });
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: an edited reused capture blocks a new application without changing history`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-reused-row-seal-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const importSource = source({ key: "submission", value: "reviewed" });
    const firstRecipe: ImportRecipe = {
      version: 1,
      routes: [route("submission", "First")],
    };
    const firstPlan = await createPreparedImport({
      path: path.join(directory, "first.ccplan"),
      database,
      recipe: firstRecipe,
      baselineRevision: 0n,
    });
    const secondRecipe: ImportRecipe = {
      version: 1,
      routes: [route("submission", "Second")],
    };
    let secondPlan:
      Awaited<ReturnType<typeof createPreparedImport>> | undefined;
    try {
      const first = await prepareImport({
        database,
        prepared: firstPlan,
        recipe: firstRecipe,
        sources: [importSource],
      });
      if (first.prepared.state !== "ready") {
        throw new Error("First plan did not become ready");
      }
      await applyImport({
        database,
        prepared: firstPlan,
        approved: first.prepared,
        requestId: `${format}-first-capture`,
      });
      const baseline = await inspectDatabase({ database });
      secondPlan = await createPreparedImport({
        path: path.join(directory, "second.ccplan"),
        database,
        recipe: secondRecipe,
        baselineRevision: baseline.revision,
      });
      const second = await prepareImport({
        database,
        prepared: secondPlan,
        recipe: secondRecipe,
        sources: [importSource],
      });
      if (second.prepared.state !== "ready") {
        throw new Error("Reused plan did not become ready");
      }
      expect(second.result.metrics).toMatchObject({
        sourcesRead: 0,
        sourcesReused: 1,
        rowsCaptured: 0,
      });
      await engineOf(database).execute(
        `UPDATE ${CAPTURE_ROW_TABLE} SET values_json = ?`,
        [
          JSON.stringify({
            Value: { kind: "string", value: "changed after preparation" },
          }),
        ],
      );
      const before = await inspectDatabase({ database });

      await expect(
        applyImport({
          database,
          prepared: secondPlan,
          approved: second.prepared,
          requestId: `${format}-reused-checksum`,
        }),
      ).rejects.toMatchObject({ code: "DB_CORRUPT_DATABASE" });
      expect(await inspectDatabase({ database })).toEqual(before);
      expect(before).toMatchObject({
        revision: 1n,
        tables: [{ name: "First", rowCount: 1n }],
        captures: 1n,
        completedImports: 1n,
      });
    } finally {
      await Promise.all([firstPlan.close(), secondPlan?.close()]);
      await database.close();
    }
  });

  test(`${format}: cancellation after an empty capture page prevents publication`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-empty-abort-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const recipe: ImportRecipe = { version: 1, routes: [] };
    const prepared = await createPreparedImport({
      path: path.join(directory, "empty.ccplan"),
      database,
      recipe,
      baselineRevision: 0n,
    });
    const controller = new AbortController();
    try {
      await prepareImport({
        database,
        prepared,
        recipe,
        sources: [emptySource()],
      });
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [
          {
            kind: "exclude",
            source: "omitted",
            selection: "Empty",
            reason: "Outside this import",
          },
        ],
      });
      if (approved.state !== "ready") {
        throw new Error("Empty plan did not become ready");
      }
      const engine = preparedEngineOf(prepared);
      const originalQuery = engine.query.bind(engine);
      engine.query = async (sql, values) => {
        const rows = await originalQuery(sql, values);
        if (sql.includes(PREPARED_ROW_TABLE)) controller.abort();
        return rows;
      };
      try {
        await expect(
          applyImport({
            database,
            prepared,
            approved,
            requestId: `${format}-empty-abort`,
            signal: controller.signal,
          }),
        ).rejects.toMatchObject({ code: "OPERATION_ABORTED" });
      } finally {
        engine.query = originalQuery;
      }
      await expect(inspectDatabase({ database })).resolves.toMatchObject({
        revision: 0n,
        tables: [],
        captures: 0n,
        completedImports: 0n,
      });
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: cancellation at the commit boundary rolls back receipt and rows`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-commit-abort-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const recipe: ImportRecipe = {
      version: 1,
      routes: [route("submission", "Imported")],
    };
    const prepared = await createPreparedImport({
      path: path.join(directory, "review.ccplan"),
      database,
      recipe,
      baselineRevision: 0n,
    });
    const controller = new AbortController();
    const target = engineOf(database);
    const originalTransaction = target.transaction.bind(target);
    try {
      const outcome = await prepareImport({
        database,
        prepared,
        recipe,
        sources: [source({ key: "submission", value: "reviewed" })],
      });
      if (outcome.prepared.state !== "ready") {
        throw new Error("Plan did not become ready");
      }
      target.transaction = async <T>(
        work: (transaction: EngineTransaction) => Promise<T>,
      ) =>
        originalTransaction((transaction) =>
          work({
            ...transaction,
            async execute(sql, values) {
              await transaction.execute(sql, values);
              if (
                sql.includes(DATABASE_METADATA_TABLE) &&
                sql.includes("revision = revision + 1")
              ) {
                controller.abort();
              }
            },
          }),
        );
      try {
        await expect(
          applyImport({
            database,
            prepared,
            approved: outcome.prepared,
            requestId: `${format}-commit-abort`,
            signal: controller.signal,
          }),
        ).rejects.toMatchObject({ code: "OPERATION_ABORTED" });
      } finally {
        target.transaction = originalTransaction;
      }
      await expect(inspectDatabase({ database })).resolves.toMatchObject({
        revision: 0n,
        tables: [],
        captures: 0n,
        completedImports: 0n,
      });
    } finally {
      target.transaction = originalTransaction;
      await prepared.close();
      await database.close();
    }
  });
}
