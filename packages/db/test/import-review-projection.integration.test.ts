import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test, vi } from "vitest";

import {
  applyImport,
  draftImportProfile,
  inspectImport,
  prepareImport,
  resolveImport,
  type ImportSource,
} from "../src/index.js";
import {
  createDatabase,
  createImportBatch,
  openImportBatch,
} from "../src/node.js";
import { engineOf } from "../src/database.js";
import { preparedEngineOf } from "../src/prepared.js";
import { inspectUpdatedImport } from "../src/import/inspection.js";

const directories: string[] = [];

afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

function source(selections: number): ImportSource {
  const bytes = new TextEncoder().encode(`synthetic projection ${selections}`);
  return {
    key: "stable-alias",
    readerVersion: "projection-test-1",
    bytes: {
      name: "Friendly workbook.xlsx",
      size: bytes.length,
      async readAt(offset, length) {
        return bytes.slice(offset, offset + length);
      },
    },
    selections: Array.from({ length: selections }, (_, index) => {
      const key = `Region_${String(index).padStart(3, "0")}`;
      return {
        key,
        label: key,
        async open() {
          return {
            columns: ["Value"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: {
                    Value: { kind: "number" as const, raw: String(index) },
                  },
                },
              ];
            },
            async close() {},
          };
        },
      };
    }),
  };
}

test.each(["sqlite", "duckdb"] as const)(
  "%s returns the saved review and applied membership without source files",
  async (format) => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "cc-review-projection-"),
    );
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `working.${format}`),
      format,
    });
    const input = source(2);
    const profile = await draftImportProfile({
      sources: [input],
      naming: { kind: "selection-label" },
    });
    const batchPath = path.join(directory, "review.ccplan");
    let prepared = await createImportBatch({
      path: batchPath,
      database,
      profile,
      baselineRevision: 0n,
    });
    try {
      const captured = await prepareImport({
        database,
        prepared,
        sources: [input],
        profile,
        reviewPage: { limit: 1 },
      });
      expect(captured.prepared).toBe(captured.inspection.prepared);
      expect(captured.inspection.application).toEqual({ state: "pending" });
      expect(captured.inspection.reviewRows).toBe(2n);
      expect(
        captured.inspection.routes.map((route) => route.displayName),
      ).toEqual(["Friendly workbook.xlsx", "Friendly workbook.xlsx"]);
      const reviewed = await resolveImport({
        database,
        prepared,
        decisions: [
          {
            kind: "exclude",
            source: input.key,
            selection: "Region_001",
            reason: "synthetic exclusion",
          },
        ],
        reviewPage: { limit: 1 },
      });
      expect(reviewed.prepared).toBe(reviewed.inspection.prepared);
      const excluded = reviewed.inspection.routes.find(
        (route) => route.selection === "Region_001",
      );
      expect(excluded).toMatchObject({
        destination: null,
        suggestedDestination: { kind: "new-table-infer", name: "Region_001" },
      });
      expect(reviewed.prepared.state).toBe("ready");
      if (reviewed.prepared.state !== "ready")
        throw new Error("Expected a ready batch");
      const applied = await applyImport({
        database,
        prepared,
        approved: reviewed.prepared,
        requestId: "projection-apply",
      });
      await prepared.close();
      prepared = await openImportBatch({ path: batchPath });
      const reopened = await inspectImport({
        database,
        prepared,
        page: { limit: 1 },
      });
      expect(reopened.prepared).toEqual(reviewed.prepared);
      expect(reopened.application).toEqual({
        state: "applied",
        captureIds: applied.captureIds,
      });
      expect(
        reopened.application.state === "applied" &&
          reopened.application.captureIds.length,
      ).toBe(2);
      expect(reopened.routes).toHaveLength(2);
      expect(
        reopened.routes.find((route) => route.selection === "Region_000")
          ?.applicationState,
      ).toBe("already-applied");
      expect(reopened.reviewRows).toBe(2n);
      expect(reopened.targetRevision).not.toBe(
        reviewed.inspection.targetRevision,
      );
    } finally {
      await prepared.close();
      await database.close();
    }
  },
);

test("route pages are bounded, independent from previews, and reject a cursor after review changes", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-route-pages-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "working.sqlite"),
    format: "sqlite",
  });
  const input = source(53);
  const profile = await draftImportProfile({ sources: [input] });
  const prepared = await createImportBatch({
    path: path.join(directory, "review.ccplan"),
    database,
    profile,
    baselineRevision: 0n,
  });
  try {
    await prepareImport({ database, prepared, sources: [input], profile });
    const first = await inspectImport({
      database,
      prepared,
      page: { limit: 1 },
    });
    expect(first.routes).toHaveLength(50);
    expect(first.routeCount).toBe(53n);
    expect(first.reviewRows).toBe(53n);
    expect(first.examples).toHaveLength(1);
    expect(first.nextRouteCursor).toBeTypeOf("string");
    const next = await inspectImport({
      database,
      prepared,
      page: { limit: 1 },
      routePage: { cursor: first.nextRouteCursor },
    });
    expect(next.routes.map((route) => route.selection)).toEqual([
      "Region_050",
      "Region_051",
      "Region_052",
    ]);
    expect(next.nextRouteCursor).toBeUndefined();
    expect(next.examples).toEqual(first.examples);
    const review = await resolveImport({
      database,
      prepared,
      decisions: [],
      reviewPage: { limit: 1 },
    });
    expect(review.prepared.reviewFingerprint).not.toBe(
      first.prepared.reviewFingerprint,
    );
    await expect(
      inspectImport({
        database,
        prepared,
        page: { limit: 1 },
        routePage: { cursor: first.nextRouteCursor },
      }),
    ).rejects.toMatchObject({ code: "DB_INVALID_CURSOR" });
    await expect(
      inspectImport({
        database,
        prepared,
        page: { limit: 1 },
        routePage: { limit: 101 },
      }),
    ).rejects.toMatchObject({ code: "DB_INVALID_PAGE_SIZE" });
    if (review.prepared.state !== "ready")
      throw new Error("Expected ready batch");
    await applyImport({
      database,
      prepared,
      approved: review.prepared,
      requestId: "paged-application",
    });
    const queries = vi.spyOn(engineOf(database), "query");
    const applied = await inspectImport({
      database,
      prepared,
      page: { limit: 1 },
    });
    expect(applied.routes).toHaveLength(50);
    expect(
      applied.routes.every(
        (route) => route.applicationState === "already-applied",
      ),
    ).toBe(true);
    const applicationLookups = queries.mock.calls.filter(([sql]) =>
      sql.includes("SELECT capture_id, table_name, import_id, application_key"),
    );
    expect(applicationLookups).toHaveLength(1);
    expect(applicationLookups[0]?.[1]).toHaveLength(100);
  } finally {
    await prepared.close();
    await database.close();
  }
});

test("counts physical captures once while retaining distinct source bindings", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-review-alias-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "working.sqlite"),
    format: "sqlite",
  });
  const first = source(1);
  const second: ImportSource = {
    ...first,
    key: "second-alias",
    bytes: { ...first.bytes, name: "Second filename.xlsx" },
  };
  const sources = [first, second];
  const profile = await draftImportProfile({ sources });
  const prepared = await createImportBatch({
    path: path.join(directory, "review.ccplan"),
    database,
    profile,
    baselineRevision: 0n,
  });
  try {
    const result = await prepareImport({
      database,
      prepared,
      sources,
      profile,
      reviewPage: { limit: 1 },
    });
    expect(result.inspection.capturedRows).toBe(1n);
    expect(result.inspection.reviewRows).toBe(1n);
    expect(result.inspection.routeCount).toBe(2n);
    expect(
      result.inspection.routes.map((route) => route.displayName).sort(),
    ).toEqual(["Friendly workbook.xlsx", "Second filename.xlsx"]);
  } finally {
    await prepared.close();
    await database.close();
  }
});

test("a failed review refresh preserves the updated batch and explains the safe retry", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-review-refresh-"));
  directories.push(directory);
  const { database } = await createDatabase({
    path: path.join(directory, "working.sqlite"),
    format: "sqlite",
  });
  const input = source(1);
  const profile = await draftImportProfile({ sources: [input] });
  const prepared = await createImportBatch({
    path: path.join(directory, "review.ccplan"),
    database,
    profile,
    baselineRevision: 0n,
  });
  try {
    const captured = await prepareImport({
      database,
      prepared,
      sources: [input],
      profile,
    });
    const failure = new Error("synthetic metadata read failure");
    const reading = vi
      .spyOn(preparedEngineOf(prepared), "readTransaction")
      .mockRejectedValueOnce(failure);
    await expect(
      inspectUpdatedImport({
        database,
        prepared,
        expected: captured.prepared,
        page: { limit: 1 },
      }),
    ).rejects.toMatchObject({
      code: "DB_BATCH_REVIEW_REFRESH_REQUIRED",
      details: { batchUpdated: true, batchId: captured.prepared.id },
      cause: failure,
    });
    reading.mockRestore();
    expect(
      (await inspectImport({ database, prepared, page: { limit: 1 } }))
        .prepared,
    ).toEqual(captured.prepared);
  } finally {
    await prepared.close();
    await database.close();
  }
});
