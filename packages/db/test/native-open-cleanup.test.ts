import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import SqliteDatabase from "better-sqlite3";
import { expect, test, vi } from "vitest";

async function fixture() {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-open-cleanup-"));
  const filename = path.join(directory, "damaged.sqlite");
  const database = new SqliteDatabase(filename);
  database.pragma("journal_mode = WAL");
  database.exec(
    "CREATE TABLE _consultchimps_database (bad TEXT); CREATE TABLE _consultchimps_prepared (bad TEXT)",
  );
  database.close();
  return { directory, filename };
}

test("inspection retains a managed owner when only final closure fails", async () => {
  vi.resetModules();
  const { NodeSqliteEngine } = await import("../src/engines/sqlite/node.js");
  const api = await import("../src/node.js");
  const directory = await mkdtemp(path.join(tmpdir(), "cc-inspect-cleanup-"));
  const filename = path.join(directory, "database.sqlite");
  const created = await api.createDatabase({
    path: filename,
    format: "sqlite",
  });
  await created.database.close();
  const failure = new Error("Synthetic inspection close failure");
  const owners: ReturnType<typeof NodeSqliteEngine.open>[] = [];
  const close = vi
    .spyOn(NodeSqliteEngine.prototype, "close")
    .mockImplementation(async function (
      this: ReturnType<typeof NodeSqliteEngine.open>,
    ) {
      owners.push(this);
      throw failure;
    });
  try {
    await expect(api.inspectFileKind({ path: filename })).rejects.toMatchObject(
      {
        code: "DB_NATIVE_HANDLE_CLEANUP_REQUIRED",
        cause: expect.objectContaining({ errors: [failure] }),
      },
    );
    expect(close).toHaveBeenCalledTimes(1);
    await expect(
      api.createDatabase({ path: filename, format: "sqlite", overwrite: true }),
    ).rejects.toMatchObject({ code: "DB_NATIVE_HANDLE_CLEANUP_REQUIRED" });
  } finally {
    close.mockRestore();
    for (const owner of owners) await owner.close();
    await rm(directory, { recursive: true, force: true });
  }
});

for (const operation of ["database", "prepared", "inspect"] as const) {
  test.each([new Error("Synthetic close failure"), undefined])(
    `${operation} retains validation and close failures and blocks replacement (%s)`,
    async (cleanupFailure) => {
      vi.resetModules();
      const { NodeSqliteEngine } =
        await import("../src/engines/sqlite/node.js");
      const api = await import("../src/node.js");
      const { directory, filename } = await fixture();
      const owners: ReturnType<typeof NodeSqliteEngine.open>[] = [];
      const close = vi
        .spyOn(NodeSqliteEngine.prototype, "close")
        .mockImplementation(async function (
          this: ReturnType<typeof NodeSqliteEngine.open>,
        ) {
          owners.push(this);
          throw cleanupFailure;
        });
      try {
        const open =
          operation === "database"
            ? api.openDatabase
            : operation === "prepared"
              ? api.openPreparedImport
              : api.inspectFileKind;
        const failure: unknown = await open({ path: filename }).catch(
          (error: unknown) => error,
        );
        expect(failure).toMatchObject({
          code: "DB_NATIVE_HANDLE_CLEANUP_REQUIRED",
          details: { path: filename, closeFailed: true },
        });
        if (
          !(failure instanceof Error) ||
          !(failure.cause instanceof AggregateError)
        )
          throw new Error("Expected combined cleanup error");
        expect(failure.cause.errors).toHaveLength(2);
        expect(failure.cause.errors[0]).toMatchObject({
          code: expect.any(String),
        });
        expect(failure.cause.errors[1]).toBe(cleanupFailure);
        expect(close).toHaveBeenCalledTimes(1);
        const before = await readFile(filename);
        await expect(
          api.createDatabase({
            path: filename,
            format: "sqlite",
            overwrite: true,
          }),
        ).rejects.toMatchObject({ code: "DB_NATIVE_HANDLE_CLEANUP_REQUIRED" });
        expect(await readFile(filename)).toEqual(before);
      } finally {
        close.mockRestore();
        for (const owner of owners) await owner.close();
        await rm(directory, { recursive: true, force: true });
      }
    },
  );

  test(`${operation} preserves constructor cleanup guidance and replacement protection`, async () => {
    vi.resetModules();
    const { ConsultChimpsError } = await import("@consultchimps/core");
    const { NodeSqliteEngine } = await import("../src/engines/sqlite/node.js");
    const api = await import("../src/node.js");
    const { directory, filename } = await fixture();
    const failure = new ConsultChimpsError(
      "DB_NATIVE_SQLITE_CLEANUP_REQUIRED",
      "Synthetic inaccessible engine",
    );
    const openEngine = vi
      .spyOn(NodeSqliteEngine, "open")
      .mockImplementation(() => {
        throw failure;
      });
    try {
      const open =
        operation === "database"
          ? api.openDatabase
          : operation === "prepared"
            ? api.openPreparedImport
            : api.inspectFileKind;
      await expect(open({ path: filename })).rejects.toBe(failure);
      await expect(
        api.createDatabase({
          path: filename,
          format: "sqlite",
          overwrite: true,
        }),
      ).rejects.toMatchObject({ code: "DB_NATIVE_HANDLE_CLEANUP_REQUIRED" });
    } finally {
      openEngine.mockRestore();
      await rm(directory, { recursive: true, force: true });
    }
  });
}
