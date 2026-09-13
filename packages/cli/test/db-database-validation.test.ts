import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { DuckDBInstance } from "@duckdb/node-api";
import Sqlite from "better-sqlite3";
import { afterEach, expect, test } from "vitest";

import { createDatabase } from "@consultchimps/db/node";

const execute = promisify(execFile);
const cli = fileURLToPath(new URL("../dist/index.js", import.meta.url));
const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

async function runFailure(args: readonly string[]): Promise<{
  readonly stdout: string;
  readonly stderr: string;
}> {
  try {
    await execute(process.execPath, [cli, ...args], { encoding: "utf8" });
  } catch (error) {
    if (
      error instanceof Error &&
      "stdout" in error &&
      typeof error.stdout === "string" &&
      "stderr" in error &&
      typeof error.stderr === "string"
    ) {
      return { stdout: error.stdout, stderr: error.stderr };
    }
    throw error;
  }
  throw new Error("The command unexpectedly succeeded");
}

async function alterFixture(
  databasePath: string,
  format: "sqlite" | "duckdb",
  sql: string,
): Promise<void> {
  if (format === "sqlite") {
    const sqlite = new Sqlite(databasePath);
    try {
      sqlite.exec(sql);
    } finally {
      sqlite.close();
    }
  } else {
    const instance = await DuckDBInstance.create(databasePath);
    try {
      const connection = await instance.connect();
      try {
        await connection.run(sql);
      } finally {
        connection.closeSync();
      }
    } finally {
      instance.closeSync();
    }
  }
}

test.each(["sqlite", "duckdb"] as const)(
  "db inspect reports damaged %s managed metadata",
  async (format) => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-cli-db-damage-"));
    directories.push(directory);
    const databasePath = path.join(directory, `private-workspace.${format}`);
    const { database } = await createDatabase({ path: databasePath, format });
    await database.close();

    await alterFixture(
      databasePath,
      format,
      "DROP TABLE _consultchimps_captures",
    );

    const failure = await runFailure(["--json", "db", "inspect", databasePath]);
    const expected = {
      ok: false,
      error: {
        code: "DB_CORRUPT_DATABASE",
        message:
          "The database is incomplete or damaged. Restore a verified database copy before retrying.",
      },
    };
    expect(JSON.parse(failure.stdout)).toEqual(expected);
    expect(JSON.parse(failure.stderr)).toEqual(expected);
    expect(failure.stdout).not.toContain(databasePath);
    expect(failure.stderr).not.toContain(databasePath);
    expect(failure.stdout).not.toContain("_consultchimps_captures");
    expect(failure.stderr).not.toContain("_consultchimps_captures");
  },
);

test.each(["sqlite", "duckdb"] as const)(
  "db deliveries reports damaged %s history with a stable error",
  async (format) => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "cc-cli-delivery-damage-"),
    );
    directories.push(directory);
    const databasePath = path.join(directory, `history.${format}`);
    const { database } = await createDatabase({ path: databasePath, format });
    await database.close();
    await alterFixture(
      databasePath,
      format,
      "INSERT INTO _consultchimps_delivery_events VALUES ('DEL-000001', 'synthetic-request', '{')",
    );

    const failure = await runFailure([
      "--json",
      "db",
      "deliveries",
      databasePath,
    ]);
    const expected = {
      ok: false,
      error: {
        code: "DB_CORRUPT_DATABASE",
        message:
          "The recorded delivery details are damaged. Restore a verified database copy before retrying.",
      },
    };
    expect(JSON.parse(failure.stdout)).toEqual(expected);
    expect(JSON.parse(failure.stderr)).toEqual(expected);
    expect(failure.stdout).not.toContain(databasePath);
    expect(failure.stderr).not.toContain("SyntaxError");
  },
);

test.each(["missing", "invalid", "duckdb"] as const)(
  "db apply reports a stable error for a %s plan without changing the database",
  async (kind) => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-cli-plan-open-"));
    directories.push(directory);
    const databasePath = path.join(directory, "workspace.sqlite");
    const planPath = path.join(directory, "review.ccplan");
    const { database } = await createDatabase({
      path: databasePath,
      format: "sqlite",
    });
    await database.close();
    if (kind === "invalid")
      await writeFile(planPath, "Synthetic non-database input");
    if (kind === "duckdb") {
      const created = await createDatabase({
        path: planPath,
        format: "duckdb",
      });
      await created.database.close();
    }
    const before = await readFile(databasePath);
    const planBefore =
      kind === "missing" ? undefined : await readFile(planPath);
    const failure = await runFailure([
      "--json",
      "db",
      "apply",
      databasePath,
      "--plan",
      planPath,
    ]);
    for (const output of [failure.stdout, failure.stderr]) {
      expect(JSON.parse(output)).toMatchObject({
        ok: false,
        error: {
          code: "DB_INVALID_PREPARED_IMPORT",
          message: expect.stringMatching(/import plan/u),
        },
      });
      expect(output).not.toContain(planPath);
      expect(output).not.toContain("SqliteError");
    }
    expect((await readFile(databasePath)).equals(before)).toBe(true);
    if (planBefore !== undefined)
      expect((await readFile(planPath)).equals(planBefore)).toBe(true);
    else await expect(stat(planPath)).rejects.toMatchObject({ code: "ENOENT" });
  },
);
