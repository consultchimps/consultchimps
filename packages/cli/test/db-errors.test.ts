import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import {
  directory,
  cleanupDirectories,
  run,
  runFailure,
  closeFailureLoader,
  workbook,
  cli,
  execute,
} from "./db-support.js";

afterEach(cleanupDirectories);

test.each([
  {
    option: "schema",
    args: (root: string, missing: string) => [
      "db",
      "create",
      "-o",
      path.join(root, "database.sqlite"),
      "--schema",
      missing,
    ],
  },
  {
    option: "profile",
    args: (root: string, missing: string) => [
      "db",
      "import",
      "update",
      path.join(root, "database.sqlite"),
      "--batch",
      path.join(root, "review.ccplan"),
      "--profile",
      missing,
    ],
  },
  {
    option: "context",
    args: (root: string, missing: string) => [
      "db",
      "import",
      "apply",
      path.join(root, "database.sqlite"),
      "--batch",
      path.join(root, "review.ccplan"),
      "--context",
      missing,
    ],
  },
])("missing $option documents have a stable JSON error", async ({ args }) => {
  const root = await directory();
  const missing = path.join(root, "private-configuration.json");
  const failure = await runFailure(["--json", ...args(root, missing)]);

  const expected = {
    ok: false,
    error: {
      code: "DB_DOCUMENT_UNREADABLE",
      message:
        "The JSON configuration file could not be opened. Check that it exists and that you can read it.",
    },
  };
  expect(JSON.parse(failure.stdout)).toEqual(expected);
  expect(JSON.parse(failure.stderr)).toEqual(expected);
  expect(failure.stdout).not.toContain("private-configuration.json");
  expect(failure.stderr).not.toContain("private-configuration.json");
  await expect(
    readFile(path.join(root, "database.sqlite")),
  ).rejects.toMatchObject({ code: "ENOENT" });
});

test("a missing database document has an actionable human error", async () => {
  const root = await directory();
  const missing = path.join(root, "private-configuration.json");
  const output = path.join(root, "database.sqlite");
  await writeFile(output, "existing output");
  const failure = await runFailure([
    "db",
    "create",
    "-o",
    output,
    "--schema",
    missing,
    "--force",
  ]);

  expect(failure.stdout).toBe("");
  expect(failure.stderr).toContain(
    "The JSON configuration file could not be opened. Check that it exists and that you can read it.",
  );
  expect(failure.stderr).toContain("DB_DOCUMENT_UNREADABLE");
  expect(failure.stderr).not.toContain("private-configuration.json");
  expect(await readFile(output, "utf8")).toBe("existing output");
});

test("a close failure after db create defers JSON and human success output", async () => {
  const root = await directory();
  const database = path.join(root, "created.sqlite");
  const loader = await closeFailureLoader(root);
  let stdout: string;
  let stderr: string;

  try {
    await execute(
      process.execPath,
      [
        "--no-warnings",
        "--experimental-loader",
        loader,
        cli,
        "--json",
        "db",
        "create",
        "-o",
        database,
      ],
      { encoding: "utf8" },
    );
    throw new Error("The command unexpectedly succeeded.");
  } catch (error) {
    if (
      !(error instanceof Error) ||
      !("stdout" in error) ||
      typeof error.stdout !== "string" ||
      !("stderr" in error) ||
      typeof error.stderr !== "string"
    )
      throw error;
    stdout = error.stdout;
    stderr = error.stderr;
  }

  expect(stdout, stderr).not.toBe("");
  expect(stdout.trim().split("\n")).toHaveLength(1);
  expect(JSON.parse(stdout)).toEqual({
    ok: false,
    error: {
      code: "CLI_DB_COMMAND_CLEANUP_REQUIRED",
      message: expect.stringContaining(
        "The database operation completed, but its local resources could not finish closing.",
      ),
    },
  });
  expect(JSON.parse(stderr)).toEqual(JSON.parse(stdout));
  expect(stdout).not.toContain('"ok":true');
  await expect(run(["inspect", database])).resolves.toMatchObject({
    format: "sqlite",
  });

  const humanDatabase = path.join(root, "human-created.sqlite");
  let humanStdout: string;
  let humanStderr: string;
  try {
    await execute(
      process.execPath,
      [
        "--no-warnings",
        "--experimental-loader",
        loader,
        cli,
        "db",
        "create",
        "-o",
        humanDatabase,
      ],
      { encoding: "utf8" },
    );
    throw new Error("The command unexpectedly succeeded.");
  } catch (error) {
    if (
      !(error instanceof Error) ||
      !("stdout" in error) ||
      typeof error.stdout !== "string" ||
      !("stderr" in error) ||
      typeof error.stderr !== "string"
    )
      throw error;
    humanStdout = error.stdout;
    humanStderr = error.stderr;
  }
  expect(humanStdout).toBe("");
  expect(humanStderr).toContain(
    "The database operation completed, but its local resources could not finish closing.",
  );
  expect(humanStderr).toContain("CLI_DB_COMMAND_CLEANUP_REQUIRED");
  await expect(run(["inspect", humanDatabase])).resolves.toMatchObject({
    format: "sqlite",
  });
});

test("an existing workbook path containing an equals sign remains a plain path", async () => {
  const root = await directory();
  const database = path.join(root, "inventory.sqlite");
  const source = "current=final.xlsx";
  await writeFile(path.join(root, source), workbook([["Name"], ["North"]]));
  await run(["create", "-o", database]);

  const imported = await run(
    ["import", "run", database, "--input", source],
    root,
  );

  expect(imported["metrics"]).toMatchObject({ rowsImported: 1 });
  expect((await run(["inspect", database]))["tables"]).toEqual([
    expect.objectContaining({ name: "current_final", rowCount: "1" }),
  ]);
});

test("a hidden-only workbook requires explicit hidden-sheet selection", async () => {
  const root = await directory();
  const database = path.join(root, "inventory.sqlite");
  const source = path.join(root, "hidden.xlsx");
  await writeFile(source, workbook([["Name"], ["North"]], true));
  const sourceBefore = await readFile(source);
  await run(["create", "-o", database]);
  const databaseBefore = await run(["inspect", database]);

  const failure = await runFailure([
    "--json",
    "db",
    "import",
    "run",
    database,
    "--input",
    source,
  ]);
  const expectedFailure = {
    ok: false,
    error: {
      code: "DB_IMPORT_NO_SELECTIONS",
      message:
        "Choose at least one source region before importing. If the workbook contains only hidden worksheets, include hidden sheets and try again.",
    },
  };
  expect(JSON.parse(failure.stdout)).toEqual(expectedFailure);
  expect(JSON.parse(failure.stderr)).toEqual(expectedFailure);
  expect(await run(["inspect", database])).toEqual(databaseBefore);
  expect((await readFile(source)).equals(sourceBefore)).toBe(true);

  const imported = await run([
    "import",
    "run",
    database,
    "--input",
    source,
    "--hidden",
  ]);
  expect(imported["metrics"]).toMatchObject({
    rowsImported: 1,
    tablesCreated: 1,
  });
  expect((await readFile(source)).equals(sourceBefore)).toBe(true);
});
