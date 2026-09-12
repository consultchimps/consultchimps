import { execFile } from "node:child_process";
import { copyFile, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { afterEach, expect, test } from "vitest";
import * as XLSX from "xlsx";
import Sqlite from "better-sqlite3";

const execute = promisify(execFile);
const cli = fileURLToPath(new URL("../dist/index.js", import.meta.url));
const directories: string[] = [];

async function directory(): Promise<string> {
  const value = await mkdtemp(path.join(tmpdir(), "cc-db-cli-test-"));
  directories.push(value);
  return value;
}

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((value) => rm(value, { recursive: true, force: true })),
  );
});

async function run(args: string[]): Promise<Record<string, unknown>> {
  const result = await execute(
    process.execPath,
    [cli, "--json", "db", ...args],
    { encoding: "utf8" },
  );
  expect(result.stderr).toBe("");
  expect(result.stdout.trim().split("\n")).toHaveLength(1);
  const envelope: unknown = JSON.parse(result.stdout);
  expect(envelope).toMatchObject({ ok: true });
  if (
    envelope === null ||
    typeof envelope !== "object" ||
    !("result" in envelope) ||
    envelope.result === null ||
    typeof envelope.result !== "object"
  )
    throw new Error("Missing command result.");
  return envelope.result as Record<string, unknown>;
}

function workbook(
  rows: readonly (readonly (string | number | boolean)[])[],
): Uint8Array {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    book,
    XLSX.utils.aoa_to_sheet(rows.map((row) => [...row])),
    "Inventory",
  );
  return XLSX.write(book, { type: "buffer", bookType: "xlsx" }) as Uint8Array;
}

test("inspects an ordinary SQLite database without adopting or changing it", async () => {
  const root = await directory();
  const file = path.join(root, "external.sqlite");
  const external = new Sqlite(file);
  external.exec(
    "CREATE TABLE existing (name TEXT NOT NULL, count INTEGER); INSERT INTO existing VALUES ('Synthetic', 2)",
  );
  external.close();
  const before = await readFile(file);
  expect(await run(["inspect", file])).toMatchObject({
    kind: "unmanaged-database",
    format: "sqlite",
    tables: [
      {
        name: "existing",
        columns: [
          { name: "name", storageType: "TEXT", nullable: false },
          { name: "count", storageType: "INTEGER", nullable: true },
        ],
      },
    ],
  });
  expect(await readFile(file)).toEqual(before);
});

for (const format of ["sqlite", "duckdb"]) {
  test(`${format}: re-reviews a stale saved plan without returning to Excel`, async () => {
    const root = await directory();
    const database = path.join(root, `inventory.${format}`);
    const source = path.join(root, "inventory.xlsx");
    const firstPlan = path.join(root, "first.ccplan");
    const secondPlan = path.join(root, "second.ccplan");
    await run(["create", "-o", database]);
    await writeFile(source, workbook([["Name"], ["North"]]));
    await run([
      "plan",
      database,
      "--input",
      `first=${source}`,
      "-o",
      firstPlan,
    ]);
    await writeFile(source, workbook([["Name"], ["South"]]));
    await run([
      "plan",
      database,
      "--input",
      `second=${source}`,
      "-o",
      secondPlan,
    ]);
    await rm(source);
    await run(["apply", database, "--plan", firstPlan]);
    await expect(
      execute(process.execPath, [
        cli,
        "--json",
        "db",
        "apply",
        database,
        "--plan",
        secondPlan,
      ]),
    ).rejects.toMatchObject({
      code: 1,
      stdout: expect.stringContaining("DB_STALE_IMPORT_PLAN"),
    });
    await run(["resolve", database, "--plan", secondPlan]);
    const applied = await run(["apply", database, "--plan", secondPlan]);
    expect(applied["metrics"]).toMatchObject({ rowsImported: 1 });
    expect((await run(["inspect", database]))["tables"]).toEqual([
      expect.objectContaining({ name: "first", rowCount: "1" }),
      expect.objectContaining({ name: "second", rowCount: "1" }),
    ]);
    const retried = await run(["apply", database, "--plan", secondPlan]);
    expect(retried["captureIds"]).toEqual(applied["captureIds"]);
    expect(retried["metrics"]).toMatchObject({
      rowsImported: 0,
      rowsReused: 1,
    });
  });

  test(`${format}: create, prepare, remove Excel, apply, and reopen through built CLI`, async () => {
    const root = await directory();
    const database = path.join(root, `inventory.${format}`);
    const source = path.join(root, "inventory.xlsx");
    const prepared = path.join(root, "review.ccplan");
    await writeFile(
      source,
      workbook([
        ["Dataset", "Attributes", "CDE"],
        ["North", 12, true],
        ["South", 9, false],
      ]),
    );
    await run(["create", "-o", database, "--format", format]);
    const plan = await run([
      "plan",
      database,
      "--input",
      `inventory=${source}`,
      "-o",
      prepared,
    ]);
    expect(plan["metrics"]).toMatchObject({ rowsCaptured: 2 });
    const preview = await run(["inspect", prepared]);
    expect(preview["capturedRows"]).toBe("2");
    const renamedPlan = path.join(root, "renamed-plan.sqlite");
    await copyFile(prepared, renamedPlan);
    expect((await run(["inspect", renamedPlan]))["capturedRows"]).toBe("2");
    await rm(source);
    const applied = await run(["apply", database, "--plan", prepared]);
    expect(applied["metrics"]).toMatchObject({ rowsImported: 2 });
    const inspected = await run(["inspect", database]);
    expect(inspected["format"]).toBe(format);
    expect(inspected["tables"]).toEqual([
      expect.objectContaining({ rowCount: "2" }),
    ]);
    const retried = await run(["apply", database, "--plan", prepared]);
    expect(retried["metrics"]).toMatchObject({ rowsImported: 0 });
  });

  test(`${format}: identical renamed workbook adds no observations`, async () => {
    const root = await directory();
    const database = path.join(root, `inventory.${format}`);
    const source = path.join(root, "first.xlsx");
    const renamed = path.join(root, "renamed.xlsx");
    await writeFile(source, workbook([["Name"], ["Synthetic dataset"]]));
    await copyFile(source, renamed);
    const original = await readFile(source);
    await run(["create", "-o", database]);
    await run(["import", database, "--input", `inventory=${source}`]);
    const repeated = await run([
      "import",
      database,
      "--input",
      `inventory=${renamed}`,
    ]);
    expect(repeated["metrics"]).toMatchObject({ rowsImported: 0 });
    const inspected = await run(["inspect", database]);
    expect(inspected["tables"]).toEqual([
      expect.objectContaining({ rowCount: "1" }),
    ]);
    expect(await readFile(source)).toEqual(original);
    expect(await readFile(renamed)).toEqual(original);
  });

  test(`${format}: changed submissions append, delivery retries are distinct from content, and conversion retains history`, async () => {
    const root = await directory();
    const database = path.join(root, `inventory.${format}`);
    const targetFormat = format === "sqlite" ? "duckdb" : "sqlite";
    const converted = path.join(root, `converted.${targetFormat}`);
    const source = path.join(root, "submission.xlsx");
    const recipe = path.join(root, "recipe.json");
    const context = path.join(root, "delivery.json");
    const schema = path.join(root, "schema.json");
    await writeFile(
      schema,
      JSON.stringify({
        version: 1,
        tables: [
          {
            name: "datasets",
            recordId: { prefix: "DS", padding: 4 },
            columns: [{ name: "name", type: "text" }],
          },
        ],
      }),
    );
    await writeFile(
      recipe,
      JSON.stringify({
        version: 1,
        routes: [
          {
            source: "inventory",
            selection: JSON.stringify({ sheet: "Inventory", headerRow: 1 }),
            destination: { kind: "existing-table", table: "datasets" },
            columns: [{ source: "Name", target: "name", type: "text" }],
          },
        ],
      }),
    );
    await writeFile(
      context,
      JSON.stringify({
        label: "Synthetic partial submission",
        scope: { kind: "partial", description: "Selected records" },
        attributes: { reportedCount: 9 },
      }),
    );
    await writeFile(source, workbook([["Name"], ["North"]]));
    await run(["create", "-o", database, "--schema", schema]);
    const first = await run([
      "import",
      database,
      "--input",
      `inventory=${source}`,
      "--recipe",
      recipe,
      "--context",
      context,
      "--request-id",
      "submission-a",
    ]);
    await writeFile(source, workbook([["Name"], ["North revised"], ["South"]]));
    const second = await run([
      "import",
      database,
      "--input",
      `inventory=${source}`,
      "--recipe",
      recipe,
      "--context",
      context,
      "--request-id",
      "submission-b",
    ]);
    expect(first["metrics"]).toMatchObject({ rowsImported: 1 });
    expect(second["metrics"]).toMatchObject({ rowsImported: 2 });
    const captures = second["captureIds"];
    if (!Array.isArray(captures) || typeof captures[0] !== "string")
      throw new Error("Import did not report its capture.");
    const deliveryArguments = [
      "delivery",
      "record",
      database,
      "--capture",
      captures[0],
      "--context",
      context,
      "--request-id",
      "submission-c",
    ];
    const delivered = await run(deliveryArguments);
    expect(delivered["metrics"]).toMatchObject({ deliveriesRecorded: 1 });
    const retried = await run(deliveryArguments);
    expect(retried["metrics"]).toMatchObject({ deliveriesRecorded: 0 });
    const before = await run(["inspect", database]);
    expect(before["tables"]).toEqual([
      expect.objectContaining({ name: "datasets", rowCount: "3" }),
    ]);
    expect(before["deliveries"]).toBe("3");
    await run(["export", database, "-o", converted]);
    const after = await run(["inspect", converted]);
    expect(after["format"]).toBe(targetFormat);
    expect(after["tables"]).toEqual(before["tables"]);
    expect(after["captures"]).toBe(before["captures"]);
    expect(await run(["deliveries", converted])).toEqual(
      await run(["deliveries", database]),
    );
  });
}

test("db help describes source-safe persistence and format choices", async () => {
  for (const args of [
    ["--help"],
    ["db", "--help"],
    ["db", "create", "--help"],
    ["db", "plan", "--help"],
  ]) {
    const result = await execute(process.execPath, [cli, ...args], {
      encoding: "utf8",
    });
    expect(result.stderr).toBe("");
    expect(result.stdout).toContain("consultchimps");
  }
});

test("a create format mismatch fails before producing a database", async () => {
  const root = await directory();
  const output = path.join(root, "inventory.sqlite");
  await expect(
    execute(process.execPath, [
      cli,
      "--json",
      "db",
      "create",
      "-o",
      output,
      "--format",
      "duckdb",
    ]),
  ).rejects.toMatchObject({
    code: 1,
    stdout: expect.stringContaining("DB_FORMAT_CONFLICT"),
  });
  await expect(readFile(output)).rejects.toMatchObject({ code: "ENOENT" });
});

test("one recipe can select several distinct regions in the same workbook", async () => {
  const root = await directory();
  const database = path.join(root, "inventory.duckdb");
  const source = path.join(root, "inventory.xlsx");
  const recipe = path.join(root, "recipe.json");
  const plan = path.join(root, "regions.ccplan");
  await writeFile(
    source,
    workbook([
      ["Left", "Right"],
      ["North", "South"],
    ]),
  );
  await writeFile(
    recipe,
    JSON.stringify({
      version: 1,
      routes: [
        {
          source: "inventory",
          selection: JSON.stringify({ range: "Inventory!A1:A2" }),
          destination: {
            kind: "new-table-infer",
            name: "left_side",
            recordId: { prefix: "L", padding: 4 },
          },
          columns: [],
        },
        {
          source: "inventory",
          selection: JSON.stringify({ range: "Inventory!B1:B2" }),
          destination: {
            kind: "new-table-infer",
            name: "right_side",
            recordId: { prefix: "R", padding: 4 },
          },
          columns: [],
        },
      ],
    }),
  );
  await run(["create", "-o", database]);
  await run([
    "plan",
    database,
    "--input",
    `inventory=${source}`,
    "--recipe",
    recipe,
    "-o",
    plan,
  ]);
  await rm(source);
  await run(["apply", database, "--plan", plan]);
  expect((await run(["inspect", database]))["tables"]).toEqual([
    expect.objectContaining({ name: "left_side", rowCount: "1" }),
    expect.objectContaining({ name: "right_side", rowCount: "1" }),
  ]);
});

test("a recipe naming an unavailable source fails instead of accepting an empty import", async () => {
  const root = await directory();
  const database = path.join(root, "inventory.sqlite");
  const source = path.join(root, "inventory.xlsx");
  const recipe = path.join(root, "recipe.json");
  await writeFile(source, workbook([["Name"], ["North"]]));
  await writeFile(
    recipe,
    JSON.stringify({
      version: 1,
      routes: [
        {
          source: "missing",
          selection: JSON.stringify({ sheet: "Inventory", headerRow: 1 }),
          destination: {
            kind: "new-table-infer",
            name: "datasets",
            recordId: { prefix: "DS", padding: 4 },
          },
          columns: [],
        },
      ],
    }),
  );
  await run(["create", "-o", database]);
  await expect(
    execute(process.execPath, [
      cli,
      "--json",
      "db",
      "import",
      database,
      "--input",
      `inventory=${source}`,
      "--recipe",
      recipe,
    ]),
  ).rejects.toMatchObject({
    code: 1,
    stdout: expect.stringContaining('"ok":false'),
  });
  expect((await run(["inspect", database]))["tables"]).toEqual([]);
});
