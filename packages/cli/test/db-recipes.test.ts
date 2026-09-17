import { readFile, readdir, rm, writeFile } from "node:fs/promises";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import {
  directory,
  cleanupDirectories,
  run,
  runHuman,
  workbook,
  workbookWithInvalidNumericCell,
  cli,
  execute,
} from "./db-support.js";

afterEach(cleanupDirectories);

test("db help describes source-safe persistence and format choices", async () => {
  for (const args of [
    ["--help"],
    ["db", "--help"],
    ["db", "create", "--help"],
    ["db", "import", "--help"],
    ["db", "import", "prepare", "--help"],
  ]) {
    const result = await execute(process.execPath, [cli, ...args], {
      encoding: "utf8",
    });
    expect(result.stderr).toBe("");
    expect(result.stdout).toContain("consultchimps");
    if (args[1] === "import" && args[2] === "prepare")
      expect(result.stdout).toMatch(/existing paths are used as\s+written/u);
  }

  const dbHelp = await execute(process.execPath, [cli, "db", "--help"], {
    encoding: "utf8",
  });
  expect(dbHelp.stdout).toContain("import");
  expect(dbHelp.stdout).not.toMatch(
    /^\s+(plan|apply|resolve|deliveries|delivery)\b/mu,
  );

  const importHelp = await execute(
    process.execPath,
    [cli, "db", "import", "--help"],
    { encoding: "utf8" },
  );
  for (const command of [
    "prepare",
    "inspect",
    "update",
    "apply",
    "run",
    "history",
    "record",
  ]) {
    expect(importHelp.stdout).toMatch(new RegExp(`^  ${command}\\b`, "mu"));
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
    "import",
    "prepare",
    database,
    "--input",
    `inventory=${source}`,
    "--profile",
    recipe,
    "-o",
    plan,
  ]);
  await rm(source);
  await run(["import", "apply", database, "--batch", plan]);
  expect((await run(["inspect", database]))["tables"]).toEqual([
    expect.objectContaining({ name: "left_side", rowCount: "1" }),
    expect.objectContaining({ name: "right_side", rowCount: "1" }),
  ]);
});

test("a replacement recipe excludes routes it omits", async () => {
  const root = await directory();
  const database = path.join(root, "routes.sqlite");
  const source = path.join(root, "routes.xlsx");
  const initialRecipe = path.join(root, "initial.json");
  const replacementRecipe = path.join(root, "replacement.json");
  const plan = path.join(root, "routes.ccplan");
  const selection = (column: string) =>
    JSON.stringify({ range: `Inventory!${column}1:${column}2` });
  const route = (column: "A" | "B", table: string) => ({
    source: "inventory",
    selection: selection(column),
    destination: {
      kind: "new-table",
      schema: {
        name: table,
        recordId: { prefix: column, padding: 4 },
        columns: [{ name: "value", type: "text" }],
      },
    },
    columns: [
      {
        source: column === "A" ? "Left" : "Right",
        target: "value",
        type: "text",
      },
    ],
  });
  const left = route("A", "left_side");
  const right = route("B", "right_side");
  const unavailable = {
    ...route("B", "unavailable"),
    source: "missing",
  };
  await writeFile(
    source,
    workbook([
      ["Left", "Right"],
      ["North", "South"],
    ]),
  );
  await writeFile(
    initialRecipe,
    JSON.stringify({ version: 1, routes: [left, right, unavailable] }),
  );
  await writeFile(
    replacementRecipe,
    JSON.stringify({ version: 1, routes: [left] }),
  );
  await run(["create", "-o", database]);
  await run([
    "import",
    "prepare",
    database,
    "--input",
    `inventory=${source}`,
    "--profile",
    initialRecipe,
    "-o",
    plan,
  ]);
  const resolved = await run([
    "import",
    "update",
    database,
    "--batch",
    plan,
    "--profile",
    replacementRecipe,
  ]);
  expect(resolved).toMatchObject({ state: "ready" });
  await rm(source);
  const applied = await run(["import", "apply", database, "--batch", plan]);
  expect(applied["metrics"]).toMatchObject({ rowsImported: 1 });
  expect((await run(["inspect", database]))["tables"]).toEqual([
    expect.objectContaining({ name: "left_side", rowCount: "1" }),
  ]);
});

test("an empty replacement recipe excludes every captured route", async () => {
  const root = await directory();
  const database = path.join(root, "empty.sqlite");
  const source = path.join(root, "empty.xlsx");
  const initialRecipe = path.join(root, "initial.json");
  const replacementRecipe = path.join(root, "empty.json");
  const plan = path.join(root, "empty.ccplan");
  await writeFile(source, workbook([["Name"], ["North"]]));
  const selection = JSON.stringify({ sheet: "Inventory", headerRow: 1 });
  const destination = (name: string) => ({
    kind: "new-table-infer",
    name,
    recordId: { prefix: name.slice(0, 3).toUpperCase(), padding: 4 },
  });
  await writeFile(
    initialRecipe,
    JSON.stringify({
      version: 1,
      routes: [
        {
          source: "inventory",
          selection,
          destination: destination("inventory"),
          columns: [],
        },
        {
          source: "missing",
          selection: "Unavailable",
          destination: destination("unavailable"),
          columns: [],
        },
      ],
    }),
  );
  await writeFile(
    replacementRecipe,
    JSON.stringify({ version: 1, routes: [] }),
  );
  await run(["create", "-o", database]);
  await run([
    "import",
    "prepare",
    database,
    "--input",
    `inventory=${source}`,
    "--profile",
    initialRecipe,
    "-o",
    plan,
  ]);
  const resolved = await run([
    "import",
    "update",
    database,
    "--batch",
    plan,
    "--profile",
    replacementRecipe,
  ]);
  expect(resolved).toMatchObject({ state: "ready" });
  const applied = await run(["import", "apply", database, "--batch", plan]);
  expect(applied["metrics"]).toMatchObject({ rowsImported: 0 });
  expect((await run(["inspect", database]))["tables"]).toEqual([]);
});

test("forced plan replacement preserves the old plan until capture succeeds", async () => {
  const root = await directory();
  const database = path.join(root, "safe.sqlite");
  const source = path.join(root, "safe.xlsx");
  const plan = path.join(root, "safe.ccplan");
  await writeFile(source, workbook([["Name"], ["North"]]));
  await run(["create", "-o", database]);
  await run([
    "import",
    "prepare",
    database,
    "--input",
    `inventory=${source}`,
    "-o",
    plan,
  ]);
  const originalPlan = await readFile(plan);
  await writeFile(source, await workbookWithInvalidNumericCell());
  await expect(
    execute(process.execPath, [
      cli,
      "--json",
      "db",
      "import",
      "prepare",
      database,
      "--input",
      `inventory=${source}`,
      "-o",
      plan,
      "-f",
    ]),
  ).rejects.toMatchObject({
    code: 1,
    stdout: expect.stringContaining("XLSX_READ_FAILED"),
  });
  expect(await readFile(plan)).toEqual(originalPlan);
  expect((await run(["import", "inspect", plan]))["capturedRows"]).toBe("1");
  expect(
    (await readdir(root)).filter((name) =>
      name.startsWith(`.${path.basename(plan)}.`),
    ),
  ).toEqual([]);

  const originalSource = await readFile(source);
  await expect(
    execute(process.execPath, [
      cli,
      "--json",
      "db",
      "import",
      "prepare",
      database,
      "--input",
      `inventory=${source}`,
      "-o",
      source,
      "-f",
    ]),
  ).rejects.toMatchObject({
    code: 1,
    stdout: expect.stringContaining("FILES_INPUT_OVERWRITE"),
  });
  expect(await readFile(source)).toEqual(originalSource);

  await writeFile(source, workbook([["Name"], ["North"], ["South"]]));
  await run([
    "import",
    "prepare",
    database,
    "--input",
    `inventory=${source}`,
    "-o",
    plan,
    "-f",
  ]);
  expect((await run(["import", "inspect", plan]))["capturedRows"]).toBe("2");
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
      "run",
      database,
      "--input",
      `inventory=${source}`,
      "--profile",
      recipe,
    ]),
  ).rejects.toMatchObject({
    code: 1,
    stdout: expect.stringContaining('"ok":false'),
  });
  expect((await run(["inspect", database]))["tables"]).toEqual([]);
});

test("db inspect explains conflicting mappings for aliases of one capture", async () => {
  const root = await directory();
  const database = path.join(root, "inventory.sqlite");
  const source = path.join(root, "inventory.xlsx");
  const recipe = path.join(root, "recipe.json");
  const plan = path.join(root, "review.ccplan");
  const selection = JSON.stringify({ sheet: "Inventory", headerRow: 1 });
  const schema = {
    name: "Records",
    recordId: { prefix: "REC", padding: 4 },
    columns: [{ name: "Value", type: "text" }],
  };
  await writeFile(
    source,
    workbook([
      ["First", "Second"],
      ["A", "B"],
    ]),
  );
  await writeFile(
    recipe,
    JSON.stringify({
      version: 1,
      routes: [
        {
          source: "first-alias",
          selection,
          destination: { kind: "new-table", schema },
          columns: [{ source: "First", target: "Value", type: "text" }],
        },
        {
          source: "second-alias",
          selection,
          destination: { kind: "existing-table", table: "Records" },
          columns: [{ source: "Second", target: "Value", type: "text" }],
        },
      ],
    }),
  );
  await run(["create", "-o", database]);
  await run([
    "import",
    "prepare",
    database,
    "--input",
    `first-alias=${source}`,
    "--input",
    `second-alias=${source}`,
    "--profile",
    recipe,
    "-o",
    plan,
  ]);

  const inspection = await run(["import", "inspect", plan]);
  expect(inspection).toMatchObject({
    prepared: { state: "needs-review" },
    conflicts: [
      { kind: "conflicting-application-mapping", source: "first-alias" },
      { kind: "conflicting-application-mapping", source: "second-alias" },
    ],
  });
  const report = await runHuman(["import", "inspect", plan]);
  expect(report).toContain(
    'maps the same captured data differently for table "Records"',
  );
  expect(report).toContain(
    "Use one consistent mapping or choose another table",
  );
  expect((await run(["inspect", database]))["tables"]).toEqual([]);
});
