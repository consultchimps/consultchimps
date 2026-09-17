import {
  chmod,
  copyFile,
  readFile,
  readdir,
  rm,
  writeFile,
} from "node:fs/promises";
import path from "node:path";

import { expect, test } from "vitest";
import Sqlite from "better-sqlite3";

import {
  directory,
  run,
  runFailure,
  workbook,
  addUnsupportedObject,
  cli,
  execute,
} from "./db-support.js";

export function defineFormatTests(format: "sqlite" | "duckdb"): void {
  test(`${format}: applies ready plans read-only and reopens unresolved plans for resolution`, async () => {
    const root = await directory();
    const database = path.join(root, `inventory.${format}`);
    const readySource = path.join(root, "ready.xlsx");
    const readyRecipe = path.join(root, "ready.json");
    const readyPlan = path.join(root, "ready.ccplan");
    await writeFile(readySource, workbook([["Name"], ["North"]]));
    await writeFile(
      readyRecipe,
      JSON.stringify({
        version: 1,
        routes: [
          {
            source: "ready",
            selection: JSON.stringify({
              sheet: "Inventory",
              headerRow: 1,
            }),
            destination: {
              kind: "new-table",
              schema: {
                name: "ready",
                recordId: { prefix: "RDY", padding: 4 },
                columns: [{ name: "name", type: "text" }],
              },
            },
            columns: [{ source: "Name", target: "name", type: "text" }],
          },
        ],
      }),
    );
    await run(["create", "-o", database, "--format", format]);
    await run([
      "import",
      "prepare",
      database,
      "--input",
      `ready=${readySource}`,
      "--profile",
      readyRecipe,
      "-o",
      readyPlan,
    ]);
    expect(await run(["import", "inspect", readyPlan])).toMatchObject({
      prepared: { state: "ready" },
    });

    const readyConnection = new Sqlite(readyPlan);
    try {
      readyConnection.pragma("journal_mode = DELETE");
    } finally {
      readyConnection.close();
    }
    const planBefore = await readFile(readyPlan);
    const filesBefore = (await readdir(root)).sort();
    await chmod(readyPlan, 0o444);
    try {
      const applied = await run([
        "import",
        "apply",
        database,
        "--batch",
        readyPlan,
      ]);
      expect(applied["metrics"]).toMatchObject({ rowsImported: 1 });
      expect(await readFile(readyPlan)).toEqual(planBefore);
      expect((await readdir(root)).sort()).toEqual(filesBefore);
    } finally {
      await chmod(readyPlan, 0o644);
    }

    const unresolvedSource = path.join(root, "unresolved.xlsx");
    const unresolvedPlan = path.join(root, "unresolved.ccplan");
    await writeFile(unresolvedSource, workbook([["Name"], ["South"]]));
    await run([
      "import",
      "prepare",
      database,
      "--input",
      `unresolved=${unresolvedSource}`,
      "-o",
      unresolvedPlan,
    ]);
    expect(await run(["import", "inspect", unresolvedPlan])).toMatchObject({
      prepared: { state: "needs-review" },
    });

    const resolved = await run([
      "import",
      "apply",
      database,
      "--batch",
      unresolvedPlan,
    ]);
    expect(resolved["metrics"]).toMatchObject({ rowsImported: 1 });
    expect(await run(["import", "inspect", unresolvedPlan])).toMatchObject({
      prepared: { state: "ready" },
    });
    expect((await run(["inspect", database]))["tables"]).toEqual([
      expect.objectContaining({ name: "ready", rowCount: "1" }),
      expect.objectContaining({ name: "unresolved", rowCount: "1" }),
    ]);
  });

  test(`${format}: dry-run reviews conversion without reserving or changing output files`, async () => {
    const root = await directory();
    const source = path.join(root, `source.${format}`);
    const targetFormat = format === "sqlite" ? "duckdb" : "sqlite";
    const output = path.join(root, `existing.${targetFormat}`);
    const sentinel = Buffer.from("existing destination");
    await run(["create", "-o", source, "--format", format]);
    if (format === "sqlite") {
      const database = new Sqlite(source);
      try {
        database.pragma("journal_mode = DELETE");
      } finally {
        database.close();
      }
    }

    const sourceBeforeAliasReview = await readFile(source);
    const filesBeforeAliasReview = (await readdir(root)).sort();
    const aliasPlan = await run([
      "export",
      source,
      "-o",
      source,
      "--format",
      format,
      "--dry-run",
    ]);
    expect(aliasPlan).toMatchObject({
      sourceFormat: format,
      targetFormat: format,
      state: "ready",
    });
    expect(await readFile(source)).toEqual(sourceBeforeAliasReview);
    expect((await readdir(root)).sort()).toEqual(filesBeforeAliasReview);

    await writeFile(output, sentinel);
    const sourceBeforeConversionReview = await readFile(source);
    const filesBeforeConversionReview = (await readdir(root)).sort();
    const conversionPlan = await run([
      "export",
      source,
      "-o",
      output,
      "--format",
      targetFormat,
      "--dry-run",
    ]);
    expect(conversionPlan).toMatchObject({
      sourceFormat: format,
      targetFormat,
      state: "ready",
      issues: [],
    });
    expect(await readFile(source)).toEqual(sourceBeforeConversionReview);
    expect(await readFile(output)).toEqual(sentinel);
    expect((await readdir(root)).sort()).toEqual(filesBeforeConversionReview);

    const publicationFailure = await runFailure([
      "--json",
      "db",
      "export",
      source,
      "-o",
      output,
      "--format",
      targetFormat,
    ]);
    expect(publicationFailure.stdout).toContain("OUTPUT_EXISTS");
    expect(await readFile(source)).toEqual(sourceBeforeConversionReview);
    expect(await readFile(output)).toEqual(sentinel);
    expect((await readdir(root)).sort()).toEqual(filesBeforeConversionReview);

    const conflict = await runFailure([
      "--json",
      "db",
      "export",
      source,
      "-o",
      output,
      "--format",
      format,
      "--dry-run",
    ]);
    expect(conflict.stdout).toContain("DB_FORMAT_CONFLICT");
    expect(await readFile(source)).toEqual(sourceBeforeConversionReview);
    expect(await readFile(output)).toEqual(sentinel);
    expect((await readdir(root)).sort()).toEqual(filesBeforeConversionReview);

    await addUnsupportedObject(source, format);
    const sourceBeforeUnsupportedReview = await readFile(source);
    const filesBeforeUnsupportedReview = (await readdir(root)).sort();
    const unsupportedPlan = await run([
      "export",
      source,
      "-o",
      output,
      "--format",
      targetFormat,
      "--dry-run",
    ]);
    expect(unsupportedPlan).toMatchObject({
      sourceFormat: format,
      targetFormat,
      state: "unsupported",
      issues: [
        expect.objectContaining({
          kind: "unsupported-object",
          name: format === "sqlite" ? "unsupported_view" : "unsupported_macro",
        }),
      ],
    });
    expect(await readFile(source)).toEqual(sourceBeforeUnsupportedReview);
    expect(await readFile(output)).toEqual(sentinel);
    expect((await readdir(root)).sort()).toEqual(filesBeforeUnsupportedReview);
  });

  test(`${format}: previews reused captures from the target database without Excel`, async () => {
    const root = await directory();
    const database = path.join(root, `inventory.${format}`);
    const source = path.join(root, "inventory.xlsx");
    const plan = path.join(root, "duplicate.ccplan");
    await writeFile(source, workbook([["Name"], ["North"], ["South"]]));
    await run(["create", "-o", database]);
    await run(["import", "run", database, "--input", source]);
    await run(["import", "prepare", database, "--input", source, "-o", plan]);
    await rm(source);

    const offline = await run(["import", "inspect", plan]);
    expect(offline["previewWarnings"]).toEqual([
      expect.objectContaining({ code: "DB_PREVIEW_DATABASE_REQUIRED" }),
    ]);
    const preview = await run([
      "import",
      "inspect",
      plan,
      "--database",
      database,
    ]);
    expect(preview["examples"]).toEqual([
      expect.objectContaining({
        sourceRow: 2,
        values: { Name: { kind: "string", value: "North" } },
      }),
      expect.objectContaining({
        sourceRow: 3,
        values: { Name: { kind: "string", value: "South" } },
      }),
    ]);
    expect(preview["previewWarnings"]).toEqual([]);
  });

  test(`${format}: re-reviews a stale saved batch without returning to Excel`, async () => {
    const root = await directory();
    const database = path.join(root, `inventory.${format}`);
    const source = path.join(root, "inventory.xlsx");
    const firstPlan = path.join(root, "first.ccplan");
    const secondPlan = path.join(root, "second.ccplan");
    await run(["create", "-o", database]);
    await writeFile(source, workbook([["Name"], ["North"]]));
    await run([
      "import",
      "prepare",
      database,
      "--input",
      `first=${source}`,
      "-o",
      firstPlan,
    ]);
    await writeFile(source, workbook([["Name"], ["South"]]));
    await run([
      "import",
      "prepare",
      database,
      "--input",
      `second=${source}`,
      "-o",
      secondPlan,
    ]);
    await rm(source);
    await run(["import", "apply", database, "--batch", firstPlan]);
    await expect(
      execute(process.execPath, [
        cli,
        "--json",
        "db",
        "import",
        "apply",
        database,
        "--batch",
        secondPlan,
      ]),
    ).rejects.toMatchObject({
      code: 1,
      stdout: expect.stringContaining("DB_STALE_IMPORT_PLAN"),
    });
    await run(["import", "update", database, "--batch", secondPlan]);
    const applied = await run([
      "import",
      "apply",
      database,
      "--batch",
      secondPlan,
    ]);
    expect(applied["metrics"]).toMatchObject({ rowsImported: 1 });
    expect((await run(["inspect", database]))["tables"]).toEqual([
      expect.objectContaining({ name: "first", rowCount: "1" }),
      expect.objectContaining({ name: "second", rowCount: "1" }),
    ]);
    const retried = await run([
      "import",
      "apply",
      database,
      "--batch",
      secondPlan,
    ]);
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
      "import",
      "prepare",
      database,
      "--input",
      `inventory=${source}`,
      "-o",
      prepared,
    ]);
    expect(plan["metrics"]).toMatchObject({ rowsCaptured: 2 });
    const preview = await run(["import", "inspect", prepared]);
    expect(preview["capturedRows"]).toBe("2");
    const renamedPlan = path.join(root, "renamed-plan.sqlite");
    await copyFile(prepared, renamedPlan);
    expect(
      (await run(["import", "inspect", renamedPlan]))["capturedRows"],
    ).toBe("2");
    await rm(source);
    const applied = await run([
      "import",
      "apply",
      database,
      "--batch",
      prepared,
    ]);
    expect(applied["metrics"]).toMatchObject({ rowsImported: 2 });
    const inspected = await run(["inspect", database]);
    expect(inspected["format"]).toBe(format);
    expect(inspected["tables"]).toEqual([
      expect.objectContaining({ rowCount: "2" }),
    ]);
    const retried = await run([
      "import",
      "apply",
      database,
      "--batch",
      prepared,
    ]);
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
    await run(["import", "run", database, "--input", `inventory=${source}`]);
    const repeated = await run([
      "import",
      "run",
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
      "run",
      database,
      "--input",
      `inventory=${source}`,
      "--profile",
      recipe,
      "--context",
      context,
      "--request-id",
      "submission-a",
    ]);
    await writeFile(source, workbook([["Name"], ["North revised"], ["South"]]));
    const second = await run([
      "import",
      "run",
      database,
      "--input",
      `inventory=${source}`,
      "--profile",
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
      "import",
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
    expect(delivered["metrics"]).toMatchObject({ batchesRecorded: 1 });
    const retried = await run(deliveryArguments);
    expect(retried["metrics"]).toMatchObject({ batchesRecorded: 0 });
    const before = await run(["inspect", database]);
    expect(before["tables"]).toEqual([
      expect.objectContaining({ name: "datasets", rowCount: "3" }),
    ]);
    expect(before["recordedBatches"]).toBe("3");
    await run(["export", database, "-o", converted]);
    const after = await run(["inspect", converted]);
    expect(after["format"]).toBe(targetFormat);
    expect(after["tables"]).toEqual(before["tables"]);
    expect(after["captures"]).toBe(before["captures"]);
    expect(await run(["import", "history", converted])).toEqual(
      await run(["import", "history", database]),
    );
  });
}
