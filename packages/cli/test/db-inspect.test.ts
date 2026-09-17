import { chmod, readFile, readdir, writeFile } from "node:fs/promises";
import path from "node:path";

import { afterEach, expect, test } from "vitest";
import Sqlite from "better-sqlite3";

import {
  directory,
  cleanupDirectories,
  run,
  runHuman,
  workbook,
} from "./db-support.js";

afterEach(cleanupDirectories);

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

test("inspects a read-only saved batch without changing its journal mode or files", async () => {
  const root = await directory();
  const database = path.join(root, "inventory.sqlite");
  const source = path.join(root, "inventory.xlsx");
  const plan = path.join(root, "review.ccplan");
  await writeFile(source, workbook([["Name"], ["Synthetic"]]));
  await run(["create", "-o", database]);
  await run(["import", "prepare", database, "--input", source, "-o", plan]);
  const connection = new Sqlite(plan);
  try {
    connection.pragma("journal_mode = DELETE");
  } finally {
    connection.close();
  }
  const before = await readFile(plan);
  const files = (await readdir(root)).sort();
  await chmod(plan, 0o444);
  try {
    expect((await run(["import", "inspect", plan]))["capturedRows"]).toBe("1");
    expect(await readFile(plan)).toEqual(before);
    expect((await readdir(root)).sort()).toEqual(files);
  } finally {
    await chmod(plan, 0o644);
  }
});

test("renders database review commands as labeled prose while JSON stays structured", async () => {
  const root = await directory();
  const database = path.join(root, "inventory.sqlite");
  const source = path.join(root, "inventory.xlsx");
  const plan = path.join(root, "review.ccplan");
  const context = path.join(root, "delivery.json");
  const schema = path.join(root, "schema.json");
  await writeFile(
    source,
    workbook([
      ["Name"],
      ...Array.from({ length: 25 }, (_, index) => [`Region ${index + 1}`]),
    ]),
  );
  await writeFile(
    context,
    JSON.stringify({ label: "Synthetic delivery", scope: { kind: "full" } }),
  );
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

  const planInspection = await runHuman([
    "import",
    "inspect",
    plan,
    "--limit",
    "25",
  ]);
  expect(planInspection).toContain("Saved import batch inspection");
  expect(planInspection).toContain("Rows newly captured in this batch: 25");
  expect(planInspection).toContain("Bounded row preview:");
  expect(planInspection).toContain('Name="Region 1"');
  expect(planInspection).toContain("source row 26");
  expect(planInspection).toContain("Selection key:");
  expect(planInspection).toContain("Capture ID:");
  expect(planInspection).toContain("captured in this batch");
  expect(planInspection).toContain("Column mappings: Name -> Name (text)");
  expect(planInspection).toContain(
    'Confirm the inferred schema for new table "inventory"',
  );
  expect(planInspection).not.toContain("inferred-schema");
  expect(planInspection).toContain(
    "Safety: This inspection did not change captured rows, routes, or decisions.",
  );
  expect(planInspection).not.toContain('"capturedRows"');

  const firstPreview = await run(["import", "inspect", plan, "--limit", "20"]);
  const nextCursor = firstPreview["nextCursor"];
  if (typeof nextCursor !== "string") {
    throw new Error("The first preview page did not return a cursor.");
  }
  const firstPreviewText = await runHuman([
    "import",
    "inspect",
    plan,
    "--limit",
    "20",
  ]);
  expect(firstPreviewText).toContain("More preview rows are available.");
  const shownCursor = /^Next preview cursor \(data\): (.+)$/mu.exec(
    firstPreviewText,
  )?.[1];
  if (shownCursor === undefined) {
    throw new Error("The human report did not include its preview cursor.");
  }
  expect(shownCursor).toContain("\\");
  const secondPreviewText = await runHuman([
    "import",
    "inspect",
    plan,
    "--limit",
    "20",
    "--cursor",
    shownCursor,
  ]);
  expect(secondPreviewText).toContain("source row 26");
  expect(secondPreviewText).not.toContain("source row 2:");

  const resolution = await runHuman([
    "import",
    "update",
    database,
    "--batch",
    plan,
  ]);
  expect(resolution).toContain("Saved import batch update");
  expect(resolution).toContain("Status: Ready");
  expect(resolution).toContain("No accepted database rows were changed.");
  expect(resolution).not.toContain('"planRevision"');

  const applied = await run([
    "import",
    "apply",
    database,
    "--batch",
    plan,
    "--context",
    context,
    "--request-id",
    "synthetic-delivery",
  ]);
  const captureIds = applied["captureIds"];
  if (!Array.isArray(captureIds) || typeof captureIds[0] !== "string") {
    throw new Error("The import did not report its capture ID.");
  }
  const deliveryConnection = new Sqlite(database);
  try {
    const insertDelivery = deliveryConnection.prepare(
      "INSERT INTO _consultchimps_delivery_events VALUES (?, ?, ?)",
    );
    const insertMembership = deliveryConnection.prepare(
      "INSERT INTO _consultchimps_delivery_memberships VALUES (?, ?)",
    );
    for (let index = 2; index <= 25; index += 1) {
      const id = `DEL-${index.toString().padStart(6, "0")}`;
      const deliveryContext =
        index === 2
          ? {
              label: "Selected region delivery",
              scope: { kind: "partial", description: "Selected regions" },
            }
          : index === 3
            ? {
                label: "Changed records delivery",
                effectiveDate: "2026-09-01",
                receivedDate: "2026-09-02",
                scope: { kind: "changes", baseline: "August delivery" },
                attributes: { reportedRows: 25 },
              }
            : { label: `Synthetic delivery ${index}`, scope: { kind: "full" } };
      insertDelivery.run(
        id,
        `synthetic-delivery-${index}`,
        JSON.stringify(deliveryContext),
      );
      insertMembership.run(id, captureIds[0]);
    }
    deliveryConnection.exec(
      "UPDATE _consultchimps_counters SET next_value = 26 WHERE counter_name = 'delivery'",
    );
  } finally {
    deliveryConnection.close();
  }
  const databaseInspection = await runHuman(["inspect", database]);
  expect(databaseInspection).toContain("Database inspection");
  expect(databaseInspection).toContain("Recorded batches: 25");
  expect(databaseInspection).toContain(
    "Safety: This inspection did not change database tables or stored data.",
  );
  expect(databaseInspection).not.toContain('"completedImports"');

  const deliveries = await runHuman([
    "import",
    "history",
    database,
    "--limit",
    "25",
  ]);
  expect(deliveries).toContain("Batch history");
  expect(deliveries).toContain("Synthetic delivery");
  expect(deliveries).toContain("DEL-000025");
  expect(deliveries).toContain("Scope: Partial, Selected regions");
  expect(deliveries).toContain("Scope: Changes since August delivery");
  expect(deliveries).toContain("Effective date: 2026-09-01");
  expect(deliveries).toContain("Received date: 2026-09-02");
  expect(deliveries).toContain("Reported attributes: reportedRows=25");
  expect(deliveries).toContain(`Capture IDs: ${captureIds[0]}`);
  expect(deliveries).not.toContain('"captureIds"');

  await writeFile(
    schema,
    JSON.stringify({
      version: 1,
      tables: [
        {
          name: "inventory",
          recordId: { prefix: "I", padding: 6 },
          columns: [{ name: "Name", type: "integer" }],
        },
        {
          name: "reviews",
          recordId: { prefix: "RVW", separator: ":", padding: 4 },
          columns: [
            { name: "inventory_id", type: "text", nullable: false },
            { name: "score", type: "integer" },
          ],
          foreignKeys: [
            { column: "inventory_id", referencesTable: "inventory" },
          ],
        },
      ],
    }),
  );

  const schemaReview = await runHuman([
    "schema",
    "apply",
    database,
    "--file",
    schema,
    "--dry-run",
  ]);
  expect(schemaReview).toContain("Database schema review");
  expect(schemaReview).toContain(
    'Column "Name" in table "inventory" is text, but the proposed schema declares integer.',
  );
  expect(schemaReview).toContain('Create table "reviews"');
  expect(schemaReview).toContain(
    'Record IDs: prefix "RVW", separator ":", padding 4',
  );
  expect(schemaReview).toContain("inventory_id text required");
  expect(schemaReview).toContain("score integer optional");
  expect(schemaReview).toContain("inventory_id -> inventory.record_id");
  expect(schemaReview).not.toContain("column-type");
  expect(schemaReview).toContain("Safety: This dry run did not change");
  expect(schemaReview).not.toContain('"schemaFingerprint"');

  const exportReview = await runHuman([
    "export",
    database,
    "-o",
    path.join(root, "converted.duckdb"),
    "--format",
    "duckdb",
    "--dry-run",
  ]);
  expect(exportReview).toContain("Database export review");
  expect(exportReview).toContain("Source format: SQLite");
  expect(exportReview).toContain(
    "Safety: This dry run did not create or replace an output file.",
  );
  expect(exportReview).not.toContain('"sourceFormat"');

  const jsonInspection = await run(["import", "inspect", plan]);
  expect(jsonInspection).toMatchObject({
    capturedRows: "25",
    prepared: { state: "ready" },
  });
});

test("escapes controls inside database report values while JSON preserves them", async () => {
  const root = await directory();
  const database = path.join(root, "controls.sqlite");
  const source = path.join(root, "controls.xlsx");
  const plan = path.join(root, "controls.ccplan");
  const context = path.join(root, "controls.json");
  const sourceAlias = "inventory\nSafety: forged source\r\t\u001B[31m";
  const label = "Synthetic delivery\nSafety: forged delivery\r\t\u001B[32m";
  const requestId = "request\nNext: forged action\r\t\u009B";
  await writeFile(source, workbook([["Name"], ["North"]]));
  await writeFile(context, JSON.stringify({ label, scope: { kind: "full" } }));
  await run(["create", "-o", database]);
  await run([
    "import",
    "prepare",
    database,
    "--input",
    `${sourceAlias}=${source}`,
    "-o",
    plan,
  ]);

  const planText = await runHuman(["import", "inspect", plan]);
  expect(planText).toContain(
    "inventory\\u000ASafety: forged source\\u000D\\u0009\\u001B[31m",
  );
  expect(planText.split("\n")).not.toContain("Safety: forged source");
  expect((await run(["import", "inspect", plan]))["routes"]).toEqual([
    expect.objectContaining({ source: sourceAlias }),
  ]);

  await run(["import", "update", database, "--batch", plan]);
  await run([
    "import",
    "apply",
    database,
    "--batch",
    plan,
    "--context",
    context,
    "--request-id",
    requestId,
  ]);
  const deliveriesText = await runHuman(["import", "history", database]);
  expect(deliveriesText).toContain(
    "Synthetic delivery\\u000ASafety: forged delivery\\u000D\\u0009\\u001B[32m",
  );
  expect(deliveriesText).toContain(
    "request\\u000ANext: forged action\\u000D\\u0009\\u009B",
  );
  expect(deliveriesText.split("\n")).not.toContain("Safety: forged delivery");
  expect(deliveriesText.split("\n")).not.toContain("Next: forged action");
  const deliveryPage = await run(["import", "history", database]);
  expect(deliveryPage["batches"]).toEqual([
    expect.objectContaining({
      requestId,
      context: expect.objectContaining({ label }),
    }),
  ]);
});
