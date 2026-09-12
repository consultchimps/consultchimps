import { expect, test } from "@playwright/test";
import { execFile } from "node:child_process";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";

import { inspectDatabase } from "@consultchimps/db";
import { openDatabase as openNativeDatabase } from "@consultchimps/db/node";

import { createWorkbookUpload } from "./fixtures";

const execFileAsync = promisify(execFile);
const CLI_PATH = path.resolve(
  import.meta.dirname,
  "../../../packages/cli/dist/index.js",
);

const SCHEMA = {
  version: 1,
  tables: [
    {
      name: "datasets",
      recordId: { prefix: "DATASET", padding: 6 },
      columns: [
        { name: "name", type: "text", nullable: false },
        { name: "reported_cde", type: "boolean" },
      ],
    },
  ],
};

async function createDatabase(
  page: import("@playwright/test").Page,
  format: "duckdb" | "sqlite",
): Promise<void> {
  await page.getByTestId("workspace-new-format").selectOption(format);
  await page.getByTestId("workspace-new-name").fill(`acceptance.${format}`);
  await page.getByTestId("workspace-new").click();
  await expect(page.getByTestId("workspace-summary")).toBeVisible();
  await expect(page.getByTestId("workspace-format")).toHaveText(format);
}

test.describe("persistent database workspace", () => {
  test("is reachable from the header and explains its storage", async ({
    page,
  }) => {
    await page.goto("/");
    await page.getByRole("link", { name: "Workspace", exact: true }).click();
    await expect(page).toHaveURL(/\/workspace$/u);
    await expect(
      page.getByRole("heading", { level: 1, name: "Data workspace" }),
    ).toBeVisible();
    await expect(page.getByTestId("workspace-start")).toContainText(
      "origin-private browser storage",
    );
  });

  for (const format of ["sqlite", "duckdb"] as const) {
    test(`creates, changes, exports, and reopens ${format}`, async ({
      page,
    }, testInfo) => {
      const nativeSchemaPath = testInfo.outputPath("native-schema.json");
      const nativeDatabasePath = testInfo.outputPath(`native.${format}`);
      await writeFile(nativeSchemaPath, JSON.stringify(SCHEMA));
      await execFileAsync(process.execPath, [
        CLI_PATH,
        "db",
        "create",
        "--output",
        nativeDatabasePath,
        "--format",
        format,
        "--schema",
        nativeSchemaPath,
      ]);

      await page.goto("/workspace");
      await page
        .getByTestId("workspace-open-input")
        .setInputFiles(nativeDatabasePath);
      await expect(page.getByTestId("workspace-format")).toHaveText(format);
      await expect(page.getByTestId("workspace-table")).toHaveCount(1);

      await createDatabase(page, format);

      await page
        .getByTestId("workspace-schema-input")
        .fill(JSON.stringify(SCHEMA));
      await page.getByTestId("workspace-schema-plan").click();
      await expect(page.getByTestId("workspace-schema-review")).toContainText(
        "1 proposed changes",
      );
      await page.getByTestId("workspace-schema-apply").click();
      await expect(page.getByTestId("workspace-table")).toHaveCount(1);

      await page.reload();
      await expect(page.getByTestId("workspace-summary")).toHaveCount(0);
      await page.getByTestId("workspace-reopen").first().click();
      await expect(page.getByTestId("workspace-table")).toHaveCount(1);

      const downloadPromise = page.waitForEvent("download");
      await page.getByTestId("workspace-export-same").click();
      const download = await downloadPromise;
      const nativeExportPath = testInfo.outputPath(`browser-export.${format}`);
      await download.saveAs(nativeExportPath);
      const bytes = await readFile(nativeExportPath);
      expect(bytes.byteLength).toBeGreaterThan(100);
      if (format === "sqlite") {
        expect(bytes.subarray(0, 15).toString("latin1")).toBe(
          "SQLite format 3",
        );
      }

      await page.getByTestId("workspace-open-input").setInputFiles({
        name: download.suggestedFilename(),
        mimeType: "application/octet-stream",
        buffer: bytes,
      });
      await expect(page.getByTestId("workspace-notice")).toContainText(
        "selected file will not change",
      );
      await expect(page.getByTestId("workspace-format")).toHaveText(format);
      await expect(page.getByTestId("workspace-table")).toHaveCount(1);

      const nativeExport = await openNativeDatabase({
        path: nativeExportPath,
        readonly: true,
      });
      try {
        const inspection = await inspectDatabase({ database: nativeExport });
        expect(inspection.format).toBe(format);
        expect(inspection.tables.map((table) => table.name)).toEqual([
          "datasets",
        ]);
      } finally {
        await nativeExport.close();
      }

      const convertedFormat = format === "sqlite" ? "duckdb" : "sqlite";
      const conversionPromise = page.waitForEvent("download");
      await page.getByTestId("workspace-export-convert").click();
      const conversion = await conversionPromise;
      const conversionPath = testInfo.outputPath(
        `browser-converted.${convertedFormat}`,
      );
      await conversion.saveAs(conversionPath);
      const convertedBytes = await readFile(conversionPath);
      expect(convertedBytes.byteLength).toBeGreaterThan(100);
      if (convertedFormat === "sqlite") {
        expect(convertedBytes.subarray(0, 15).toString("latin1")).toBe(
          "SQLite format 3",
        );
      }
      await page.getByTestId("workspace-open-input").setInputFiles({
        name: conversion.suggestedFilename(),
        mimeType: "application/octet-stream",
        buffer: convertedBytes,
      });
      await expect(page.getByTestId("workspace-format")).toHaveText(
        convertedFormat,
      );
      await expect(page.getByTestId("workspace-table")).toHaveCount(1);

      const nativeConversion = await openNativeDatabase({
        path: conversionPath,
        readonly: true,
      });
      try {
        const inspection = await inspectDatabase({
          database: nativeConversion,
        });
        expect(inspection.format).toBe(convertedFormat);
        expect(inspection.tables.map((table) => table.name)).toEqual([
          "datasets",
        ]);
      } finally {
        await nativeConversion.close();
      }
    });
  }

  test("shows a schema conflict before writing", async ({ page }) => {
    await page.goto("/workspace");
    await createDatabase(page, "sqlite");
    await page
      .getByTestId("workspace-schema-input")
      .fill(JSON.stringify(SCHEMA));
    await page.getByTestId("workspace-schema-plan").click();
    await page.getByTestId("workspace-schema-apply").click();

    const conflicting = structuredClone(SCHEMA);
    conflicting.tables[0]!.columns[0]!.type = "integer";
    await page
      .getByTestId("workspace-schema-input")
      .fill(JSON.stringify(conflicting));
    await page.getByTestId("workspace-schema-plan").click();
    await expect(page.getByTestId("workspace-schema-review")).toContainText(
      /type|conflict/iu,
    );
    await expect(page.getByTestId("workspace-schema-apply")).toBeDisabled();
    await expect(page.getByTestId("workspace-table")).toHaveCount(1);
  });

  test("publishes browser replacements without losing the prior database", async ({
    page,
    context,
  }) => {
    await page.goto("/workspace");
    await page.getByTestId("workspace-new-format").selectOption("sqlite");
    await page.getByTestId("workspace-new-name").fill("collision.sqlite");
    await page.getByTestId("workspace-new").click();
    await page
      .getByTestId("workspace-schema-input")
      .fill(JSON.stringify(SCHEMA));
    await page.getByTestId("workspace-schema-plan").click();
    await page.getByTestId("workspace-schema-apply").click();
    await expect(page.getByTestId("workspace-table")).toHaveCount(1);

    await page.getByTestId("workspace-import-input").setInputFiles(
      await createWorkbookUpload("old-data.xlsx", [
        {
          name: "datasets",
          rows: [["name"], ["Persisted row"]],
        },
      ]),
    );
    await page
      .getByTestId("workspace-import-source")
      .getByTestId("workspace-import-role")
      .fill("inventory");
    await page
      .getByTestId("workspace-import-source")
      .getByTestId("workspace-import-revision")
      .fill("Initial");
    await page.getByTestId("workspace-import-prepare").click();
    await expect(page.getByTestId("workspace-import-review")).toBeVisible();
    await page.getByTestId("workspace-import-route").selectOption("append");
    await page.getByTestId("workspace-import-table").fill("datasets");
    await page.getByTestId("workspace-import-resolve").click();
    await expect(page.getByTestId("workspace-import-review")).toContainText(
      "ready",
    );
    await page.getByTestId("workspace-delivery-vendor").fill("Vendor A");
    await page.getByTestId("workspace-delivery-entity").fill("Entity North");
    await page.getByTestId("workspace-delivery-phase").fill("Initial");
    await page.getByTestId("workspace-delivery-coverage").selectOption("full");
    await page.getByTestId("workspace-import-apply").click();
    await expect(page.getByTestId("workspace-table")).toContainText("1 row");

    await page.getByTestId("workspace-new-format").selectOption("duckdb");
    await page.getByTestId("workspace-new-name").fill("collision.sqlite");
    await page.getByTestId("workspace-new").click();
    await expect(page.getByTestId("workspace-error")).toContainText(
      /already exists/iu,
    );
    const workerUrl = page.workers()[0]?.url();
    expect(workerUrl).toBeDefined();
    expect(workerUrl).toContain("/_next/");
    await page.close();
    await new Promise((resolve) => setTimeout(resolve, 500));

    const direct = await context.newPage();
    await direct.goto("/workspace");
    const result = await direct.evaluate(
      async ({ url, schema }) => {
        const worker = new Worker(url);
        const request = (command: Record<string, unknown>) =>
          new Promise<Record<string, unknown>>((resolve, reject) => {
            const timer = setTimeout(
              () => reject(new Error("Direct workspace worker timed out")),
              10_000,
            );
            const receive = (event: MessageEvent<Record<string, unknown>>) => {
              if (event.data["id"] !== command["id"]) return;
              if (event.data["type"] === "progress") return;
              clearTimeout(timer);
              worker.removeEventListener("message", receive);
              resolve(event.data);
            };
            worker.addEventListener(
              "error",
              (event) => {
                clearTimeout(timer);
                reject(new Error(event.message));
              },
              { once: true },
            );
            worker.addEventListener("message", receive);
            worker.postMessage(command);
          });
        const invalid = await request({
          id: 1,
          type: "create",
          name: "collision.sqlite",
          format: "sqlite",
          overwrite: true,
          schema,
        });
        const reopened = await request({
          id: 2,
          type: "reopen",
          name: "collision.sqlite",
        });
        await request({ id: 3, type: "close" });
        const cancellation = request({
          id: 4,
          type: "create",
          name: "collision.sqlite",
          format: "sqlite",
          overwrite: true,
          schema: {
            version: 1,
            tables: Array.from({ length: 100 }, (_, index) => ({
              name: `replacement_${String(index)}`,
              recordId: { prefix: `R${String(index)}`, padding: 4 },
              columns: [{ name: "value", type: "text" }],
            })),
          },
        });
        worker.postMessage({ id: 5, type: "cancel", targetId: 4 });
        const cancelled = await cancellation;
        const reopenedAfterCancel = await request({
          id: 6,
          type: "reopen",
          name: "collision.sqlite",
        });
        await request({ id: 7, type: "close" });
        await request({
          id: 8,
          type: "create",
          name: "import-source.duckdb",
          format: "duckdb",
          schema: {
            version: 1,
            tables: [
              {
                name: "import_source",
                recordId: { prefix: "SOURCE", padding: 4 },
                columns: [{ name: "value", type: "text" }],
              },
            ],
          },
        });
        const sourceExport = await request({
          id: 9,
          type: "export",
          format: "duckdb",
        });
        await request({ id: 10, type: "close" });
        const importCollision = await request({
          id: 11,
          type: "open",
          name: "collision.sqlite",
          file: sourceExport["file"],
        });
        const reopenedAfterImportCollision = await request({
          id: 12,
          type: "reopen",
          name: "collision.sqlite",
        });
        await request({ id: 13, type: "close" });
        const importedReplacement = await request({
          id: 14,
          type: "open",
          name: "collision.sqlite",
          file: sourceExport["file"],
          overwrite: true,
        });
        await request({ id: 15, type: "close" });
        const reopenedImport = await request({
          id: 16,
          type: "reopen",
          name: "collision.sqlite",
        });
        await request({ id: 17, type: "close" });
        const replaced = await request({
          id: 18,
          type: "create",
          name: "collision.sqlite",
          format: "sqlite",
          overwrite: true,
          schema: {
            version: 1,
            tables: [
              {
                name: "replacement",
                recordId: { prefix: "NEW", padding: 4 },
                columns: [{ name: "value", type: "text" }],
              },
            ],
          },
        });
        await request({ id: 19, type: "close" });
        const reopenedReplacement = await request({
          id: 20,
          type: "reopen",
          name: "collision.sqlite",
        });
        await request({ id: 21, type: "close" });
        await request({
          id: 22,
          type: "reopen",
          name: "collision.sqlite",
          readonly: true,
        });
        const readonlyPlan = await request({
          id: 23,
          type: "planSchema",
          schema: {
            version: 1,
            tables: [
              {
                name: "replacement",
                recordId: { prefix: "NEW", padding: 4 },
                columns: [{ name: "value", type: "text" }],
              },
              {
                name: "readonly_write",
                recordId: { prefix: "READONLY", padding: 4 },
                columns: [{ name: "value", type: "text" }],
              },
            ],
          },
        });
        const readonlyApply = await request({
          id: 24,
          type: "applySchema",
          planId: (readonlyPlan["plan"] as Record<string, unknown>)["id"],
        });
        await request({ id: 25, type: "close" });
        const reopenedAfterReadonly = await request({
          id: 26,
          type: "reopen",
          name: "collision.sqlite",
        });
        await request({ id: 27, type: "close" });
        const firstCreate = request({
          id: 28,
          type: "create",
          name: "race.sqlite",
          format: "sqlite",
          schema: {
            version: 1,
            tables: [
              {
                name: "first_winner",
                recordId: { prefix: "FIRST", padding: 4 },
                columns: [{ name: "value", type: "text" }],
              },
            ],
          },
        });
        const secondCreate = request({
          id: 29,
          type: "create",
          name: "race.sqlite",
          format: "sqlite",
          schema: {
            version: 1,
            tables: [
              {
                name: "second_winner",
                recordId: { prefix: "SECOND", padding: 4 },
                columns: [{ name: "value", type: "text" }],
              },
            ],
          },
        });
        const raced = await Promise.all([firstCreate, secondCreate]);
        await request({ id: 30, type: "close" });
        const reopenedRace = await request({
          id: 31,
          type: "reopen",
          name: "race.sqlite",
        });
        worker.terminate();
        return {
          invalid,
          reopened,
          cancelled,
          reopenedAfterCancel,
          importCollision,
          reopenedAfterImportCollision,
          importedReplacement,
          reopenedImport,
          replaced,
          reopenedReplacement,
          readonlyApply,
          reopenedAfterReadonly,
          raced,
          reopenedRace,
        };
      },
      {
        url: workerUrl!,
        schema: {
          version: 1,
          tables: [
            {
              name: "left_table",
              recordId: { prefix: "LEFT", padding: 4 },
              columns: [{ name: "right_id", type: "text" }],
              foreignKeys: [
                { column: "right_id", referencesTable: "right_table" },
              ],
            },
            {
              name: "right_table",
              recordId: { prefix: "RIGHT", padding: 4 },
              columns: [{ name: "left_id", type: "text" }],
              foreignKeys: [
                { column: "left_id", referencesTable: "left_table" },
              ],
            },
          ],
        },
      },
    );
    expect(result.invalid["type"]).toBe("error");
    expect(result.reopened["type"]).toBe("ready");
    expect(result.reopened["summary"]).toMatchObject({
      format: "sqlite",
      tables: [{ name: "datasets", rowCount: 1 }],
    });
    expect(result.cancelled).toMatchObject({
      type: "error",
      code: "OPERATION_ABORTED",
    });
    expect(result.reopenedAfterCancel["summary"]).toMatchObject({
      format: "sqlite",
      tables: [{ name: "datasets", rowCount: 1 }],
    });
    expect(result.importCollision).toMatchObject({
      type: "error",
      code: "DB_OUTPUT_EXISTS",
    });
    expect(result.reopenedAfterImportCollision["summary"]).toMatchObject({
      format: "sqlite",
      tables: [{ name: "datasets", rowCount: 1 }],
    });
    expect(result.importedReplacement["type"]).toBe("ready");
    expect(result.reopenedImport["summary"]).toMatchObject({
      format: "duckdb",
      tables: [{ name: "import_source", rowCount: 0 }],
    });
    expect(result.replaced["type"]).toBe("ready");
    expect(result.reopenedReplacement["summary"]).toMatchObject({
      format: "sqlite",
      tables: [{ name: "replacement", rowCount: 0 }],
    });
    expect(result.readonlyApply).toMatchObject({
      type: "error",
      message: expect.stringMatching(/read.?only/iu),
    });
    expect(result.reopenedAfterReadonly["summary"]).toMatchObject({
      format: "sqlite",
      tables: [{ name: "replacement", rowCount: 0 }],
    });
    expect(result.raced.map((event) => event["type"]).sort()).toEqual([
      "error",
      "ready",
    ]);
    expect(
      result.raced.find((event) => event["type"] === "error"),
    ).toMatchObject({ code: "DB_OUTPUT_EXISTS" });
    expect(result.reopenedRace["summary"]).toEqual(
      result.raced.find((event) => event["type"] === "ready")?.["summary"],
    );
  });

  test("refuses a file that is not a database", async ({ page }) => {
    await page.goto("/workspace");
    await page.getByTestId("workspace-open-input").setInputFiles({
      name: "broken.sqlite",
      mimeType: "application/vnd.sqlite3",
      buffer: Buffer.from("not a database\n", "utf8"),
    });
    await expect(page.getByTestId("workspace-error")).toContainText(
      /database|SQLite/iu,
    );
    await expect(page.getByTestId("workspace-summary")).toHaveCount(0);
  });
});
