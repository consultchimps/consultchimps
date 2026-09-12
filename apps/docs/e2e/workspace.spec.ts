import { expect, test } from "@playwright/test";
import { execFile } from "node:child_process";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";

import { inspectDatabase } from "@consultchimps/db";
import { openDatabase as openNativeDatabase } from "@consultchimps/db/node";

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
