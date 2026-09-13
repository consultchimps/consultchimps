import { expect, test } from "@playwright/test";

import { createDatabase } from "@consultchimps/db/node";

import { createWorkbookUpload } from "./fixtures";

const SUPPLEMENTARY_SHEET_NAME = `A${"𐐨".repeat(4)}`;
const LONG_FILE_STEM = `${"A".repeat(79)}𐐀`;

async function createWorkspace(
  page: import("@playwright/test").Page,
  format: "duckdb" | "sqlite",
): Promise<void> {
  await page.goto("/tools/db");
  await page.getByTestId("workspace-new-format").selectOption(format);
  await page.getByTestId("workspace-new-name").fill(`unicode.${format}`);
  await page.getByTestId("workspace-new").click();
  await expect(page.getByTestId("workspace-summary")).toBeVisible();
}

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format} imports a sheet whose generated record prefix crosses a UTF-16 boundary`, async ({
    page,
  }) => {
    await createWorkspace(page, format);
    const workbook = await createWorkbookUpload("unicode-sheet.xlsx", [
      {
        name: SUPPLEMENTARY_SHEET_NAME,
        rows: [["Value"], ["Synthetic"]],
      },
    ]);

    await page.getByTestId("workspace-import-input").setInputFiles(workbook);
    await page.getByTestId("workspace-import-role").fill("inventory");
    await page.getByTestId("workspace-import-revision").fill("Iteration 1");
    await page.getByTestId("workspace-import-prepare").click();
    await expect(page.getByTestId("workspace-import-review")).toBeVisible();
    await expect(page.getByTestId("workspace-import-region")).toContainText(
      SUPPLEMENTARY_SHEET_NAME,
    );

    await page.getByTestId("workspace-import-resolve").click();
    await expect(page.getByTestId("workspace-import-review")).toContainText(
      "ready",
    );
    await page.getByTestId("workspace-import-apply").click();
    await expect(page.getByTestId("workspace-import-result")).toContainText(
      "Added 1 row",
    );
    await expect(page.getByTestId("workspace-table")).toContainText(
      SUPPLEMENTARY_SHEET_NAME,
    );
  });

  test(`${format} opens a default working-copy name without splitting a supplementary character`, async ({
    page,
  }, testInfo) => {
    const fileName = `${LONG_FILE_STEM}.${format}`;
    const databasePath = testInfo.outputPath(fileName);
    const created = await createDatabase({ path: databasePath, format });
    await created.database.close();

    await page.goto("/tools/db");
    await page.getByTestId("workspace-open-input").setInputFiles(databasePath);
    await expect(page.getByTestId("workspace-summary")).toBeVisible();

    const workingCopyName = await page
      .getByTestId("workspace-summary")
      .getByRole("heading", { level: 2 })
      .textContent();
    expect(workingCopyName).toMatch(
      new RegExp(`^${"A".repeat(79)}-[0-9a-f-]{36}\\.${format}$`, "u"),
    );
    expect(workingCopyName).not.toContain("�");
  });
}
