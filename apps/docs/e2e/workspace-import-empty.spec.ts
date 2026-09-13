import { expect, test } from "@playwright/test";

import { createWorkbookUpload } from "./fixtures";

test("hidden-only workbooks do not create an empty import plan", async ({
  page,
}) => {
  const workbook = await createWorkbookUpload("hidden-only.xlsx", [
    {
      name: "Hidden data",
      state: "hidden",
      rows: [
        ["Region", "Value"],
        ["North", "100"],
      ],
    },
  ]);
  await page.goto("/tools/db");
  await page.getByTestId("workspace-new-format").selectOption("sqlite");
  await page.getByTestId("workspace-new-name").fill("empty-import.sqlite");
  await page.getByTestId("workspace-new").click();
  await expect(page.getByTestId("workspace-summary")).toBeVisible();

  await page.getByTestId("workspace-import-input").setInputFiles(workbook);
  await expect(page.getByTestId("workspace-import-source")).toHaveCount(1);
  await page.getByTestId("workspace-import-prepare").click();

  await expect(page.getByTestId("workspace-error")).toContainText(
    "Choose at least one source region before importing",
  );
  await expect(page.getByTestId("workspace-error")).toContainText(
    "include hidden sheets and try again",
  );
  await expect(page.getByTestId("workspace-import-review")).toHaveCount(0);
  await expect(page.getByTestId("workspace-import-result")).toHaveCount(0);
  await expect(page.getByTestId("workspace-delivery-count")).toHaveText("0");

  await page.reload();
  await page.getByTestId("workspace-reopen").first().click();
  await expect(page.getByTestId("workspace-import-resume")).toHaveCount(0);
  await expect(page.getByTestId("workspace-delivery-count")).toHaveText("0");
});
