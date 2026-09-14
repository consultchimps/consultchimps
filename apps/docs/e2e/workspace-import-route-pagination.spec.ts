import { expect, test } from "@playwright/test";

import { createWorkbookUpload } from "./fixtures";

test("pages batch routes without losing reviewed decisions", async ({
  page,
}) => {
  const workbook = await createWorkbookUpload(
    "many-sheets.xlsx",
    Array.from({ length: 51 }, (_, index) => ({
      name: `Sheet ${String(index + 1).padStart(2, "0")}`,
      rows: [["Value"], [`Row ${String(index + 1)}`]],
    })),
  );

  await page.goto("/tools/db");
  await page.getByTestId("workspace-new-format").selectOption("sqlite");
  await page.getByTestId("workspace-new-name").fill("route-pages.sqlite");
  await page.getByTestId("workspace-new").click();
  await page.getByTestId("workspace-import-input").setInputFiles([workbook]);
  await page.getByTestId("workspace-import-role").fill("sheets");
  await page.getByTestId("workspace-import-revision").fill("Iteration 1");
  await page.getByTestId("workspace-import-prepare").click();

  await expect(page.getByTestId("workspace-import-route-count")).toHaveText(
    "Showing 50 of 51 routes",
  );
  const firstTable = page.getByTestId("workspace-import-table").first();
  const originalFirstTable = await firstTable.inputValue();
  await firstTable.fill(`${originalFirstTable}_edited`);
  await expect(page.getByTestId("workspace-import-routes-next")).toBeDisabled();
  await firstTable.fill(originalFirstTable);
  await expect(page.getByTestId("workspace-import-routes-next")).toBeEnabled();

  await page.getByTestId("workspace-import-routes-next").click();
  await expect(page.getByTestId("workspace-import-route-count")).toHaveText(
    "Showing 1 of 51 routes",
  );
  await expect(page.getByTestId("workspace-import-review")).toContainText(
    "Sheet 51",
  );
  const secondPageTable = page.getByTestId("workspace-import-table");
  await secondPageTable.fill("Reviewed_sheet_51");
  await expect(
    page.getByTestId("workspace-import-routes-first"),
  ).toBeDisabled();
  await page.getByTestId("workspace-import-resolve").click();

  await expect(page.getByTestId("workspace-import-route-count")).toHaveText(
    "Showing 50 of 51 routes",
  );
  await expect(page.getByTestId("workspace-import-routes-next")).toBeEnabled();
  await page.getByTestId("workspace-import-routes-next").click();
  await expect(page.getByTestId("workspace-import-route-count")).toHaveText(
    "Showing 1 of 51 routes",
  );
  await expect(page.getByTestId("workspace-import-table")).toHaveValue(
    "Reviewed_sheet_51",
  );
});
