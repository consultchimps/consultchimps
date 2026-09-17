import { expect, test } from "@playwright/test";

import { createWorkbookUpload } from "./fixtures";

/**
 * The row browser is a bounded, read-only look at one table: it opens from a
 * table card, pages in record ID order, reloads after an applied import
 * batch, and closes with the database.
 */
test("browses stored rows one page at a time", async ({ page }) => {
  test.setTimeout(120_000);
  await page.goto("/tools/db");
  await page.getByTestId("workspace-new-name").fill("browse.sqlite");
  await page.getByTestId("workspace-new").click();
  await expect(page.getByTestId("workspace-summary")).toBeVisible();

  const rows = Array.from({ length: 55 }, (_, index) => [
    `Reading ${String(index + 1)}`,
    String((index + 1) * 10),
  ]);
  await page
    .getByTestId("workspace-import-input")
    .setInputFiles(
      await createWorkbookUpload("readings.xlsx", [
        { name: "Readings", rows: [["Label", "Amount"], ...rows] },
      ]),
    );
  await page.getByTestId("workspace-import-prepare").click();
  // Reading a workbook and staging its rows is slow on a loaded machine.
  await expect(page.getByTestId("workspace-import-review")).toBeVisible({
    timeout: 90_000,
  });
  await page.getByTestId("workspace-import-resolve").click();
  await expect(page.getByTestId("workspace-import-review")).toContainText(
    "ready",
  );
  await page.getByTestId("workspace-import-apply").click();
  await expect(page.getByTestId("workspace-import-result")).toContainText(
    "Added 55 rows",
  );

  // Browse from the table card.
  await expect(page.getByTestId("workspace-table")).toHaveCount(1);
  await page.getByTestId("workspace-table-browse").click();
  const browser = page.getByTestId("workspace-browse");
  await expect(browser).toContainText("Rows in Readings");
  await expect(page.getByTestId("workspace-browse-row")).toHaveCount(50);
  await expect(page.getByTestId("workspace-browse-count")).toHaveText(
    "Showing 50 rows, more follow",
  );
  await expect(browser.locator("th").first()).toContainText("record_id");
  await expect(browser).toContainText("Reading 1");
  // The source column carries the capture ID of the workbook, not its name.
  await expect(browser).toContainText("SRC-");
  await expect(page.getByTestId("workspace-browse-first")).toHaveCount(0);

  // Page forward, then back to the first page.
  await page.getByTestId("workspace-browse-next").click();
  await expect(page.getByTestId("workspace-browse-row")).toHaveCount(5);
  await expect(page.getByTestId("workspace-browse-count")).toHaveText(
    "Showing 5 rows",
  );
  await expect(page.getByTestId("workspace-browse-next")).toHaveCount(0);
  await expect(browser).toContainText("Reading 55");
  await page.getByTestId("workspace-browse-first").click();
  await expect(page.getByTestId("workspace-browse-row")).toHaveCount(50);
  await expect(page.getByTestId("workspace-browse-first")).toHaveCount(0);

  // Nothing here writes: the summary count is unchanged.
  await expect(page.getByTestId("workspace-table")).toContainText("55 rows");

  await page.getByTestId("workspace-browse-close").click();
  await expect(browser).toHaveCount(0);

  // A different database closes the panel on its own.
  await page.getByTestId("workspace-table-browse").click();
  await expect(browser).toBeVisible();
  await page.getByTestId("workspace-new-name").fill("other.sqlite");
  await page.getByTestId("workspace-new").click();
  await expect(page.getByTestId("workspace-summary")).toContainText(
    "other.sqlite",
  );
  await expect(browser).toHaveCount(0);
});
