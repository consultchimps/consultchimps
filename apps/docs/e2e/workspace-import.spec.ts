import { expect, test, type Page } from "@playwright/test";
import { readFile } from "node:fs/promises";

import { createWorkbookUpload, type UploadFile } from "./fixtures";

/**
 * Importing a workbook and a `.csv` into the data workspace.
 *
 * The shell's own spec covers create, save, and reopen; this one covers what a
 * workspace is for. A two-sheet workbook exercises the sheet picker (one sheet
 * chosen, one left behind), the `.csv` exercises the single-table path, and the
 * table listing is the acceptance signal for both: until the grid lands it is
 * the only place a person sees that their rows arrived, with the Record ID
 * prefix and the column types the import inferred.
 *
 * The pickers are removed before the page loads, as in the shell's spec, so
 * saving downloads a copy and the saved bytes can be reopened here.
 */
async function forceDownloadFallback(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const globals = window as unknown as Record<string, unknown>;
    delete globals["showOpenFilePicker"];
    delete globals["showSaveFilePicker"];
  });
}

/** A workbook with two worksheets, so the picker has a real choice to make. */
function customerWorkbook(): Promise<UploadFile> {
  return createWorkbookUpload("book.xlsx", [
    {
      name: "Customers",
      rows: [
        ["Customer", "Region", "Active", "Score", "Opened"],
        ["Acme", "North", "true", 12.5, "2026-01-31"],
        ["Beta", "South", "false", 8, "2026-02-01"],
      ],
    },
    {
      name: "Regions",
      rows: [
        ["Region", "Manager"],
        ["North", "Team A"],
      ],
    },
  ]);
}

/**
 * A `.csv` whose values are chosen to pin the inference rules down: a padded
 * reference code stays text, a whole number becomes an integer, and a date
 * written the ISO way becomes a date.
 */
const ORDERS_CSV: UploadFile = {
  name: "orders.csv",
  mimeType: "text/csv",
  buffer: Buffer.from(
    'Reference,Quantity,Placed,Note\n007,3,2026-03-01,"Two lines\nin one field"\n008,11,2026-03-02,plain\n',
    "utf8",
  ),
};

function importRow(page: Page, index: number) {
  return page.getByTestId("workspace-import-source").nth(index);
}

test.describe("/workspace import", () => {
  test("imports a chosen worksheet and a csv, then keeps them across a save", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();
    await expect(page.getByTestId("workspace-tables-empty")).toBeVisible();

    // 1. The workbook: both worksheets are offered, and only one is imported.
    await page
      .getByTestId("workspace-import-input")
      .setInputFiles(await customerWorkbook());
    await expect(page.getByTestId("workspace-import-form")).toBeVisible();
    await expect(page.getByTestId("workspace-import-source")).toHaveCount(2);
    await expect(
      page.getByTestId("workspace-import-source-name").first(),
    ).toHaveText("Customers");
    // The suggested name and prefix are visible before anything is created.
    await expect(
      importRow(page, 0).getByTestId("workspace-import-name"),
    ).toHaveValue("Customers");
    await expect(
      importRow(page, 0).getByTestId("workspace-import-prefix"),
    ).toHaveValue("CUST");

    await importRow(page, 1).getByTestId("workspace-import-selected").uncheck();
    await importRow(page, 0).getByTestId("workspace-import-prefix").fill("CUS");
    await page.getByTestId("workspace-import-run").click();

    await expect(page.getByTestId("workspace-notice")).toHaveText(
      "Imported 1 table with 2 rows",
    );
    await expect(page.getByTestId("workspace-table")).toHaveCount(1);
    await expect(page.getByTestId("workspace-table-name")).toHaveText(
      "Customers",
    );
    await expect(page.getByTestId("workspace-table-rows")).toHaveText("2 rows");
    await expect(page.getByTestId("workspace-table-prefix")).toHaveText("CUS");
    // Every column took the type its own values agree on.
    await expect(page.getByTestId("workspace-table-columns")).toHaveText(
      "Customer (text), Region (text), Active (boolean), Score (real), Opened (date)",
    );

    // 2. The csv: one table, and the inference rules that matter most.
    await page.getByTestId("workspace-import-input").setInputFiles(ORDERS_CSV);
    await expect(page.getByTestId("workspace-import-source")).toHaveCount(1);
    await expect(
      importRow(page, 0).getByTestId("workspace-import-name"),
    ).toHaveValue("orders");
    await importRow(page, 0)
      .getByTestId("workspace-import-name")
      .fill("Orders");
    await importRow(page, 0).getByTestId("workspace-import-prefix").fill("ORD");
    await page.getByTestId("workspace-import-run").click();

    await expect(page.getByTestId("workspace-notice")).toHaveText(
      "Imported 1 table with 2 rows",
    );
    await expect(page.getByTestId("workspace-table")).toHaveCount(2);
    await expect(page.getByTestId("workspace-table-count")).toHaveText("2");
    await expect(
      page
        .getByTestId("workspace-table")
        .nth(1)
        .getByTestId("workspace-table-columns"),
      // A padded reference stays text; a plain count is an integer; an ISO date
      // is a date; a quoted field with a line break in it is ordinary text.
    ).toHaveText(
      "Reference (text), Quantity (integer), Placed (date), Note (text)",
    );

    // 3. Save the workspace and reopen exactly those bytes.
    const downloadPromise = page.waitForEvent("download");
    await page.getByTestId("workspace-save-as").click();
    const download = await downloadPromise;
    const bytes = await readFile(await download.path());

    await page.getByTestId("file-input").setInputFiles({
      name: "reopened.sqlite",
      mimeType: "application/vnd.sqlite3",
      buffer: bytes,
    });

    await expect(page.getByTestId("workspace-notice")).toHaveText(
      "Opened the workspace",
    );
    await expect(page.getByTestId("workspace-table")).toHaveCount(2);
    await expect(page.getByTestId("workspace-table-name").first()).toHaveText(
      "Customers",
    );
    await expect(page.getByTestId("workspace-table-rows").first()).toHaveText(
      "2 rows",
    );
    await expect(page.getByTestId("workspace-table-prefix").first()).toHaveText(
      "CUS",
    );
    await expect(page.getByTestId("workspace-table-columns").last()).toHaveText(
      "Reference (text), Quantity (integer), Placed (date), Note (text)",
    );
  });

  test("refuses a name the workspace already holds, before anything is written", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();

    await page
      .getByTestId("workspace-import-input")
      .setInputFiles(await customerWorkbook());
    await importRow(page, 1).getByTestId("workspace-import-selected").uncheck();
    await page.getByTestId("workspace-import-run").click();
    await expect(page.getByTestId("workspace-table")).toHaveCount(1);

    // The same worksheet again, under a name that differs only by case: the
    // page reads it as the name the workspace already holds and says so, with
    // the Import button held back rather than a failure after the fact.
    await page
      .getByTestId("workspace-import-input")
      .setInputFiles(await customerWorkbook());
    await importRow(page, 1).getByTestId("workspace-import-selected").uncheck();
    await importRow(page, 0)
      .getByTestId("workspace-import-name")
      .fill("customers");

    await expect(page.getByTestId("workspace-import-problem")).toContainText(
      "the same name",
    );
    await expect(page.getByTestId("workspace-import-run")).toBeDisabled();
    await expect(page.getByTestId("workspace-table")).toHaveCount(1);
  });

  test("reports a file that holds nothing to import", async ({ page }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();

    // A header line and nothing under it: there is no data to make a table
    // from, so the page says that rather than creating an empty one.
    await page.getByTestId("workspace-import-input").setInputFiles({
      name: "empty.xlsx",
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      buffer: Buffer.from("Customer,Region\n", "utf8"),
    });

    await expect(page.getByTestId("workspace-import-error")).toContainText(
      "has a header row with rows under it, so there is nothing to import",
    );
    await expect(page.getByTestId("workspace-import-form")).toHaveCount(0);
    await expect(page.getByTestId("workspace-tables-empty")).toBeVisible();
  });
});
