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
 *
 * The last two tests cover the shell's unsaved-changes guard rather than import
 * itself. They live here because import is the first command that can leave a
 * workspace holding work no file has: the guard is the shell's, and the record
 * grid will lean on the same flag.
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

/**
 * Hold the worker's import commands for a while, so the page can be inspected
 * while one is genuinely in flight.
 *
 * A real import of a real file finishes in milliseconds, which is too short to
 * assert anything about reliably; making the fixture big enough to be slow
 * would trade a timing assumption for a slower one. Delaying the command on its
 * way to the worker is the same kind of stand-in as removing the file pickers
 * above: the page is untouched, and what is being tested is exactly what the
 * page does while a command has not come back.
 */
async function delayWorkerImports(page: Page, ms: number): Promise<void> {
  await page.addInitScript((delay: number) => {
    const post = Worker.prototype.postMessage;
    Worker.prototype.postMessage = function (
      this: Worker,
      message: unknown,
      transfer?: unknown,
    ): void {
      const type = (message as { type?: string } | null)?.type;
      if (type === "describeImport" || type === "import") {
        window.setTimeout(() => {
          (post as (m: unknown, t?: unknown) => void).call(
            this,
            message,
            transfer,
          );
        }, delay);
        return;
      }
      (post as (m: unknown, t?: unknown) => void).call(this, message, transfer);
    } as typeof Worker.prototype.postMessage;
  }, ms);
}

/**
 * A workbook whose totals column holds formulas nothing has calculated. Excel
 * writes a formula and its last result together; a file written by a generator,
 * or saved with calculation off, carries the formula alone.
 */
function uncalculatedWorkbook(): Promise<UploadFile> {
  return createWorkbookUpload("totals.xlsx", [
    {
      name: "Customers",
      rows: [
        // The last header is a formula too, so a header that reads as blank and
        // would import under an invented name is caught with the rest.
        ["Customer", "Region", { formula: 'CONCATENATE("Sc","ore")' }],
        ["Acme", "North", { formula: "10+2" }],
        ["Beta", "South", { formula: "3+5" }],
      ],
    },
  ]);
}

/**
 * A workbook whose dates are held the way Excel holds them: a count of days
 * from the workbook's epoch, wearing a date number format. Nothing in the file
 * spells the date out, so what reaches the workspace is what the reader made
 * of the number.
 */
function dateSerialWorkbook(): Promise<UploadFile> {
  return createWorkbookUpload("opened.xlsx", [
    {
      name: "Customers",
      rows: [
        ["Customer", "Opened"],
        // 1 January 2024, and the day after it.
        ["Acme", { serial: 45292 }],
        ["Beta", { serial: 45293 }],
      ],
    },
  ]);
}

/**
 * A workbook whose dates declare themselves dates and write ISO 8601 text,
 * wearing no number format at all. Nothing but the declaration says they are
 * dates, which is the case a reader that judges by format alone gets wrong.
 */
function declaredDateWorkbook(): Promise<UploadFile> {
  return createWorkbookUpload("declared.xlsx", [
    {
      name: "Customers",
      rows: [
        ["Customer", "Opened"],
        ["Acme", { date: "2024-01-01" }],
        ["Beta", { date: "2024-01-02" }],
      ],
    },
  ]);
}

/**
 * A workbook whose amounts are errors. A reader built on a spreadsheet engine
 * sees the internal code Excel numbers each error by, so `#REF!` would import
 * as 23 and `#DIV/0!` as 7, and the column would infer a numeric type.
 */
function errorValueWorkbook(): Promise<UploadFile> {
  return createWorkbookUpload("amounts.xlsx", [
    {
      name: "Customers",
      rows: [
        ["Customer", "Region", "Amount"],
        ["Acme", "North", { error: "#REF!" }],
        ["Beta", "South", { error: "#DIV/0!" }],
      ],
    },
  ]);
}

/**
 * Press Back and wait for the page to have seen it. A same-document traversal
 * changes nothing a navigation assertion could wait on, so the press is only
 * observable through the event it fires.
 */
async function pressBack(page: Page): Promise<void> {
  await page.evaluate(
    () =>
      new Promise<void>((resolve) => {
        window.addEventListener("popstate", () => resolve(), { once: true });
        window.history.back();
      }),
  );
}

function importRow(page: Page, index: number) {
  return page.getByTestId("workspace-import-source").nth(index);
}

/** Import the workbook's first worksheet and leave the second behind. */
async function importCustomersSheet(page: Page): Promise<void> {
  await page
    .getByTestId("workspace-import-input")
    .setInputFiles(await customerWorkbook());
  await importRow(page, 1).getByTestId("workspace-import-selected").uncheck();
  await page.getByTestId("workspace-import-run").click();
  await expect(page.getByTestId("workspace-table")).toHaveCount(1);
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
    // The workspace now holds work that no file has.
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible();
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

    // Saving cleared the flag, so reopening asks nothing.
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);

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

  test("will not replace an imported workspace until the loss is confirmed", async ({
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
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible();

    // New asks first, and the imported table is still there while it asks.
    await page.getByTestId("workspace-new").click();
    await expect(page.getByTestId("workspace-confirm")).toBeVisible();
    await expect(page.getByTestId("workspace-table")).toHaveCount(1);

    // Keeping the workspace leaves it exactly as it was.
    await page.getByTestId("workspace-confirm-cancel").click();
    await expect(page.getByTestId("workspace-confirm")).toHaveCount(0);
    await expect(page.getByTestId("workspace-table")).toHaveCount(1);
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible();

    // Open asks the same question, and discarding then replaces the workspace.
    await page.getByTestId("workspace-open").click();
    await expect(page.getByTestId("workspace-confirm")).toBeVisible();
    await page.getByTestId("workspace-confirm-cancel").click();

    await page.getByTestId("workspace-new").click();
    await page.getByTestId("workspace-confirm-discard").click();
    await expect(page.getByTestId("workspace-tables-empty")).toBeVisible();
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);
  });

  test("asks nothing once the imported workspace has been saved", async ({
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
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible();

    const downloadPromise = page.waitForEvent("download");
    await page.getByTestId("workspace-save-as").click();
    await downloadPromise;
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);

    // Nothing is at risk any more, so New replaces the workspace at once.
    await page.getByTestId("workspace-new").click();
    await expect(page.getByTestId("workspace-tables-empty")).toBeVisible();
    await expect(page.getByTestId("workspace-confirm")).toHaveCount(0);
  });

  test("holds the rest of the page while an import is in flight", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await delayWorkerImports(page, 1200);
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();
    await expect(page.getByTestId("workspace-new")).toBeEnabled();

    // Reading the file is a page command too, so the same buttons are held.
    await page
      .getByTestId("workspace-import-input")
      .setInputFiles(await customerWorkbook());
    await expect(page.getByTestId("workspace-new")).toBeDisabled();
    await expect(page.getByTestId("workspace-open")).toBeDisabled();
    await expect(page.getByTestId("workspace-save")).toBeDisabled();
    await expect(page.getByTestId("workspace-import-form")).toBeVisible();
    await expect(page.getByTestId("workspace-new")).toBeEnabled();

    // The import itself. Nothing may replace the workspace while it runs: a
    // click that got through here would land behind the import and discard
    // the tables it was still creating.
    await importRow(page, 1).getByTestId("workspace-import-selected").uncheck();
    await page.getByTestId("workspace-import-run").click();
    await expect(page.getByTestId("workspace-new")).toBeDisabled();
    await expect(page.getByTestId("workspace-open")).toBeDisabled();
    await expect(page.getByTestId("workspace-save")).toBeDisabled();
    await expect(page.getByTestId("workspace-save-as")).toBeDisabled();
    await expect(page.getByTestId("workspace-import-run")).toBeDisabled();

    // Once it lands the page is usable again, and the workspace is dirty.
    await expect(page.getByTestId("workspace-notice")).toHaveText(
      "Imported 1 table with 2 rows",
    );
    await expect(page.getByTestId("workspace-new")).toBeEnabled();
    await expect(page.getByTestId("workspace-open")).toBeEnabled();
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible();
    await expect(page.getByTestId("workspace-table")).toHaveCount(1);
  });

  test("refuses a worksheet whose formulas were never calculated", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();

    await page
      .getByTestId("workspace-import-input")
      .setInputFiles(await uncalculatedWorkbook());

    // The worksheet is listed rather than hidden, so the visitor learns why it
    // cannot be imported instead of wondering where it went, and the tick is
    // refused before any choice is made.
    await expect(page.getByTestId("workspace-import-source")).toHaveCount(1);
    await expect(page.getByTestId("workspace-import-blocked")).toContainText(
      "3 cells hold a formula this workbook carries no calculated value for",
    );
    await expect(page.getByTestId("workspace-import-blocked")).toContainText(
      "Open the workbook in Excel, let it calculate, save it",
    );
    // The tick is refused, so the worksheet can never be chosen, and with
    // nothing else in the workbook there is nothing left to import.
    await expect(
      importRow(page, 0).getByTestId("workspace-import-selected"),
    ).toBeDisabled();
    await expect(page.getByTestId("workspace-import-run")).toBeDisabled();
    await expect(page.getByTestId("workspace-import-problem")).toHaveText(
      "Choose at least one table to import",
    );

    // Nothing was created, and the workspace is untouched.
    await expect(page.getByTestId("workspace-tables-empty")).toBeVisible();
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);
  });

  test.describe("in a browser east of UTC", () => {
    // The page runs in UTC+4 for this block. A reader that turned a serial into
    // a date through the browser's own zone would produce 31 December here for
    // a cell that says 1 January, so the import runs where that would show.
    test.use({ timezoneId: "Asia/Dubai" });

    test("imports a date a worksheet declares without a format", async ({
      page,
    }) => {
      await forceDownloadFallback(page);
      await page.goto("/workspace");
      await page.getByTestId("workspace-new").click();

      await page
        .getByTestId("workspace-import-input")
        .setInputFiles(await declaredDateWorkbook());
      await page.getByTestId("workspace-import-run").click();
      await expect(page.getByTestId("workspace-table")).toHaveCount(1);

      // The column can only take the date type if the reader saw the cell
      // declare itself one. Judging by the number format alone, which is what
      // it used to do, made this an integer column holding two serials.
      await expect(page.getByTestId("workspace-table-columns")).toHaveText(
        "Customer (text), Opened (date)",
      );
      await expect(page.getByTestId("workspace-table-rows")).toHaveText(
        "2 rows",
      );
    });

    test("imports a date Excel holds as a number", async ({ page }) => {
      await forceDownloadFallback(page);
      await page.goto("/workspace");
      await page.getByTestId("workspace-new").click();

      await page
        .getByTestId("workspace-import-input")
        .setInputFiles(await dateSerialWorkbook());
      await page.getByTestId("workspace-import-run").click();
      await expect(page.getByTestId("workspace-table")).toHaveCount(1);

      // Nothing in the file spells the dates out, so the column can only take
      // the date type if the reader turned both serials into ISO 8601 text
      // that the database accepts as a date. A serial that arrived as a number
      // would have made this an integer column.
      await expect(page.getByTestId("workspace-table-columns")).toHaveText(
        "Customer (text), Opened (date)",
      );
      await expect(page.getByTestId("workspace-table-rows")).toHaveText(
        "2 rows",
      );
    });
  });

  test("refuses a worksheet holding error values", async ({ page }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();

    await page
      .getByTestId("workspace-import-input")
      .setInputFiles(await errorValueWorkbook());

    // Listed rather than hidden, for the same reason: an error cell reads as an
    // ordinary number, so importing it would report a clean success and put
    // numbers nobody entered in the table.
    await expect(page.getByTestId("workspace-import-source")).toHaveCount(1);
    await expect(page.getByTestId("workspace-import-blocked")).toContainText(
      "2 cells hold an error value",
    );
    await expect(page.getByTestId("workspace-import-blocked")).toContainText(
      "Fix or clear the errors in Excel, save the workbook",
    );
    await expect(
      importRow(page, 0).getByTestId("workspace-import-selected"),
    ).toBeDisabled();
    await expect(page.getByTestId("workspace-import-run")).toBeDisabled();

    // Nothing was created, and the workspace is untouched.
    await expect(page.getByTestId("workspace-tables-empty")).toBeVisible();
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);
  });

  test("holds a link out of the page until the loss is confirmed", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();
    await importCustomersSheet(page);
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible();

    // A client-side transition unloads nothing, so the browser's own warning
    // never fires; the page has to hold the click itself.
    await page.getByTestId("guide-link").click();
    await expect(page.getByTestId("workspace-confirm")).toBeVisible();
    await expect(page).toHaveURL(/\/workspace$/u);
    await expect(page.getByTestId("workspace-table")).toHaveCount(1);

    // Keeping the workspace leaves the page exactly where it was.
    await page.getByTestId("workspace-confirm-cancel").click();
    await expect(page.getByTestId("workspace-confirm")).toHaveCount(0);
    await expect(page).toHaveURL(/\/workspace$/u);
    await expect(page.getByTestId("workspace-table")).toHaveCount(1);

    // Discarding follows the link that was held.
    await page.getByTestId("guide-link").click();
    await page.getByTestId("workspace-confirm-discard").click();
    await expect(page).toHaveURL(/\/docs\/libraries/u);
  });

  test("follows a link at once when nothing is unsaved", async ({ page }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();
    await expect(page.getByTestId("workspace-summary")).toBeVisible();

    await page.getByTestId("guide-link").click();
    await expect(page).toHaveURL(/\/docs\/libraries/u);
    await expect(page.getByTestId("workspace-confirm")).toHaveCount(0);
  });

  test("holds the Back button while the workspace is unsaved", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/tools");
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();
    await importCustomersSheet(page);

    // Back cannot be cancelled, so the page keeps a spare history entry to
    // absorb the first press. Driven through history rather than Playwright's
    // goBack, which waits for a navigation that deliberately does not happen.
    await page.evaluate(() => {
      window.history.back();
    });
    await expect(page.getByTestId("workspace-confirm")).toBeVisible();
    await expect(page).toHaveURL(/\/workspace$/u);

    await page.getByTestId("workspace-confirm-discard").click();
    await expect(page).toHaveURL(/\/tools$/u);
  });

  test("holds a link out of the page while an import is still running", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await delayWorkerImports(page, 1500);
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();
    await expect(page.getByTestId("workspace-new")).toBeEnabled();

    await page
      .getByTestId("workspace-import-input")
      .setInputFiles(await customerWorkbook());
    await expect(page.getByTestId("workspace-import-form")).toBeVisible();
    await importRow(page, 1).getByTestId("workspace-import-selected").uncheck();
    // Nothing is unsaved yet: what is at stake is the import itself, which
    // exists nowhere but this tab until it lands.
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);
    await page.getByTestId("workspace-import-run").click();

    await page.getByTestId("guide-link").click();
    await expect(page.getByTestId("workspace-confirm")).toContainText(
      "An import is still running",
    );
    await expect(page).toHaveURL(/\/workspace$/u);

    // The click was held rather than obeyed, so the worker was never torn down
    // and the import finished.
    await expect(page.getByTestId("workspace-notice")).toHaveText(
      "Imported 1 table with 2 rows",
    );
    await expect(page.getByTestId("workspace-table")).toHaveCount(1);
  });

  test("arms the Back guard again after a save has spent it", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/tools");
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();
    await importCustomersSheet(page);

    const downloadPromise = page.waitForEvent("download");
    await page.getByTestId("workspace-save-as").click();
    await downloadPromise;
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);

    // Nothing is at stake, so this press is simply spent. The page still has to
    // notice that it went, or it will believe it is still holding the spare.
    await pressBack(page);
    await expect(page.getByTestId("workspace-confirm")).toHaveCount(0);
    await expect(page).toHaveURL(/\/workspace$/u);

    // A second import has to arm the guard again rather than trust a flag left
    // over from the first.
    await page.getByTestId("workspace-import-input").setInputFiles(ORDERS_CSV);
    await importRow(page, 0)
      .getByTestId("workspace-import-name")
      .fill("Orders");
    await page.getByTestId("workspace-import-run").click();
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible();

    await pressBack(page);
    await expect(page.getByTestId("workspace-confirm")).toBeVisible();
    await expect(page).toHaveURL(/\/workspace$/u);
  });

  test("names a one-worksheet workbook after the file it came from", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();

    // One worksheet, and its name says nothing to the visitor: the file they
    // chose is the name they know.
    await page
      .getByTestId("workspace-import-input")
      .setInputFiles(
        await createWorkbookUpload("customers.xlsx", [
          { name: "Sheet1", rows: [["Customer"], ["Acme"]] },
        ]),
      );
    await expect(
      importRow(page, 0).getByTestId("workspace-import-name"),
    ).toHaveValue("customers");
    await expect(
      importRow(page, 0).getByTestId("workspace-import-prefix"),
    ).toHaveValue("CUST");

    // More than one worksheet, and the worksheet names are what tell them
    // apart, so those are used instead.
    await page
      .getByTestId("workspace-import-input")
      .setInputFiles(await customerWorkbook());
    await expect(
      importRow(page, 0).getByTestId("workspace-import-name"),
    ).toHaveValue("Customers");
    await expect(
      importRow(page, 1).getByTestId("workspace-import-name"),
    ).toHaveValue("Regions");
  });

  test("imports a file the browser named, not the file name", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();

    // No extension on the name, and a media type the browser supplied. The
    // picker accepted this before and the reader then refused it, because they
    // decided what the file was in two different ways.
    await page.getByTestId("workspace-import-input").setInputFiles({
      name: "report",
      mimeType: "text/csv",
      buffer: Buffer.from("Customer,Region\nAcme,North\n", "utf8"),
    });

    await expect(page.getByTestId("workspace-import-form")).toBeVisible();
    await expect(
      importRow(page, 0).getByTestId("workspace-import-name"),
    ).toHaveValue("report");
    await page.getByTestId("workspace-import-run").click();

    await expect(page.getByTestId("workspace-notice")).toHaveText(
      "Imported 1 table with 1 row",
    );
    await expect(page.getByTestId("workspace-table-columns")).toHaveText(
      "Customer (text), Region (text)",
    );
  });

  test("leaves one entry behind when a held link is followed", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/tools");
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();
    await importCustomersSheet(page);

    // The guard armed a spare entry for this page. Following the link has to
    // take its place rather than stack on top of it, or Back reaches a second
    // copy of the workspace before it reaches the page before it.
    await page.getByTestId("guide-link").click();
    await page.getByTestId("workspace-confirm-discard").click();
    await expect(page).toHaveURL(/\/docs\/libraries/u);

    await page.goBack();
    await expect(page).toHaveURL(/\/workspace$/u);
    await page.goBack();
    await expect(page).toHaveURL(/\/tools$/u);
  });

  test("leaves one entry behind when a link is followed after saving", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/tools");
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();
    await importCustomersSheet(page);

    const downloadPromise = page.waitForEvent("download");
    await page.getByTestId("workspace-save-as").click();
    await downloadPromise;
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);

    // Saving clears the flag but cannot remove the entry the change armed, so
    // the clean way out has to retire it too.
    await page.getByTestId("guide-link").click();
    await expect(page).toHaveURL(/\/docs\/libraries/u);

    await page.goBack();
    await expect(page).toHaveURL(/\/workspace$/u);
    await page.goBack();
    await expect(page).toHaveURL(/\/tools$/u);
  });

  test("reports a workbook that holds nothing to import", async ({ page }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();

    // A real workbook whose worksheet holds no rows: there is nothing to make
    // a table from, so the page says that rather than creating an empty one.
    await page
      .getByTestId("workspace-import-input")
      .setInputFiles(
        await createWorkbookUpload("empty.xlsx", [
          { name: "Sheet1", rows: [] },
        ]),
      );

    await expect(page.getByTestId("workspace-import-error")).toContainText(
      "has a header row with rows under it, so there is nothing to import",
    );
    await expect(page.getByTestId("workspace-import-form")).toHaveCount(0);
    await expect(page.getByTestId("workspace-tables-empty")).toBeVisible();
  });

  test("reports a file that is not a readable workbook", async ({ page }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await page.getByTestId("workspace-new").click();

    // Named like a workbook, but not an OOXML package at all.
    await page.getByTestId("workspace-import-input").setInputFiles({
      name: "broken.xlsx",
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      buffer: Buffer.from("Customer,Region\n", "utf8"),
    });

    await expect(page.getByTestId("workspace-import-error")).toContainText(
      "Could not read workbook",
    );
    await expect(page.getByTestId("workspace-import-form")).toHaveCount(0);
    await expect(page.getByTestId("workspace-tables-empty")).toBeVisible();
  });
});
