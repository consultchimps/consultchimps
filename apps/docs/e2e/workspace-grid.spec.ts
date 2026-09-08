import { Database, type TableSchema } from "@consultchimps/db";
import { expect, test, type Locator, type Page } from "@playwright/test";
import { readFile } from "node:fs/promises";

/**
 * The record grid on /workspace: view a table, edit it, have an impossible
 * value refused, and find the surviving edits in the saved file.
 *
 * The workspace under test is built here with `@consultchimps/db` rather than
 * imported through the page, so this spec exercises the grid and nothing else:
 * it does not wait on an import surface, and a change to import cannot make it
 * fail. The bytes it produces are an ordinary workspace file, which is what the
 * page's own file input opens.
 *
 * Everything runs through the download and file-input fallback, for the reason
 * the shell spec gives: the File System Access API's pickers cannot be driven
 * from a test, and the fallback is the path every browser without the API uses.
 */
async function forceDownloadFallback(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const globals = window as unknown as Record<string, unknown>;
    delete globals["showOpenFilePicker"];
    delete globals["showSaveFilePicker"];
  });
}

const region: TableSchema = {
  name: "Region",
  columns: [{ name: "name", type: "text", nullable: false }],
  foreignKeys: [],
  recordId: { prefix: "REG", padding: 4 },
};

const customer: TableSchema = {
  name: "Customer",
  columns: [
    { name: "name", type: "text", nullable: false },
    { name: "region", type: "text" },
    { name: "headcount", type: "integer" },
    { name: "active", type: "boolean" },
  ],
  foreignKeys: [{ column: "region", referencesTable: "Region" }],
  recordId: { prefix: "CUST", padding: 4 },
};

/** Two related tables with a few neutral records, as saved workspace bytes. */
async function workspaceFixture(): Promise<Buffer> {
  const database = await Database.create();
  database.createTable(region);
  database.createTable(customer);
  database.insertRecord("Region", { name: "North" });
  database.insertRecord("Region", { name: "South" });
  database.insertRecord("Customer", {
    name: "Acme",
    region: "REG-0001",
    headcount: 12,
    active: true,
  });
  database.insertRecord("Customer", {
    name: "Globex",
    region: "REG-0002",
    headcount: 7,
    active: false,
  });
  const bytes = Buffer.from(database.serialize());
  database.close();
  return bytes;
}

async function openWorkspace(page: Page, bytes: Buffer): Promise<void> {
  await page.getByTestId("file-input").setInputFiles({
    name: "records.sqlite",
    mimeType: "application/vnd.sqlite3",
    buffer: bytes,
  });
  await expect(page.getByTestId("workspace-notice")).toHaveText(
    "Opened the workspace",
  );
  await expect(page.getByTestId("workspace-grid-section")).toBeVisible();
}

/** A cell addressed the way the grid keys its rows: by Record ID. */
function cellOf(page: Page, recordId: string, field: string): Locator {
  return page.locator(
    `[data-record-id="${recordId}"] [tabulator-field="${field}"]`,
  );
}

/**
 * Open a cell's editor and hand back its input.
 *
 * The click is retried rather than made once. Tabulator renders rows into the
 * DOM as they are needed and re-lays the columns out after the first paint, so
 * a click can land on a cell element that is being replaced, which opens no
 * editor and reports no error. Retrying until the editor is up tests what a
 * visitor experiences (clicking a cell edits it) without pinning the test to
 * the grid's internal render timing.
 */
async function openEditor(cell: Locator): Promise<Locator> {
  const input = cell.locator("input");
  await expect(async () => {
    await cell.click();
    await expect(input).toBeVisible({ timeout: 1000 });
  }).toPass({ timeout: 20_000 });
  return input;
}

/** Type into a cell and commit with Enter, the way a visitor would. */
async function typeInCell(
  page: Page,
  recordId: string,
  field: string,
  value: string,
): Promise<void> {
  const input = await openEditor(cellOf(page, recordId, field));
  await input.fill(value);
  await input.press("Enter");
}

async function downloadedWorkspace(page: Page): Promise<Buffer> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("workspace-save-as").click();
  const download = await downloadPromise;
  return readFile(await download.path());
}

test.describe("/workspace record grid", () => {
  test("shows a table and lets a visitor switch between tables", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    // Tables come back ordered by name, so Customer is the one first shown.
    await expect(page.getByTestId("workspace-table-select")).toHaveValue(
      "Customer",
    );
    await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Acme");
    await expect(cellOf(page, "CUST-0002", "name")).toHaveText("Globex");
    await expect(cellOf(page, "CUST-0001", "headcount")).toHaveText("12");
    // A foreign key shows the referenced record's label and stores its Record ID.
    await expect(cellOf(page, "CUST-0001", "region")).toHaveText(
      "North (REG-0001)",
    );

    await page.getByTestId("workspace-table-select").selectOption("Region");
    await expect(cellOf(page, "REG-0001", "name")).toHaveText("North");
    await expect(cellOf(page, "CUST-0001", "name")).toHaveCount(0);
  });

  test("never offers an editor for the Record ID", async ({ page }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    const recordId = cellOf(page, "CUST-0001", "record_id");
    await expect(recordId).toHaveText("CUST-0001");
    await recordId.click();
    await expect(recordId.locator("input")).toHaveCount(0);

    // Tab from an editable cell lands on the next editable one, never on a
    // Record ID.
    await openEditor(cellOf(page, "CUST-0001", "name"));
    await page.keyboard.press("Tab");
    await expect(
      cellOf(page, "CUST-0001", "region").locator("input"),
    ).toBeVisible();
  });

  test("persists cell edits into the saved workspace", async ({ page }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    await typeInCell(page, "CUST-0001", "name", "Acme Holdings");
    await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Acme Holdings");

    await typeInCell(page, "CUST-0001", "headcount", "15");
    await expect(cellOf(page, "CUST-0001", "headcount")).toHaveText("15");

    // The foreign-key picker shows labels and stores the Record ID behind them.
    await openEditor(cellOf(page, "CUST-0001", "region"));
    await page
      .locator(".tabulator-edit-list-item", { hasText: "South (REG-0002)" })
      .click();
    await expect(cellOf(page, "CUST-0001", "region")).toHaveText(
      "South (REG-0002)",
    );

    // The neighbouring record is untouched: an edit is written by Record ID,
    // not by row position.
    await expect(cellOf(page, "CUST-0002", "name")).toHaveText("Globex");
    await expect(cellOf(page, "CUST-0002", "headcount")).toHaveText("7");

    const saved = await downloadedWorkspace(page);
    await openWorkspace(page, saved);

    await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Acme Holdings");
    await expect(cellOf(page, "CUST-0001", "headcount")).toHaveText("15");
    await expect(cellOf(page, "CUST-0001", "region")).toHaveText(
      "South (REG-0002)",
    );
  });

  test("reverts a value the database refuses and says why", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    await typeInCell(page, "CUST-0001", "headcount", "15");
    await expect(cellOf(page, "CUST-0001", "headcount")).toHaveText("15");

    // A whole-number column cannot hold 1.5. Nothing in the page decides that:
    // the library refuses the write, and the grid puts the stored value back.
    await typeInCell(page, "CUST-0001", "headcount", "1.5");
    await expect(page.getByTestId("workspace-error")).toContainText(
      "not a whole number",
    );
    await expect(page.getByTestId("workspace-error")).toContainText(
      'column "headcount"',
    );
    await expect(cellOf(page, "CUST-0001", "headcount")).toHaveText("15");

    // Emptying a column the schema declares non-nullable is refused the same
    // way, and leaves the stored value in place.
    await typeInCell(page, "CUST-0001", "name", "");
    await expect(page.getByTestId("workspace-error")).toContainText(
      "was left empty",
    );
    await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Acme");

    // The refusals reached the page, not the file: the saved workspace still
    // holds the value that was accepted.
    const saved = await downloadedWorkspace(page);
    await openWorkspace(page, saved);
    await expect(cellOf(page, "CUST-0001", "headcount")).toHaveText("15");
    await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Acme");
  });
});
