import { Database, type TableSchema } from "@consultchimps/db";
import { expect, test, type Locator, type Page } from "@playwright/test";
import { readFile } from "node:fs/promises";

import { createWorkbookUpload, type UploadFile } from "./fixtures";

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

/**
 * A table with a foreign key back to itself, whose first ordinary text column
 * is what names a record. Editing that column changes what every cell pointing
 * at the record shows, which no other table's edit can do.
 */
const employee: TableSchema = {
  name: "Employee",
  columns: [
    { name: "name", type: "text", nullable: false },
    { name: "manager", type: "text" },
  ],
  foreignKeys: [{ column: "manager", referencesTable: "Employee" }],
  recordId: { prefix: "EMP", padding: 4 },
};

const customer: TableSchema = {
  name: "Customer",
  columns: [
    { name: "name", type: "text", nullable: false },
    { name: "region", type: "text" },
    { name: "headcount", type: "integer" },
    { name: "active", type: "boolean" },
    // A legal column name with a dot in it. The schema accepts it, so the grid
    // has to address it literally rather than as a path into nested data.
    { name: "billing.address", type: "text" },
  ],
  foreignKeys: [{ column: "region", referencesTable: "Region" }],
  recordId: { prefix: "CUST", padding: 4 },
};

/** Two related tables with a few neutral records, as saved workspace bytes. */
async function workspaceFixture(customerName = "Acme"): Promise<Buffer> {
  const database = await Database.create();
  database.createTable(region);
  database.createTable(customer);
  database.createTable(employee);
  database.insertRecord("Region", { name: "North" });
  database.insertRecord("Region", { name: "South" });
  database.insertRecord("Employee", { name: "Ada", manager: null });
  database.insertRecord("Employee", { name: "Grace", manager: "EMP-0001" });
  database.insertRecord("Customer", {
    name: customerName,
    region: "REG-0001",
    headcount: 12,
    active: true,
    "billing.address": "1 North Street",
  });
  database.insertRecord("Customer", {
    name: "Globex",
    region: "REG-0002",
    headcount: 7,
    active: false,
    "billing.address": "2 South Street",
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
 * Do something to a cell, from the click that opens its editor through to the
 * keystroke that commits, retrying the whole interaction rather than any one
 * step of it.
 *
 * Tabulator renders rows into the DOM as they are needed and re-lays its
 * columns out when the table is resized, which moves the element under a click
 * and closes an editor that happens to be open. That is the grid's own render
 * timing, not behaviour worth asserting, and a test that pins itself to one
 * uninterrupted interaction is testing the timing instead of the feature.
 *
 * Every step therefore carries a short timeout of its own, so an action that
 * catches the grid mid-layout gives up quickly and the whole interaction is
 * tried again, rather than one waiting action holding the test open until it
 * times out. Every `interaction` must be safe to repeat: typing the same value
 * twice commits once, because the second attempt changes nothing.
 */
const STEP_TIMEOUT = 2_000;

async function editCell(
  cell: Locator,
  interaction: (input: Locator) => Promise<void>,
): Promise<void> {
  const input = cell.locator("input");
  await expect(async () => {
    await cell.click({ timeout: STEP_TIMEOUT });
    await expect(input).toBeVisible({ timeout: STEP_TIMEOUT });
    await interaction(input);
  }).toPass({ timeout: 30_000 });
}

/**
 * Open a cell's editor and type into it, leaving it open and uncommitted. The
 * caller decides what commits it, which for these tests is a click somewhere
 * else on the page.
 */
async function typeWithoutCommitting(
  cell: Locator,
  value: string,
): Promise<void> {
  const input = cell.locator("input");
  await expect(async () => {
    await cell.click({ timeout: STEP_TIMEOUT });
    await expect(input).toBeVisible({ timeout: STEP_TIMEOUT });
    await input.fill(value, { timeout: STEP_TIMEOUT });
  }).toPass({ timeout: 30_000 });
}

/** Type into a cell and commit with Enter, the way a visitor would. */
async function typeInCell(
  page: Page,
  recordId: string,
  field: string,
  value: string,
): Promise<void> {
  await editCell(cellOf(page, recordId, field), async (input) => {
    await input.fill(value, { timeout: STEP_TIMEOUT });
    await input.press("Enter", { timeout: STEP_TIMEOUT });
  });
}

/**
 * A workbook to import, so a table can appear after the grid is already up. The
 * file and its worksheet are named alike, because what the import calls the
 * table it creates is its own business and this spec is not the place to pin
 * that rule down.
 */
function supplierWorkbook(): Promise<UploadFile> {
  return createWorkbookUpload("Supplier.xlsx", [
    {
      name: "Supplier",
      rows: [
        ["Name", "Region"],
        ["Acme Supply", "North"],
      ],
    },
  ]);
}

/**
 * Hold named worker commands on their way out, so the page can be inspected
 * while one is genuinely in flight.
 *
 * The same stand-in the import spec uses, and for the same reason. An import of
 * a real file finishes in milliseconds and a cell edit finishes in less: both
 * are answered before React has rendered the question that was raised about
 * them, so a test that waits for the page to settle sees only the state after.
 * Making the work slow enough to observe would trade one timing assumption for
 * a worse one. Delaying the command on its way to the worker leaves the page
 * untouched, and what is being tested is exactly what the page does while a
 * command has not come back.
 */
async function delayWorkerCommands(
  page: Page,
  types: readonly string[],
  ms: number,
): Promise<void> {
  await page.addInitScript(
    ({ delay, held }: { delay: number; held: readonly string[] }) => {
      const post = Worker.prototype.postMessage;
      Worker.prototype.postMessage = function (
        this: Worker,
        message: unknown,
        transfer?: unknown,
      ): void {
        const type = (message as { type?: string } | null)?.type;
        if (type !== undefined && held.includes(type)) {
          window.setTimeout(() => {
            (post as (m: unknown, t?: unknown) => void).call(
              this,
              message,
              transfer,
            );
          }, delay);
          return;
        }
        (post as (m: unknown, t?: unknown) => void).call(
          this,
          message,
          transfer,
        );
      } as typeof Worker.prototype.postMessage;
    },
    { delay: ms, held: types },
  );
}

/** Press Back the way the shell's guard sees it, with nothing blurred first. */
async function pressBack(page: Page): Promise<void> {
  await page.evaluate(
    () =>
      new Promise<void>((resolve) => {
        window.addEventListener("popstate", () => resolve(), { once: true });
        window.history.back();
      }),
  );
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

  test("addresses a column name with a dot in it literally", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    // Tabulator reads a dot in a field as a path into nested data unless it is
    // told otherwise, which would render this column blank and send its edits
    // to a property nobody asked for. The schema accepts the name, so the grid
    // has to show and write it as the column it is.
    await expect(cellOf(page, "CUST-0001", "billing.address")).toHaveText(
      "1 North Street",
    );

    await typeInCell(page, "CUST-0001", "billing.address", "9 East Street");
    await expect(cellOf(page, "CUST-0001", "billing.address")).toHaveText(
      "9 East Street",
    );
    await expect(cellOf(page, "CUST-0002", "billing.address")).toHaveText(
      "2 South Street",
    );

    const saved = await downloadedWorkspace(page);
    await openWorkspace(page, saved);
    await expect(cellOf(page, "CUST-0001", "billing.address")).toHaveText(
      "9 East Street",
    );
  });

  test("replaces the grid when another workspace is opened", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture("Acme"));
    await typeInCell(page, "CUST-0001", "headcount", "15");
    await expect(cellOf(page, "CUST-0001", "headcount")).toHaveText("15");

    // And a refusal standing over that workspace, about CUST-0001's headcount.
    await typeInCell(page, "CUST-0001", "headcount", "1.5");
    await expect(page.getByTestId("workspace-grid-error")).toContainText(
      "not a whole number",
    );

    // A second workspace holds the same table and the same Record IDs, which is
    // exactly why an edit is bound to the workspace it was made in. The grid is
    // replaced with the new one, and nothing from the first survives in it:
    // neither the values, nor an explanation whose every key names something
    // that exists here and means something else.
    await openWorkspace(page, await workspaceFixture("Globex"));
    await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Globex");
    await expect(cellOf(page, "CUST-0001", "headcount")).toHaveText("12");
    await expect(page.getByTestId("workspace-grid-error")).toHaveCount(0);
    await expect(page.getByTestId("workspace-error")).toHaveCount(0);
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
    await editCell(cellOf(page, "CUST-0001", "name"), async () => {
      await page.keyboard.press("Tab");
      await expect(
        cellOf(page, "CUST-0001", "region").locator("input"),
      ).toBeVisible({ timeout: STEP_TIMEOUT });
    });
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
    await editCell(cellOf(page, "CUST-0001", "region"), async () => {
      await page
        .locator(".tabulator-edit-list-item", { hasText: "South (REG-0002)" })
        .click({ timeout: STEP_TIMEOUT });
    });
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
    await expect(page.getByTestId("workspace-grid-error")).toContainText(
      "not a whole number",
    );
    await expect(page.getByTestId("workspace-grid-error")).toContainText(
      'column "headcount"',
    );
    await expect(cellOf(page, "CUST-0001", "headcount")).toHaveText("15");

    // Emptying a column the schema declares non-nullable is refused the same
    // way, and leaves the stored value in place. Both refusals stand: the
    // newest in full, the other counted.
    await typeInCell(page, "CUST-0001", "name", "");
    await expect(page.getByTestId("workspace-grid-error")).toContainText(
      "was left empty",
    );
    await expect(page.getByTestId("workspace-grid-error")).toContainText(
      "1 other edit was refused as well",
    );
    await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Acme");

    // Putting one of them right takes its explanation away and brings the other
    // back in full, because each belongs to the cell it named.
    await typeInCell(page, "CUST-0001", "name", "Acme Two");
    await expect(page.getByTestId("workspace-grid-error")).toContainText(
      "not a whole number",
    );
    await expect(page.getByTestId("workspace-grid-error")).not.toContainText(
      "other edit was refused",
    );

    // And the visitor can put the rest away.
    await page.getByTestId("workspace-grid-error-dismiss").click();
    await expect(page.getByTestId("workspace-grid-error")).toHaveCount(0);

    // The refusals reached the page, not the file: what the saved workspace
    // holds is the values that were accepted, the 15 and the name that replaced
    // the empty one, and neither of the two that were refused.
    const saved = await downloadedWorkspace(page);
    await openWorkspace(page, saved);
    await expect(cellOf(page, "CUST-0001", "headcount")).toHaveText("15");
    await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Acme Two");
  });

  test("reports an edit to the shell, so leaving asks first", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    // Nothing is unsaved until an edit lands, so New goes straight through.
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);

    await typeInCell(page, "CUST-0001", "name", "Acme Holdings");
    // The grid keeps no flag of its own: this is the shell's, set by the one
    // marker the grid calls when the worker accepts an edit.
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible();

    await page.getByTestId("workspace-new").click();
    await expect(page.getByTestId("workspace-confirm")).toBeVisible();

    // Cancelling leaves the workspace and the edit exactly where they were.
    await page.getByTestId("workspace-confirm-cancel").click();
    await expect(page.getByTestId("workspace-confirm")).toHaveCount(0);
    await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Acme Holdings");
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible();
  });

  test("asks nothing after an edit that has been saved", async ({ page }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    await typeInCell(page, "CUST-0001", "headcount", "15");
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible();

    // Saving is what clears the flag, and it is cleared in the shell, not by
    // the grid noticing anything.
    await downloadedWorkspace(page);
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);

    await page.getByTestId("workspace-new").click();
    await expect(page.getByTestId("workspace-confirm")).toHaveCount(0);
    await expect(page.getByTestId("workspace-tables-empty")).toBeVisible();
  });

  test("locks editing while an import is in flight", async ({ page }) => {
    await forceDownloadFallback(page);
    await delayWorkerCommands(page, ["describeImport", "import"], 2_000);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    await page
      .getByTestId("workspace-import-input")
      .setInputFiles(await supplierWorkbook());
    await expect(page.getByTestId("workspace-import-form")).toBeVisible();
    await page.getByTestId("workspace-import-run").click();

    // The import is now in the worker's queue. Editing is the shell's single
    // busy state away, so the cell offers no editor at all: the grid holds no
    // second opinion about whether the workspace is free.
    const cell = cellOf(page, "CUST-0001", "name");
    await cell.click();
    await expect(cell.locator("input")).toHaveCount(0);
    await expect(page.getByTestId("workspace-table-select")).toBeDisabled();

    // Once it lands, editing is offered again.
    await expect(page.getByTestId("workspace-table")).toHaveCount(4, {
      timeout: 30_000,
    });
    await typeInCell(page, "CUST-0001", "name", "Acme Holdings");
    await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Acme Holdings");
  });

  test("shows a table imported after the grid was already up", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    // The grid is showing Customer, and knows only the tables the file held.
    await expect(page.getByTestId("workspace-table-select")).toHaveValue(
      "Customer",
    );
    await expect(
      page.getByTestId("workspace-table-select").locator("option"),
    ).toHaveCount(3);

    await page
      .getByTestId("workspace-import-input")
      .setInputFiles(await supplierWorkbook());
    await expect(page.getByTestId("workspace-import-form")).toBeVisible();
    await page.getByTestId("workspace-import-run").click();
    await expect(page.getByTestId("workspace-table")).toHaveCount(4);

    // The switcher reads the shell's summary, which the import replaced, so the
    // new table is there without the grid asking anyone for a second listing.
    await expect(
      page.getByTestId("workspace-table-select").locator("option"),
    ).toHaveCount(4);
    await page.getByTestId("workspace-table-select").selectOption("Supplier");

    // Located by column rather than by Record ID: what the import numbers the
    // rows is its own business, and this test is about the table arriving.
    // Scoped to a row rather than the whole grid: the column header carries the
    // same field attribute, and only rows carry a Record ID.
    const supplierName = page.locator(
      '[data-record-id] [tabulator-field="Name"]',
    );
    await expect(supplierName).toHaveText("Acme Supply");

    // And the grid is editing the workspace as it now stands: the read that
    // followed the import carries the generation the worker will accept, so
    // this edit is not refused as belonging to the workspace before it.
    await editCell(supplierName, async (input) => {
      await input.fill("Acme Supplies", { timeout: STEP_TIMEOUT });
      await input.press("Enter", { timeout: STEP_TIMEOUT });
    });
    await expect(supplierName).toHaveText("Acme Supplies");
    await expect(page.getByTestId("workspace-grid-error")).toHaveCount(0);
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible();
  });

  test("holds New for an edit that has been sent but not answered", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    // Held so the window under test is open long enough to look at. Without
    // this the edit is answered before the question has rendered, and the page
    // would be asserted in the state after it rather than during it.
    await delayWorkerCommands(page, ["updateCell"], 2_000);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    // The editor commits when it loses focus, so this click is both what
    // commits the edit and what asks to replace the workspace. Between the two
    // the workspace still reads as clean, and that is the window being tested:
    // without the shell counting the edit, New would go through, the edit would
    // land on the database being discarded, and nothing would say so.
    await typeWithoutCommitting(cellOf(page, "CUST-0001", "name"), "Acme Two");
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);
    await page.getByTestId("workspace-new").click();

    await expect(page.getByTestId("workspace-confirm")).toBeVisible();
    // And it says what is actually at stake. Nothing is unsaved yet, because
    // the edit has not been answered, so a question explaining this state as
    // anything else would be describing a workspace the visitor does not have.
    await expect(page.getByTestId("workspace-confirm")).toContainText(
      "An edit is still being applied",
    );
    await expect(page.getByTestId("workspace-confirm")).not.toContainText(
      "An import is still running",
    );
    // The workspace is untouched behind the question.
    await expect(page.getByTestId("workspace-file-name")).toHaveText(
      "records.sqlite",
    );

    await page.getByTestId("workspace-confirm-cancel").click();
    await expect(page.getByTestId("workspace-confirm")).toHaveCount(0);
    // And the edit it committed on the way is in the workspace, not lost to it.
    await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Acme Two");
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible({
      timeout: 30_000,
    });

    // Answering it the other way is still available, and still discards.
    await page.getByTestId("workspace-new").click();
    await page.getByTestId("workspace-confirm-discard").click();
    await expect(page.getByTestId("workspace-tables-empty")).toBeVisible();
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);
  });

  test("holds Open for an edit that has been sent but not answered", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await delayWorkerCommands(page, ["updateCell"], 2_000);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    // An accepted edit first, so the workspace is already unsaved when the
    // second one is still on its way: two reasons at once, which is the case a
    // question that names only one of them gets wrong.
    await typeInCell(page, "CUST-0001", "name", "Acme Two");
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible({
      timeout: 30_000,
    });

    await typeWithoutCommitting(cellOf(page, "CUST-0001", "headcount"), "21");
    await page.getByTestId("workspace-open").click();

    await expect(page.getByTestId("workspace-confirm")).toBeVisible();
    await expect(page.getByTestId("workspace-confirm")).toContainText(
      "have not been saved to a file and an edit is still being applied",
    );

    await page.getByTestId("workspace-confirm-cancel").click();
    await expect(cellOf(page, "CUST-0001", "headcount")).toHaveText("21");
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible();
  });

  test("names the edit alone once it has landed", async ({ page }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    // No delay this time, so the edit is answered before the question renders
    // and the workspace really is unsaved by then. The wording follows the
    // state it is shown over rather than the one that raised it, which is what
    // deriving it from the same answer buys.
    await typeInCell(page, "CUST-0001", "name", "Acme Two");
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible();
    await page.getByTestId("workspace-new").click();
    await expect(page.getByTestId("workspace-confirm")).toContainText(
      "This workspace has changes that have not been saved to a file.",
    );
    await expect(page.getByTestId("workspace-confirm")).not.toContainText(
      "an edit is still being applied",
    );
  });

  test("renames a record everywhere it is referred to, without a reload", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());
    await page.getByTestId("workspace-table-select").selectOption("Employee");

    // The manager column points back at this same table, so what it shows is
    // read from the rows on screen rather than from a listing taken when the
    // table was read.
    await expect(cellOf(page, "EMP-0002", "manager")).toHaveText(
      "Ada (EMP-0001)",
    );

    // Renaming the record the other row points at is the case a captured
    // listing gets wrong: the worker accepts and stores it, nothing else on
    // screen changes, and a plain edit does not move the generation, so nothing
    // would ever re-read it.
    await typeInCell(page, "EMP-0001", "name", "Ada Lovelace");
    await expect(cellOf(page, "EMP-0001", "name")).toHaveText("Ada Lovelace");
    await expect(cellOf(page, "EMP-0002", "manager")).toHaveText(
      "Ada Lovelace (EMP-0001)",
    );

    // And the picker offers the new name, because it and the cell read the one
    // source rather than two that can drift apart.
    await editCell(cellOf(page, "EMP-0002", "manager"), async () => {
      await expect(
        page.locator(".tabulator-edit-list-item", {
          hasText: "Ada Lovelace (EMP-0001)",
        }),
      ).toBeVisible({ timeout: STEP_TIMEOUT });
      await page.keyboard.press("Escape");
    });

    // It is the workspace that changed, not just the screen.
    const saved = await downloadedWorkspace(page);
    await openWorkspace(page, saved);
    await page.getByTestId("workspace-table-select").selectOption("Employee");
    await expect(cellOf(page, "EMP-0002", "manager")).toHaveText(
      "Ada Lovelace (EMP-0001)",
    );
  });

  test("leaves a reference to another table alone when this one is edited", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    // Customer's region points at Region, which is not on screen and so cannot
    // change while this table is: its labels are read once and hold.
    await expect(cellOf(page, "CUST-0001", "region")).toHaveText(
      "North (REG-0001)",
    );
    await typeInCell(page, "CUST-0001", "name", "Acme Holdings");
    await expect(cellOf(page, "CUST-0001", "region")).toHaveText(
      "North (REG-0001)",
    );
    await expect(cellOf(page, "CUST-0002", "region")).toHaveText(
      "South (REG-0002)",
    );
  });

  test("re-reads a reference to another table after an import", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    // Rename a region, which is the label source for Customer's region column,
    // while Region is the table on screen.
    await page.getByTestId("workspace-table-select").selectOption("Region");
    await typeInCell(page, "REG-0001", "name", "Northern");
    await expect(cellOf(page, "REG-0001", "name")).toHaveText("Northern");

    // Switching back re-reads Customer, so its snapshot of Region is the one
    // taken now.
    await page.getByTestId("workspace-table-select").selectOption("Customer");
    await expect(cellOf(page, "CUST-0001", "region")).toHaveText(
      "Northern (REG-0001)",
    );

    // An import moves the generation, which re-reads the table on screen and
    // with it the labels it holds for tables that are not.
    await page
      .getByTestId("workspace-import-input")
      .setInputFiles(await supplierWorkbook());
    await expect(page.getByTestId("workspace-import-form")).toBeVisible();
    await page.getByTestId("workspace-import-run").click();
    await expect(page.getByTestId("workspace-table")).toHaveCount(4);
    await expect(cellOf(page, "CUST-0001", "region")).toHaveText(
      "Northern (REG-0001)",
    );
  });

  test("keeps a refused edit explained when another one succeeds", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    // Held so both edits are in flight together, which is the only way their
    // replies can arrive in an order that matters.
    await delayWorkerCommands(page, ["updateCell"], 1_500);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    // An invalid value first, then a valid one in a different cell. The second
    // is accepted and the first refused, and it is the second that answers
    // last: with one slot for the explanation, the acceptance would clear the
    // refusal and leave a cell snapped back with nothing saying why.
    await typeInCell(page, "CUST-0001", "headcount", "1.5");
    await typeInCell(page, "CUST-0002", "name", "Globex Two");

    await expect(page.getByTestId("workspace-unsaved")).toBeVisible({
      timeout: 30_000,
    });
    await expect(cellOf(page, "CUST-0002", "name")).toHaveText("Globex Two");
    await expect(cellOf(page, "CUST-0001", "headcount")).toHaveText("12");
    await expect(page.getByTestId("workspace-grid-error")).toContainText(
      "not a whole number",
    );

    // It survives the save too, since nothing about saving answers it.
    await downloadedWorkspace(page);
    await expect(page.getByTestId("workspace-grid-error")).toContainText(
      "not a whole number",
    );
  });

  test("holds the page while a cell is open for editing", async ({ page }) => {
    await forceDownloadFallback(page);
    // Somewhere to come from, so the shell has an entry it can guard.
    await page.goto("/tools");
    await page.getByRole("link", { name: "Workspace", exact: true }).click();
    await openWorkspace(page, await workspaceFixture());

    // Nothing sent and nothing unsaved: what is at stake is only what is in the
    // input element, and until now nothing outside the grid knew it was there.
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);
    const beforeEditing = await page.evaluate(() => window.history.length);

    await typeWithoutCommitting(cellOf(page, "CUST-0001", "name"), "Acme Two");

    // The shell arms a spare history entry whenever it is holding, from the
    // same rendered state that installs the browser's own warning before the
    // tab closes or reloads. That warning is the guard this is really about and
    // the one no test can open, so the entry is how its presence is asserted:
    // an editor being open is now enough for the shell to be holding.
    await expect
      .poll(() => page.evaluate(() => window.history.length))
      .toBe(beforeEditing + 1);
  });

  test("holds a link out of the page for work an editor still holds", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    await typeWithoutCommitting(cellOf(page, "CUST-0001", "name"), "Acme Two");
    await page.getByTestId("guide-link").click();

    // Held, and the page not left. Which reason it names depends on how far the
    // work got: clicking the link blurs the editor, which commits the draft, so
    // by the time the click is handled this is an edit on its way rather than
    // one still being typed. Either way it is the same hold, asked once.
    await expect(page.getByTestId("workspace-confirm")).toBeVisible();
    await expect(page).toHaveURL(/\/workspace$/u);

    await page.getByTestId("workspace-confirm-cancel").click();
    await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Acme Two");
  });

  test("holds the Back button for a cell that is still open for editing", async ({
    page,
  }) => {
    // #174. A Back press is the one way of leaving that never blurs the editor,
    // so the draft is still in the input element when the question is raised.
    // The confirmation used to lock the grid, the lock cancelled that editor,
    // and the cancel then dismissed the question the press had raised: the
    // draft went with no warning. A question is answered by the visitor, never
    // by a side effect of the navigation that raised it.
    await forceDownloadFallback(page);
    await page.goto("/tools");
    await page.getByRole("link", { name: "Workspace", exact: true }).click();
    await openWorkspace(page, await workspaceFixture());

    const beforeEditing = await page.evaluate(() => window.history.length);
    await typeWithoutCommitting(cellOf(page, "CUST-0001", "name"), "Acme Two");
    // The spare entry the press has to land on, which is also how the page says
    // it is holding for the editor. Waited for, so the press below is one of
    // the page's own rather than a press it never armed for.
    await expect
      .poll(() => page.evaluate(() => window.history.length))
      .toBe(beforeEditing + 1);

    await pressBack(page);

    await expect(page.getByTestId("workspace-confirm")).toBeVisible();
    await expect(page.getByTestId("workspace-confirm")).toContainText(
      "A cell is still open for editing",
    );
    await expect(page).toHaveURL(/\/workspace$/u);

    // The editor is still open behind the question, still holding the draft.
    // This is the assertion the bug was: the question used to lock the grid,
    // and the lock cancelled this editor and reverted what is in it.
    await expect(
      cellOf(page, "CUST-0001", "name").locator("input"),
    ).toHaveValue("Acme Two");

    // Keeping the workspace keeps the draft. The click that answers is also
    // what blurs the editor, and Tabulator commits on blur, so by the time the
    // question is answered the draft is an edit on its way rather than one
    // still being typed. Either way it is in the workspace and not lost to it.
    await page.getByTestId("workspace-confirm-cancel").click();
    await expect(page.getByTestId("workspace-confirm")).toHaveCount(0);
    await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Acme Two");
    await expect(page.getByTestId("workspace-unsaved")).toBeVisible();
  });

  test("leaves when the visitor discards the draft it asked about", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/tools");
    await page.getByRole("link", { name: "Workspace", exact: true }).click();
    await openWorkspace(page, await workspaceFixture());

    const beforeEditing = await page.evaluate(() => window.history.length);
    await typeWithoutCommitting(cellOf(page, "CUST-0001", "name"), "Acme Two");
    await expect
      .poll(() => page.evaluate(() => window.history.length))
      .toBe(beforeEditing + 1);

    await pressBack(page);
    await expect(page.getByTestId("workspace-confirm")).toBeVisible();

    // Answered the other way, the press is honoured: back over the spare entry
    // and this page together, to the page the visitor came from. Nothing warns
    // a second time, because answering released the whole hold rather than the
    // unsaved flag alone.
    await page.getByTestId("workspace-confirm-discard").click();
    await expect(page).toHaveURL(/\/tools$/u);
  });

  test("stops holding once the editor is closed with Escape", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    // Opened and abandoned. Tabulator reports the cancel like any other, so the
    // hold is released through the one handler and nothing is left waiting on a
    // draft that no longer exists.
    await typeWithoutCommitting(cellOf(page, "CUST-0001", "name"), "Acme Two");
    await page.keyboard.press("Escape");
    await expect(
      cellOf(page, "CUST-0001", "name").locator("input"),
    ).toHaveCount(0);
    await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Acme");
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);

    // With nothing at stake the page is left without a question, which is what
    // says the hold was released rather than merely hidden.
    await page.getByTestId("guide-link").click();
    await expect(page).toHaveURL(/\/docs\/libraries/u);
    await expect(page.getByTestId("workspace-confirm")).toHaveCount(0);
  });
});
