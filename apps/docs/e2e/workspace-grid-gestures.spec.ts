import { Database, type TableSchema } from "@consultchimps/db";
import { expect, test, type Locator, type Page } from "@playwright/test";
import { readFile } from "node:fs/promises";

/**
 * The record grid's spreadsheet gestures on /workspace: selecting a range,
 * copying it as the text Excel reads, pasting a block back, and dragging the
 * fill handle.
 *
 * Its own spec rather than more of `workspace-grid.spec.ts`, because these are
 * one feature with one rule behind them: a gesture is one movement, so it is one
 * command, planned whole and refused whole. What each test pins is the outcome a
 * visitor sees, and the values that reach the saved file.
 *
 * The clipboard is driven by dispatching the browser's own copy and paste events
 * with a `DataTransfer` on them, rather than by the system clipboard, which
 * needs permissions the static export never asks for and which would make the
 * test about the operating system. The events are the same ones a real copy and
 * paste deliver, and the page's listeners are what answer them.
 */
async function forceDownloadFallback(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const globals = window as unknown as Record<string, unknown>;
    delete globals["showOpenFilePicker"];
    delete globals["showSaveFilePicker"];
  });
}

/**
 * A table with one of each kind a gesture treats differently: text ending in a
 * number, plain text, a whole number, and a date.
 */
const task: TableSchema = {
  name: "Task",
  columns: [
    { name: "code", type: "text", nullable: false },
    { name: "owner", type: "text", nullable: false },
    { name: "effort", type: "integer" },
    { name: "due", type: "date" },
  ],
  foreignKeys: [],
  recordId: { prefix: "TSK", padding: 4 },
};

/** Five neutral records, as saved workspace bytes. */
async function workspaceFixture(): Promise<Buffer> {
  const database = await Database.create();
  database.createTable(task);
  database.insertRecord("Task", {
    code: "T-007",
    owner: "North",
    effort: 10,
    due: "2026-01-01",
  });
  database.insertRecord("Task", {
    code: "B-1",
    owner: "South",
    effort: 20,
    due: "2026-01-08",
  });
  for (const [index, owner] of ["East", "West", "North"].entries()) {
    database.insertRecord("Task", {
      code: `X-${String(index + 1)}`,
      owner,
      effort: index + 1,
      due: null,
    });
  }
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

/** Select one rectangle: click a corner, shift-click the other. */
async function selectRange(
  page: Page,
  from: { recordId: string; field: string },
  to?: { recordId: string; field: string },
): Promise<void> {
  await cellOf(page, from.recordId, from.field).click();
  if (to !== undefined) {
    await cellOf(page, to.recordId, to.field).click({ modifiers: ["Shift"] });
  }
}

/** What a copy would put on the clipboard, as the page writes it. */
function copySelection(page: Page): Promise<string> {
  return page.evaluate(() => {
    const holder = document.querySelector(".tabulator-tableholder");
    if (holder === null) {
      return "";
    }
    const data = new DataTransfer();
    holder.dispatchEvent(
      new ClipboardEvent("copy", {
        bubbles: true,
        cancelable: true,
        clipboardData: data,
      }),
    );
    return data.getData("text/plain");
  });
}

/** Paste text into the grid, exactly as a spreadsheet would deliver it. */
async function pasteText(page: Page, text: string): Promise<void> {
  await page.evaluate((clipboard) => {
    const holder = document.querySelector(".tabulator-tableholder");
    if (holder === null) {
      return;
    }
    const data = new DataTransfer();
    data.setData("text/plain", clipboard);
    holder.dispatchEvent(
      new ClipboardEvent("paste", {
        bubbles: true,
        cancelable: true,
        clipboardData: data,
      }),
    );
  }, text);
}

/** Drag the fill handle onto a cell, the way a pointer does it. */
async function dragFillHandleTo(
  page: Page,
  recordId: string,
  field: string,
): Promise<void> {
  const handle = page.getByTestId("workspace-fill-handle");
  await expect(handle).toBeVisible();
  const from = await handle.boundingBox();
  const to = await cellOf(page, recordId, field).boundingBox();
  expect(from).not.toBeNull();
  expect(to).not.toBeNull();
  if (from === null || to === null) {
    return;
  }
  await page.mouse.move(from.x + from.width / 2, from.y + from.height / 2);
  await page.mouse.down();
  await page.mouse.move(to.x + to.width / 2, to.y + to.height / 2, {
    steps: 10,
  });
  await page.mouse.up();
}

/**
 * Make the page send a workspace number the worker does not hold, so the guard
 * every gesture carries is exercised end to end.
 *
 * A real workspace cannot be replaced underneath a gesture from the page: the
 * client runs one command at a time, so a New or an Open is not even posted
 * until the gesture has been answered. The number the command carries is what
 * the guard reads, so that is what this changes, and everything else (the page,
 * the worker, the refusal, the report) is the real path.
 */
async function sendStaleGeneration(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const post = Worker.prototype.postMessage;
    Worker.prototype.postMessage = function (
      this: Worker,
      message: unknown,
      transfer?: unknown,
    ): void {
      const command = message as { type?: string; generation?: number } | null;
      if (
        command?.type === "updateCells" &&
        typeof command.generation === "number"
      ) {
        command.generation += 1;
      }
      (post as (m: unknown, t?: unknown) => void).call(this, message, transfer);
    } as typeof Worker.prototype.postMessage;
  });
}

async function downloadedWorkspace(page: Page): Promise<Buffer> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("workspace-save-as").click();
  const download = await downloadPromise;
  return readFile(await download.path());
}

test.describe("/workspace record grid gestures", () => {
  test("copies a range and pastes it back with Windows line endings", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    await selectRange(
      page,
      { recordId: "TSK-0001", field: "code" },
      { recordId: "TSK-0002", field: "owner" },
    );

    // Tabs between cells, CRLF between rows, no header row, and no trailing
    // newline: the text a spreadsheet reads back as the block it was.
    expect(await copySelection(page)).toBe("T-007\tNorth\r\nB-1\tSouth");

    // A block from Excel on Windows, ending in a line break as one usually does.
    // The built-in parser leaves that "\r" on the last field of every row, which
    // is why the grid has its own.
    await selectRange(page, { recordId: "TSK-0003", field: "code" });
    await pasteText(page, "P-001\tWest\r\nP-002\tEast\r\n");

    await expect(cellOf(page, "TSK-0003", "code")).toHaveText("P-001");
    await expect(cellOf(page, "TSK-0003", "owner")).toHaveText("West");
    await expect(cellOf(page, "TSK-0004", "code")).toHaveText("P-002");
    await expect(cellOf(page, "TSK-0004", "owner")).toHaveText("East");
    // The row below the block is untouched.
    await expect(cellOf(page, "TSK-0005", "code")).toHaveText("X-3");

    const saved = await downloadedWorkspace(page);
    await openWorkspace(page, saved);
    await expect(cellOf(page, "TSK-0004", "owner")).toHaveText("East");
  });

  test("refuses a paste that would run past the last record", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    // Anchored on the last row, so two of the three rows have nowhere to go.
    // Records are not added by a paste, so the whole block is refused.
    await selectRange(page, { recordId: "TSK-0005", field: "code" });
    await pasteText(page, "A-1\r\nA-2\r\nA-3");

    await expect(page.getByTestId("workspace-grid-error")).toContainText(
      "2 more rows",
    );
    await expect(cellOf(page, "TSK-0005", "code")).toHaveText("X-3");
    // Nothing was written at all, so the workspace holds no change to save.
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);
  });

  test("fills a number series from the corner of a selection", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    await selectRange(
      page,
      { recordId: "TSK-0001", field: "effort" },
      { recordId: "TSK-0002", field: "effort" },
    );
    await dragFillHandleTo(page, "TSK-0004", "effort");

    await expect(cellOf(page, "TSK-0003", "effort")).toHaveText("30");
    await expect(cellOf(page, "TSK-0004", "effort")).toHaveText("40");
    // The drag stopped there, so the row below keeps what it had.
    await expect(cellOf(page, "TSK-0005", "effort")).toHaveText("3");

    const saved = await downloadedWorkspace(page);
    await openWorkspace(page, saved);
    await expect(cellOf(page, "TSK-0004", "effort")).toHaveText("40");
  });

  test("fills a date series by the step the source states", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    await selectRange(
      page,
      { recordId: "TSK-0001", field: "due" },
      { recordId: "TSK-0002", field: "due" },
    );
    await dragFillHandleTo(page, "TSK-0004", "due");

    await expect(cellOf(page, "TSK-0003", "due")).toHaveText("2026-01-15");
    await expect(cellOf(page, "TSK-0004", "due")).toHaveText("2026-01-22");
  });

  test("fills text by stepping the number on the end of it", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    await selectRange(page, { recordId: "TSK-0001", field: "code" });
    await dragFillHandleTo(page, "TSK-0003", "code");

    await expect(cellOf(page, "TSK-0002", "code")).toHaveText("T-008");
    await expect(cellOf(page, "TSK-0003", "code")).toHaveText("T-009");
  });

  test("keeps the cells a fill could write and explains the one it could not", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    // Sideways from a text cell across the columns beside it. The source ends in
    // a number, so the series steps it: the text column takes "X-3", and the
    // whole-number column cannot take "X-4".
    await selectRange(page, { recordId: "TSK-0004", field: "code" });
    await dragFillHandleTo(page, "TSK-0004", "effort");

    await expect(cellOf(page, "TSK-0004", "owner")).toHaveText("X-3");
    await expect(page.getByTestId("workspace-grid-error")).toContainText(
      "whole number",
    );
    // The refused cell still holds what the workspace holds, and says so.
    await expect(cellOf(page, "TSK-0004", "effort")).toHaveText("2");

    const saved = await downloadedWorkspace(page);
    await openWorkspace(page, saved);
    await expect(cellOf(page, "TSK-0004", "owner")).toHaveText("X-3");
    await expect(cellOf(page, "TSK-0004", "effort")).toHaveText("2");
  });

  test("refuses a whole gesture that names a workspace the worker does not hold", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await sendStaleGeneration(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    await selectRange(page, { recordId: "TSK-0003", field: "code" });
    await pasteText(page, "P-001\tWest\r\nP-002\tEast");

    await expect(page.getByTestId("workspace-grid-error")).toContainText(
      "no longer open",
    );
    // Not one cell of it landed: a stale gesture is refused as a whole, before
    // anything is written.
    await expect(cellOf(page, "TSK-0003", "code")).toHaveText("X-1");
    await expect(cellOf(page, "TSK-0003", "owner")).toHaveText("East");
    await expect(cellOf(page, "TSK-0004", "code")).toHaveText("X-2");
    await expect(page.getByTestId("workspace-unsaved")).toHaveCount(0);
  });

  test("selects on a single click and edits on a double one", async ({
    page,
  }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");
    await openWorkspace(page, await workspaceFixture());

    const cell = cellOf(page, "TSK-0001", "owner");
    // One click selects, because a drag from here is how a range is made.
    await cell.click();
    await expect(cell.locator("input")).toHaveCount(0);

    await cell.dblclick();
    const input = cell.locator("input");
    await expect(input).toBeVisible();
    await input.fill("Upper North");
    await input.press("Enter");
    await expect(cell).toHaveText("Upper North");
  });
});
