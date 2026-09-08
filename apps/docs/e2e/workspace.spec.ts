import { expect, test, type Page } from "@playwright/test";
import { readFile } from "node:fs/promises";

/**
 * The data workspace shell, exercised through the fallback (download and file
 * input) path so it runs in headless Chromium without the File System Access
 * API's native pickers. The whole shell is here: start an empty workspace, save
 * it to a `.sqlite` file, and reopen those very bytes.
 *
 * Removing the pickers before the page loads is deliberate. Their dialogs cannot
 * be driven from a test, and it is the fallback that every browser without the
 * API relies on, so it is the path most worth covering.
 */
async function forceDownloadFallback(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const globals = window as unknown as Record<string, unknown>;
    delete globals["showOpenFilePicker"];
    delete globals["showSaveFilePicker"];
  });
}

/** The header bytes every SQLite file starts with, so a valid save is checkable. */
const SQLITE_HEADER = "SQLite format 3";

async function downloadedWorkspace(
  page: Page,
  trigger: () => Promise<void>,
): Promise<Buffer> {
  const downloadPromise = page.waitForEvent("download");
  await trigger();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("workspace.sqlite");
  const bytes = await readFile(await download.path());
  expect(bytes.byteLength).toBeGreaterThan(0);
  return bytes;
}

test.describe("/workspace", () => {
  test("is reachable from the site header", async ({ page }) => {
    await page.goto("/");
    await page.getByRole("link", { name: "Workspace", exact: true }).click();
    await expect(page).toHaveURL(/\/workspace$/u);
    await expect(
      page.getByRole("heading", { level: 1, name: "Data workspace" }),
    ).toBeVisible();
    await expect(page.getByTestId("workspace-empty")).toBeVisible();
  });

  test("creates, saves, and reopens an empty workspace", async ({ page }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");

    // Start a new, empty workspace.
    await page.getByTestId("workspace-new").click();
    await expect(page.getByTestId("workspace-summary")).toBeVisible();
    await expect(page.getByTestId("workspace-file-name")).toHaveText(
      "New workspace",
    );
    await expect(page.getByTestId("workspace-table-count")).toHaveText("0");

    // Save it. Without the File System Access API this downloads a copy, which
    // also covers the worker's serialize path and the byte saver.
    const bytes = await downloadedWorkspace(page, () =>
      page.getByTestId("workspace-save-as").click(),
    );
    expect(bytes.subarray(0, SQLITE_HEADER.length).toString("latin1")).toBe(
      SQLITE_HEADER,
    );

    // Reopen exactly those saved bytes through the fallback file input.
    await page.getByTestId("file-input").setInputFiles({
      name: "reopened.sqlite",
      mimeType: "application/vnd.sqlite3",
      buffer: bytes,
    });

    await expect(page.getByTestId("workspace-notice")).toHaveText(
      "Opened the workspace",
    );
    await expect(page.getByTestId("workspace-file-name")).toHaveText(
      "reopened.sqlite",
    );
    await expect(page.getByTestId("workspace-table-count")).toHaveText("0");
  });

  test("reports a file that is not a readable database", async ({ page }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");

    // A file with the right extension but bytes that are not a database: the
    // worker's open rejects with a stable error, which the page shows.
    await page.getByTestId("file-input").setInputFiles({
      name: "broken.sqlite",
      mimeType: "application/vnd.sqlite3",
      buffer: Buffer.from("this is not a database\n", "utf8"),
    });

    await expect(page.getByTestId("workspace-error")).toContainText(
      "not a readable database file",
    );
    await expect(page.getByTestId("workspace-summary")).toHaveCount(0);
  });

  test("refuses a file that is not a workspace type", async ({ page }) => {
    await forceDownloadFallback(page);
    await page.goto("/workspace");

    await page.getByTestId("file-input").setInputFiles({
      name: "notes.txt",
      mimeType: "text/plain",
      buffer: Buffer.from("just text\n", "utf8"),
    });

    await expect(page.getByTestId("workspace-error")).toContainText(
      "not a .sqlite workspace file",
    );
    await expect(page.getByTestId("workspace-summary")).toHaveCount(0);
  });
});
