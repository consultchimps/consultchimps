import { expect, test } from "@playwright/test";

/**
 * The database tool runs in one tab at a time per browser profile. The page
 * claims a Web Lock before it starts an engine; a second page of the same
 * origin sees the conflict notice instead of a raw file-handle error and takes
 * over on its own once the first page is gone. A closed page can keep its
 * SQLite file handles for several seconds, so the worker waits for them.
 */
test.describe("database tool tab ownership", () => {
  test("second tab reports the conflict and takes over after the first closes", async ({
    context,
  }) => {
    // Two engine downloads, a lock hand-over, and a busy wait for the closed
    // tab's handles add up to more than the default minute on a slow runner.
    test.setTimeout(180_000);
    const first = await context.newPage();
    await first.goto("/tools/db");
    await expect(first.getByTestId("workspace-tab-conflict")).toHaveCount(0);
    await expect(first.getByTestId("workspace-new")).toBeEnabled();

    const second = await context.newPage();
    await second.goto("/tools/db");
    const conflict = second.getByTestId("workspace-tab-conflict");
    await expect(conflict).toBeVisible();
    await expect(conflict).toContainText("open in another tab");
    await expect(second.getByTestId("workspace-new")).toBeDisabled();
    await expect(second.getByTestId("workspace-open")).toBeDisabled();
    await expect(second.getByTestId("workspace-error")).toHaveCount(0);

    // The first tab keeps working while the second waits.
    await first.getByTestId("workspace-new-name").fill("ownership-a.sqlite");
    await first.getByTestId("workspace-new").click();
    await expect(first.getByTestId("workspace-summary")).toBeVisible();
    await expect(conflict).toBeVisible();

    await first.close();
    // A closed tab releases its Web Lock after about 1.5 s on an idle
    // machine and later under load; the takeover is automatic either way.
    await expect(conflict).toHaveCount(0, { timeout: 60_000 });
    await expect(second.getByTestId("workspace-new")).toBeEnabled();
    await second.getByTestId("workspace-new-name").fill("ownership-b.sqlite");
    await second.getByTestId("workspace-new").click();
    await expect(second.getByTestId("workspace-summary")).toBeVisible({
      timeout: 60_000,
    });
    await expect(second.getByTestId("workspace-error")).toHaveCount(0);
  });

  test("a busy engine start waits for the closed tab's handles", async ({
    context,
  }) => {
    test.setTimeout(180_000);
    // Bypass the page guard so the worker itself meets the busy pool, the
    // way a browser without Web Locks would. While the first tab is open the
    // start fails with a stable code; after it closes the next start waits
    // for the handles and succeeds.
    const first = await context.newPage();
    await first.goto("/tools/db");
    await first.getByTestId("workspace-new-name").fill("engine-a.sqlite");
    await first.getByTestId("workspace-new").click();
    await expect(first.getByTestId("workspace-summary")).toBeVisible();

    const second = await context.newPage();
    await second.addInitScript(() => {
      Object.defineProperty(navigator, "locks", { value: undefined });
    });
    await second.goto("/tools/db");
    await expect(second.getByTestId("workspace-tab-conflict")).toHaveCount(0);
    await second.getByTestId("workspace-new-name").fill("engine-b.sqlite");
    await second.getByTestId("workspace-new").click();
    await expect(second.getByTestId("workspace-progress")).toContainText(
      "Waiting for another tab",
    );
    await second.getByTestId("workspace-cancel").click();
    await expect(second.getByTestId("workspace-notice")).toContainText(
      "Cancelled",
    );

    await first.close();
    await second.getByTestId("workspace-new").click();
    await expect(second.getByTestId("workspace-summary")).toBeVisible({
      timeout: 60_000,
    });
    await expect(second.getByTestId("workspace-format")).toHaveText("sqlite");
  });

  test("a failed engine download is retried on the next action", async ({
    context,
  }) => {
    // Worker requests are only intercepted at the context level, and the
    // engine loader falls back to a second fetch when streaming compilation
    // fails, so every request stays blocked until the failure is reported.
    let offline = true;
    await context.route("**/database-wasm/sqlite3.wasm", async (route) => {
      if (offline) {
        await route.abort("failed");
        return;
      }
      await route.continue();
    });
    const page = await context.newPage();
    await page.goto("/tools/db");
    await page.getByTestId("workspace-new-name").fill("download.sqlite");
    await page.getByTestId("workspace-new").click();
    const error = page.getByTestId("workspace-error");
    await expect(error).toBeVisible({ timeout: 30_000 });
    await expect(error).toContainText("DB_BROWSER_ENGINE_UNAVAILABLE");
    await expect(error).toContainText("could not start");

    offline = false;
    await page.getByTestId("workspace-new").click();
    await expect(page.getByTestId("workspace-summary")).toBeVisible({
      timeout: 30_000,
    });
  });
});
