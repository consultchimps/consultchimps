import { expect, test } from "@playwright/test";

async function createDatabase(
  page: import("@playwright/test").Page,
  name: string,
): Promise<void> {
  await page.getByTestId("workspace-new-name").fill(name);
  await page.getByTestId("workspace-new").click();
  await expect(page.getByTestId("workspace-summary")).toContainText(name);
}

/**
 * The working-copy list is read from browser storage, not from a remembered
 * list, so it survives cleared page storage, orders by recent use when that
 * hint exists, and lets a closed copy be deleted.
 */
test("lists working copies from storage, closes, and deletes them", async ({
  page,
}) => {
  test.setTimeout(120_000);
  await page.goto("/tools/db");
  const copies = page.getByTestId("workspace-copy");
  await expect(page.getByTestId("workspace-recent")).toContainText(
    "No working copies are stored",
  );

  await createDatabase(page, "copy-a.sqlite");
  await createDatabase(page, "copy-b.sqlite");
  await expect(copies).toHaveCount(2);
  // Most recently used first while the ordering hint exists.
  await expect(copies.nth(0)).toContainText("copy-b.sqlite");
  await expect(copies.nth(1)).toContainText("copy-a.sqlite");
  await expect(page.getByTestId("workspace-storage-usage")).toContainText(
    "Browser storage used",
  );
  // The open copy offers Close and cannot be deleted from the list.
  await expect(copies.nth(0).getByTestId("workspace-close")).toBeVisible();
  await expect(copies.nth(0).getByTestId("workspace-copy-delete")).toHaveCount(
    0,
  );
  await expect(
    copies.nth(1).getByTestId("workspace-copy-delete"),
  ).toBeVisible();

  // Without the remembered hint the list still comes from storage.
  await page.evaluate(() => window.localStorage.clear());
  await page.reload();
  await expect(page.getByTestId("workspace-summary")).toHaveCount(0);
  await expect(copies).toHaveCount(2);
  await expect(copies.nth(0)).toContainText("copy-a.sqlite");
  await expect(copies.nth(1)).toContainText("copy-b.sqlite");

  // Delete needs confirmation, and Keep backs out.
  await copies.nth(0).getByTestId("workspace-copy-delete").click();
  await copies.nth(0).getByTestId("workspace-copy-delete-cancel").click();
  await expect(copies).toHaveCount(2);
  await copies.nth(0).getByTestId("workspace-copy-delete").click();
  await copies.nth(0).getByTestId("workspace-copy-delete-confirm").click();
  await expect(page.getByTestId("workspace-notice")).toContainText(
    'Deleted "copy-a.sqlite"',
  );
  await expect(copies).toHaveCount(1);
  await expect(copies.nth(0)).toContainText("copy-b.sqlite");

  // Reopen, close, then delete the last copy.
  await page.getByTestId("workspace-reopen").click();
  await expect(page.getByTestId("workspace-summary")).toContainText(
    "copy-b.sqlite",
  );
  await page.getByTestId("workspace-close").click();
  await expect(page.getByTestId("workspace-summary")).toHaveCount(0);
  await expect(page.getByTestId("workspace-notice")).toContainText(
    'Closed "copy-b.sqlite"',
  );
  await copies.nth(0).getByTestId("workspace-copy-delete").click();
  await copies.nth(0).getByTestId("workspace-copy-delete-confirm").click();
  await expect(copies).toHaveCount(0);
  // The empty state is the runtime's own listing of the storage pool, so
  // it also shows that removal reached storage.
  await expect(page.getByTestId("workspace-recent")).toContainText(
    "No working copies are stored",
  );
  await page.reload();
  await expect(page.getByTestId("workspace-recent")).toContainText(
    "No working copies are stored",
  );
});
