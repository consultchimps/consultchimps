import { expect, test } from "@playwright/test";

test.describe("/tools", () => {
  test("lists the in-browser tools", async ({ page }) => {
    await page.goto("/tools");
    await expect(
      page.getByRole("heading", { level: 1, name: "Run a tool right here." }),
    ).toBeVisible();
    await expect(
      page.getByRole("link", { name: "Split PDF pages" }),
    ).toBeVisible();
    await expect(
      page.getByRole("link", { name: "Merge PDF packs" }),
    ).toBeVisible();
  });

  test("moves between the tool pages from the sub-bar", async ({ page }) => {
    await page.goto("/tools");
    // The landing cards link to the same routes, so the tabs are addressed
    // through the sub-bar's accessible name.
    const tabs = page.getByRole("navigation", { name: "Online tools" });

    await tabs.getByRole("link", { name: "Split PDF" }).click();
    await expect(page).toHaveURL(/\/tools\/pdf-split$/u);
    await expect(
      page.getByRole("heading", { level: 1, name: "Split a PDF" }),
    ).toBeVisible();

    await tabs.getByRole("link", { name: "Merge PDFs" }).click();
    await expect(page).toHaveURL(/\/tools\/pdf-merge$/u);
    await expect(
      page.getByRole("heading", { level: 1, name: "Merge PDFs" }),
    ).toBeVisible();

    await tabs.getByRole("link", { name: "All tools" }).click();
    await expect(page).toHaveURL(/\/tools$/u);
    await expect(
      page.getByRole("heading", { level: 1, name: "Run a tool right here." }),
    ).toBeVisible();
  });
});
