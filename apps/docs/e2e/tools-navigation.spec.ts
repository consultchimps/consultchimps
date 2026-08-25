import { expect, test } from "@playwright/test";

/** Every in-browser tool: its sub-bar tab, its route, and its page heading. */
const BROWSER_TOOLS = [
  {
    tab: "Merge tabs",
    route: "/tools/excel-merge",
    heading: "Merge Excel workbooks",
    card: "Merge workbook tabs",
  },
  {
    tab: "Split Excel",
    route: "/tools/excel-split",
    heading: "Split an Excel workbook",
    card: "Split spreadsheets",
  },
  {
    tab: "Split PDF",
    route: "/tools/pdf-split",
    heading: "Split a PDF",
    card: "Split PDF pages",
  },
  {
    tab: "Merge PDFs",
    route: "/tools/pdf-merge",
    heading: "Merge PDFs",
    card: "Merge PDF packs",
  },
] as const;

test.describe("/tools", () => {
  test("lists the in-browser tools", async ({ page }) => {
    await page.goto("/tools");
    await expect(
      page.getByRole("heading", { level: 1, name: "Run a tool right here." }),
    ).toBeVisible();
    for (const tool of BROWSER_TOOLS) {
      await expect(page.getByRole("link", { name: tool.card })).toBeVisible();
    }
  });

  test("moves between the tool pages from the sub-bar", async ({ page }) => {
    await page.goto("/tools");
    // The landing cards link to the same routes, so the tabs are addressed
    // through the sub-bar's accessible name.
    const tabs = page.getByRole("navigation", { name: "Online tools" });

    for (const tool of BROWSER_TOOLS) {
      await tabs.getByRole("link", { name: tool.tab, exact: true }).click();
      await expect(page).toHaveURL(new RegExp(`${tool.route}$`, "u"));
      await expect(
        page.getByRole("heading", { level: 1, name: tool.heading }),
      ).toBeVisible();
    }

    await tabs.getByRole("link", { name: "All tools" }).click();
    await expect(page).toHaveURL(/\/tools$/u);
    await expect(
      page.getByRole("heading", { level: 1, name: "Run a tool right here." }),
    ).toBeVisible();
  });
});

test.describe("tool guides", () => {
  const GUIDES = [
    {
      url: "/docs/tools/pdf-split",
      tool: "/tools/pdf-split",
      label: "Split PDF",
    },
    {
      url: "/docs/tools/pdf-merge",
      tool: "/tools/pdf-merge",
      label: "Merge PDFs",
    },
    {
      url: "/docs/tools/spreadsheet-split",
      tool: "/tools/excel-split",
      label: "Split Excel",
    },
    {
      url: "/docs/tools/workbook-merge",
      tool: "/tools/excel-merge",
      label: "Merge tabs",
    },
  ] as const;

  for (const guide of GUIDES) {
    test(`${guide.url} offers the online tool`, async ({ page }) => {
      await page.goto(guide.url);
      const tryOnline = page.getByRole("link", {
        name: `Try ${guide.label} online`,
      });
      await expect(tryOnline).toBeVisible();
      await expect(tryOnline).toHaveAttribute("href", guide.tool);
    });
  }

  // Guides for operations without a browser surface must not promise an
  // online tool at all — neither consolidation nor the populate guide's two
  // operations (populate and template inspection) run in the browser.
  const GUIDES_WITHOUT_TOOL = [
    {
      url: "/docs/tools/spreadsheet-consolidate",
      heading: "Consolidate spreadsheets into one table",
    },
    {
      url: "/docs/tools/powerpoint-populate",
      heading: "Populate a PowerPoint template",
    },
  ] as const;

  for (const guide of GUIDES_WITHOUT_TOOL) {
    test(`${guide.url} offers no online tool`, async ({ page }) => {
      await page.goto(guide.url);
      await expect(
        page.getByRole("heading", { level: 1, name: guide.heading }),
      ).toBeVisible();
      await expect(
        page.getByRole("link", { name: /^Try .+ online$/u }),
      ).toHaveCount(0);
    });
  }
});
