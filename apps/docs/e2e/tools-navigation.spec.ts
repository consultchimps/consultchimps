import { expect, test } from "@playwright/test";

/** Every in-browser tool: its sub-bar tab, its route, and its page heading. */
const BROWSER_TOOLS = [
  {
    tab: "Consolidate",
    route: "/tools/excel-consolidate",
    heading: "Consolidate Excel workbooks",
    card: "Consolidate spreadsheets",
  },
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
  {
    tab: "PowerPoint",
    route: "/tools/pptx-populate",
    heading: "Populate a PowerPoint template",
    card: "Populate PowerPoint templates",
  },
  {
    tab: "Inspect template",
    route: "/tools/pptx-inspect",
    heading: "Inspect a PowerPoint template",
    card: "Inspect PowerPoint templates",
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
    {
      url: "/docs/tools/spreadsheet-consolidate",
      tool: "/tools/excel-consolidate",
      label: "Consolidate",
    },
    {
      url: "/docs/tools/powerpoint-populate",
      tool: "/tools/pptx-populate",
      label: "PowerPoint",
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

  // There is deliberately no "this guide offers no online tool" table here:
  // every operation the registry declares now has a working browser surface,
  // so such a table would have no rows. The rule it used to assert — a guide
  // may only offer a button for an operation whose browser surface works — is
  // enforced for every entry by scripts/check-registry-site.ts, which reads
  // the registry the pages render from.

  // Two registry entries — populate and template inspection — point their
  // docHref at the same PowerPoint guide, the inspection one through an
  // anchor. `findToolByDocUrl` strips the fragment and returns the first
  // browser tool whose page matches, which is populate because it precedes
  // inspection in TOOLS. A guide page therefore offers exactly one button, so
  // readers are never asked to choose between two "Try … online" links; the
  // inspection tool is reached from the /tools sub-bar and its index card
  // instead.
  test("/docs/tools/powerpoint-populate offers only the populate tool", async ({
    page,
  }) => {
    await page.goto("/docs/tools/powerpoint-populate");

    const tryOnline = page.getByRole("link", { name: /^Try .+ online$/u });
    await expect(tryOnline).toHaveCount(1);
    await expect(tryOnline).toHaveAttribute("href", "/tools/pptx-populate");
    await expect(
      page.getByRole("link", { name: "Try Inspect template online" }),
    ).toHaveCount(0);
  });
});
