import { expect, test, type Page } from "@playwright/test";

/**
 * The shortcut page's two inputs, driven the way a visitor drives them: real
 * key presses into the capture area, and the key buttons beside it. The
 * matching itself is covered by the unit tests; what is only testable here is
 * that the presses reach the page at all, that the buttons and the keyboard
 * agree, and that the sequence can be undone.
 */

const rows = (page: Page) => page.getByTestId("shortcut-row");
const palette = (page: Page) => page.getByTestId("key-palette");

async function resultCount(page: Page): Promise<number> {
  const label = (await page.getByTestId("result-count").innerText()).trim();
  return Number(label.split(" ").at(0));
}

test.describe("/shortcuts", () => {
  test("is reachable from the site header", async ({ page }) => {
    await page.goto("/");
    await page.getByRole("link", { name: "Shortcuts", exact: true }).click();
    await expect(page).toHaveURL(/\/shortcuts$/u);
    await expect(
      page.getByRole("heading", { level: 1, name: "Press what you remember" }),
    ).toBeVisible();
    // The whole database is on the page before anything is entered.
    expect(await resultCount(page)).toBeGreaterThanOrEqual(200);
  });

  test("filters by word", async ({ page }) => {
    await page.goto("/shortcuts");
    const before = await resultCount(page);

    await page.getByTestId("word-search").fill("freeze");
    await expect(rows(page)).toHaveCount(3);
    expect(await resultCount(page)).toBeLessThan(before);
    await expect(page.getByText("Freeze the top row")).toBeVisible();

    // A word nothing carries leaves the page in its empty state.
    await page.getByTestId("word-search").fill("kerning");
    await expect(page.getByTestId("empty-state")).toBeVisible();
    await expect(rows(page)).toHaveCount(0);

    await page.getByTestId("word-search").fill("");
    expect(await resultCount(page)).toBe(before);
  });

  test("narrows step by step as the key buttons are clicked", async ({
    page,
  }) => {
    await page.goto("/shortcuts");
    const everything = await resultCount(page);

    await palette(page)
      .getByRole("button", { name: "Ctrl", exact: true })
      .click();
    const afterCtrl = await resultCount(page);
    expect(afterCtrl).toBeLessThan(everything);

    await palette(page)
      .getByRole("button", { name: "Shift", exact: true })
      .click();
    const afterShift = await resultCount(page);
    expect(afterShift).toBeLessThan(afterCtrl);

    await palette(page).getByRole("button", { name: "L", exact: true }).click();
    await expect(rows(page)).toHaveCount(1);
    await expect(page.getByText("Add or remove the filter row")).toBeVisible();
  });

  test("walks a ribbon sequence one step at a time", async ({ page }) => {
    await page.goto("/shortcuts");

    await palette(page)
      .getByRole("button", { name: "Alt", exact: true })
      .click();
    const afterAlt = await resultCount(page);

    // "Then" closes the current step, which is what separates Alt, H from
    // the Alt+... chords.
    await page.getByTestId("next-step").click();
    const afterStep = await resultCount(page);
    expect(afterStep).toBeLessThan(afterAlt);

    await palette(page).getByRole("button", { name: "H", exact: true }).click();
    await page.getByTestId("next-step").click();
    await palette(page).getByRole("button", { name: "O", exact: true }).click();
    await page.getByTestId("next-step").click();
    await palette(page).getByRole("button", { name: "I", exact: true }).click();

    await expect(rows(page)).toHaveCount(1);
    await expect(
      page.getByText("Fit the column width to the widest entry"),
    ).toBeVisible();
  });

  test("narrows when the keys are actually pressed", async ({ page }) => {
    await page.goto("/shortcuts");
    const capture = page.getByTestId("key-capture");
    await capture.click();
    await expect(page.getByTestId("capture-state")).toHaveText("Listening");

    // Held together, the way the chord is pressed in Excel.
    await page.keyboard.down("Control");
    const afterCtrl = await resultCount(page);
    await page.keyboard.down("Shift");
    const afterShift = await resultCount(page);
    expect(afterShift).toBeLessThan(afterCtrl);

    await page.keyboard.press("KeyL");
    await page.keyboard.up("Shift");
    await page.keyboard.up("Control");

    await expect(rows(page)).toHaveCount(1);
    await expect(page.getByText("Add or remove the filter row")).toBeVisible();
    await expect(page.getByTestId("entered-sequence")).toContainText("Ctrl");
  });

  test("takes the modifiers in whatever order they were pressed", async ({
    page,
  }) => {
    await page.goto("/shortcuts");
    await page.getByTestId("key-capture").click();

    // Shift first, against a chord the database stores as Alt+Shift+Right.
    await page.keyboard.down("Shift");
    await page.keyboard.down("Alt");
    await page.keyboard.press("ArrowRight");
    await page.keyboard.up("Alt");
    await page.keyboard.up("Shift");

    await expect(rows(page)).toHaveCount(1);
    await expect(
      page.getByText("Group the selected PivotTable items"),
    ).toBeVisible();
  });

  test("presses each step of a ribbon route in turn", async ({ page }) => {
    await page.goto("/shortcuts");
    await page.getByTestId("key-capture").click();

    await page.keyboard.press("Alt");
    await page.keyboard.press("KeyH");
    await page.keyboard.press("KeyO");
    const afterThreeSteps = await resultCount(page);
    expect(afterThreeSteps).toBeGreaterThan(0);

    await page.keyboard.press("KeyI");
    await expect(rows(page)).toHaveCount(1);
    await expect(
      page.getByText("Fit the column width to the widest entry"),
    ).toBeVisible();
  });

  test("undoes the last key with Backspace and clears with Escape", async ({
    page,
  }) => {
    await page.goto("/shortcuts");
    const everything = await resultCount(page);
    const capture = page.getByTestId("key-capture");
    await capture.click();

    await page.keyboard.down("Control");
    await page.keyboard.down("Shift");
    await page.keyboard.press("KeyL");
    await page.keyboard.up("Shift");
    await page.keyboard.up("Control");
    await expect(rows(page)).toHaveCount(1);

    // Backspace pops the L and widens the list back out to the Ctrl+Shift
    // chords, rather than clearing everything.
    await page.keyboard.press("Backspace");
    const afterPop = await resultCount(page);
    expect(afterPop).toBeGreaterThan(1);
    expect(afterPop).toBeLessThan(everything);
    await expect(page.getByTestId("entered-sequence")).toContainText("Shift");

    await page.keyboard.press("Escape");
    expect(await resultCount(page)).toBe(everything);
    await expect(page.getByTestId("entered-sequence")).toHaveCount(0);
  });

  test("replaces the last key of a chord after Backspace", async ({ page }) => {
    await page.goto("/shortcuts");
    await page.getByTestId("key-capture").click();

    await page.keyboard.down("Control");
    await page.keyboard.down("Shift");
    await page.keyboard.press("KeyL");
    await page.keyboard.up("Shift");
    await page.keyboard.up("Control");
    await expect(rows(page)).toHaveCount(1);

    // Backspace edits the chord rather than finishing it, so the next key
    // joins Ctrl+Shift instead of starting a second step.
    await page.keyboard.press("Backspace");
    await page.keyboard.press("KeyO");

    await expect(rows(page)).toHaveCount(1);
    await expect(
      page.getByText("Select the cells that carry a note or a comment"),
    ).toBeVisible();
  });

  test("lets a keyboard visitor tab out of the capture area", async ({
    page,
  }) => {
    await page.goto("/shortcuts");
    const capture = page.getByTestId("key-capture");
    await capture.click();
    await expect(capture).toBeFocused();

    // Tab keeps its default here: capturing it would leave anyone navigating
    // by keyboard stuck inside the area.
    await page.keyboard.press("Tab");
    await expect(capture).not.toBeFocused();
    // Nothing was recorded from the press either.
    await expect(page.getByTestId("entered-sequence")).toHaveCount(0);

    // Shift+Tab navigates backwards, and the Shift it is pressed with is part
    // of leaving rather than part of a search, so it leaves no filter behind.
    const everything = await resultCount(page);
    await capture.click();
    await page.keyboard.press("Shift+Tab");
    await expect(capture).not.toBeFocused();
    await expect(page.getByTestId("entered-sequence")).toHaveCount(0);
    expect(await resultCount(page)).toBe(everything);
  });

  test("keeps a finished chord when Shift+Tab moves focus away", async ({
    page,
  }) => {
    await page.goto("/shortcuts");
    const capture = page.getByTestId("key-capture");
    await capture.click();

    await page.keyboard.down("Control");
    await page.keyboard.down("Shift");
    await page.keyboard.press("KeyL");
    await page.keyboard.up("Shift");
    await page.keyboard.up("Control");
    await expect(rows(page)).toHaveCount(1);

    // Navigating away leaves the search the visitor was reading alone.
    await page.keyboard.press("Shift+Tab");
    await expect(capture).not.toBeFocused();
    await expect(rows(page)).toHaveCount(1);
    await expect(page.getByText("Add or remove the filter row")).toBeVisible();
  });

  test("keeps a key the step already held when Shift+Tab moves focus away", async ({
    page,
  }) => {
    await page.goto("/shortcuts");

    // Ctrl and Shift entered with the buttons, so the step is open and
    // already carries the Shift that the navigation gesture also presses.
    await palette(page)
      .getByRole("button", { name: "Ctrl", exact: true })
      .click();
    await palette(page)
      .getByRole("button", { name: "Shift", exact: true })
      .click();
    const chosen = await resultCount(page);

    const capture = page.getByTestId("key-capture");
    await capture.click();
    await page.keyboard.press("Shift+Tab");
    await expect(capture).not.toBeFocused();

    await expect(page.getByTestId("entered-sequence")).toContainText("Shift");
    expect(await resultCount(page)).toBe(chosen);
  });

  test("clears the sequence from the button, and combines both filters", async ({
    page,
  }) => {
    await page.goto("/shortcuts");

    await page.getByTestId("word-search").fill("row");
    await palette(page)
      .getByRole("button", { name: "Shift", exact: true })
      .click();
    const combined = await resultCount(page);
    expect(combined).toBeGreaterThan(0);

    // Every row on the page satisfies both halves of the search.
    const shown = await rows(page).count();
    expect(shown).toBe(combined);

    await page.getByTestId("clear-sequence").click();
    expect(await resultCount(page)).toBeGreaterThan(combined);
    await expect(page.getByTestId("word-search")).toHaveValue("row");
  });
});
