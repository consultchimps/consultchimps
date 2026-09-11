import { Database, type TableSchema } from "@consultchimps/db";
import { contrastRatio } from "@consultchimps/theme";
import { expect, test, type Locator, type Page } from "@playwright/test";
import { writeFile } from "node:fs/promises";

/**
 * The record grid wearing the site's theme, in light and in dark.
 *
 * Tabulator ships hardcoded hex, so `src/app/workspace-grid.css` binds every
 * colour it draws to a site token. Two kinds of promise can only be checked
 * here rather than in a unit test, because both are about what a browser
 * actually paints once the cascade, the alpha, and the `color-mix` recipes have
 * been resolved:
 *
 * - the grid's surfaces are the site's own tokens, in both modes, with none of
 *   Tabulator's hexes surviving anywhere on screen;
 * - the text on those surfaces clears 4.5 to 1 and the state indicators clear
 *   3 to 1, measured with `contrastRatio` from `@consultchimps/theme`, which is
 *   the same ratio that gates a palette.
 *
 * The mode switch is asserted the way a visitor meets it (a fresh load under
 * `prefers-color-scheme: dark`) and the way the site's own toggle drives it (the
 * `.dark` class arriving on `<html>` while the page is up). The second is what
 * shows the grid is themed by CSS rather than re-rendered: the same cell
 * elements, marked before the switch, are still there afterwards wearing the
 * dark surface.
 */

const region: TableSchema = {
  name: "Region",
  columns: [{ name: "name", type: "text", nullable: false }],
  foreignKeys: [],
  recordId: { prefix: "REG", padding: 4 },
};

/** A table with no records, so the empty-table placeholder can be measured. */
const note: TableSchema = {
  name: "Note",
  columns: [{ name: "text", type: "text" }],
  foreignKeys: [],
  recordId: { prefix: "NOTE", padding: 4 },
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

/**
 * Enough rows to show a striped table, a foreign-key column with a picker, and
 * a boolean column with both a tick and a cross, plus one table with nothing in
 * it so the placeholder is on screen too.
 */
async function workspaceFixture(): Promise<Buffer> {
  const database = await Database.create();
  database.createTable(region);
  database.createTable(customer);
  database.createTable(note);
  database.insertRecord("Region", { name: "North" });
  database.insertRecord("Region", { name: "South" });
  for (const [index, name] of [
    "Client A",
    "Client B",
    "Client C",
    "Client D",
  ].entries()) {
    database.insertRecord("Customer", {
      name,
      region: index % 2 === 0 ? "REG-0001" : "REG-0002",
      headcount: (index + 1) * 5,
      active: index % 2 === 0,
    });
  }
  const bytes = Buffer.from(database.serialize());
  database.close();
  return bytes;
}

async function forceDownloadFallback(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const globals = window as unknown as Record<string, unknown>;
    delete globals["showOpenFilePicker"];
    delete globals["showSaveFilePicker"];
  });
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

function cellOf(page: Page, recordId: string, field: string): Locator {
  return page.locator(
    `[data-record-id="${recordId}"] [tabulator-field="${field}"]`,
  );
}

/**
 * Read a colour the way the screen reads it, as an sRGB hex value.
 *
 * A computed colour on this page is not always `rgb(...)`: a `color-mix` recipe
 * computes as `color(srgb ...)` and a Tailwind opacity modifier as `oklab(...)`.
 * Rather than teach this spec every colour space, it paints the values onto a
 * one-pixel canvas, in order, and reads the pixel back. That is the browser's
 * own conversion, and the same pass composites a translucent colour onto what
 * is behind it exactly as the compositor would.
 */
function paint(page: Page, ...colours: string[]): Promise<string> {
  return page.evaluate((values) => {
    const canvas = document.createElement("canvas");
    canvas.width = 1;
    canvas.height = 1;
    const context = canvas.getContext("2d", { willReadFrequently: true });
    if (context === null) {
      throw new Error("This browser gave the test no 2d canvas to measure in");
    }
    // A canvas ignores a value it cannot parse and keeps the one it had, which
    // would quietly turn a misread colour into whatever was painted before it.
    // Offering the same value against two different standing colours catches
    // that: a value the canvas takes reads back the same both times.
    const parses = (colour: string): boolean => {
      context.fillStyle = "#000000";
      context.fillStyle = colour;
      const first = context.fillStyle;
      context.fillStyle = "#ffffff";
      context.fillStyle = colour;
      return context.fillStyle === first;
    };
    // The page is painted on the browser's white base, so anything still
    // translucent once every layer is applied lands on white here too.
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, 1, 1);
    for (const colour of values) {
      if (!parses(colour)) {
        throw new Error(`This browser cannot read the colour "${colour}"`);
      }
      context.fillStyle = colour;
      context.fillRect(0, 0, 1, 1);
    }
    const pixel = context.getImageData(0, 0, 1, 1).data;
    return `#${[pixel[0]!, pixel[1]!, pixel[2]!]
      .map((part) => part.toString(16).padStart(2, "0"))
      .join("")}`;
  }, colours);
}

/**
 * The colour of a computed `box-shadow`, which the browser reports first, ahead
 * of the offsets: "rgb(252, 249, 243) 0px 0px 0px 2px inset". Only the first
 * shadow of a list is read, which is all the layer ever sets. A misread would
 * not slip through: `paint` refuses a value the canvas cannot parse.
 */
function shadowColour(shadow: string): string {
  const colour = /^(.*?\))\s+[-\d]/.exec(shadow)?.[1];
  if (colour === undefined) {
    throw new Error(`No colour at the front of the shadow "${shadow}"`);
  }
  return colour;
}

function styleOf(locator: Locator, property: string): Promise<string> {
  return locator.evaluate(
    (element, name) => getComputedStyle(element).getPropertyValue(name),
    property,
  );
}

/** A painted property (a background, an svg fill) as hex. */
async function paintedStyle(
  page: Page,
  locator: Locator,
  property: string,
): Promise<string> {
  return paint(page, await styleOf(locator, property));
}

/**
 * A CSS declaration resolved through the page's own custom properties, as hex.
 * Reading `--color-fd-card` as text would give the token's `hsl(...)` source;
 * painting it gives what a visitor sees, and it resolves a `color-mix` too.
 */
async function token(page: Page, declaration: string): Promise<string> {
  const computed = await page.evaluate((value) => {
    const probe = document.createElement("div");
    probe.style.backgroundColor = value;
    document.body.append(probe);
    const painted = getComputedStyle(probe).backgroundColor;
    probe.remove();
    return painted;
  }, declaration);
  return paint(page, computed);
}

/**
 * What is actually behind an element: every ancestor background from the
 * document down, composited, so a cell that sets none of its own reports the
 * row's, and a translucent panel reports what it lets through.
 */
async function surfaceOf(page: Page, locator: Locator): Promise<string> {
  const layers = await locator.evaluate((element) => {
    const found: string[] = [];
    let node: HTMLElement | null = element as HTMLElement;
    while (node !== null) {
      found.unshift(getComputedStyle(node).backgroundColor);
      node = node.parentElement;
    }
    return found;
  });
  return paint(page, ...layers);
}

async function contrastOf(
  page: Page,
  text: Locator,
  behind?: string,
): Promise<{ ratio: number; ink: string; surface: string }> {
  const surface = behind ?? (await surfaceOf(page, text));
  const ink = await paint(page, surface, await styleOf(text, "color"));
  return { ratio: contrastRatio(ink, surface), ink, surface };
}

/** Measurements collected per mode, reported so the numbers are on the record. */
interface Measurement {
  what: string;
  ratio: number;
  floor: number;
}

/**
 * Open a workspace with the grid showing, in the mode asked for. Dark is
 * reached the way a visitor on a dark desktop reaches it: the site's theme is
 * "system", so the preference alone puts `.dark` on the document.
 */
async function openGrid(page: Page, mode: "light" | "dark"): Promise<void> {
  await page.emulateMedia({ colorScheme: mode });
  await forceDownloadFallback(page);
  await page.goto("/workspace");
  await openWorkspace(page, await workspaceFixture());
  if (mode === "dark") {
    await expect(page.locator("html")).toHaveClass(/\bdark\b/);
  } else {
    await expect(page.locator("html")).not.toHaveClass(/\bdark\b/);
  }
  await expect(cellOf(page, "CUST-0001", "name")).toHaveText("Client A");
}

/**
 * The foreign-key picker. Tabulator gives the list element the popup container's
 * class as well as its own, and the tooltip is another element wearing that
 * container class, so the picker is addressed by both of its classes rather
 * than by the one they share.
 */
function pickerOf(page: Page): Locator {
  return page.locator(".tabulator-popup-container.tabulator-edit-list");
}

/**
 * Open the picker on the first customer, retrying the whole interaction rather
 * than any one step of it: Tabulator re-lays its columns out as the table sizes
 * itself, which moves the cell under a click. The same reason
 * `workspace-grid.spec.ts` retries its edits.
 */
async function openPicker(page: Page, popup: Locator): Promise<void> {
  await expect(async () => {
    await cellOf(page, "CUST-0001", "region").click({ timeout: 2_000 });
    await expect(popup).toBeVisible({ timeout: 2_000 });
  }).toPass({ timeout: 30_000 });
}

/** Every colour Tabulator's own stylesheet declares, as hex. */
const TABULATOR_HEXES = [
  "#ffffff",
  "#000000",
  "#888888",
  "#e6e6e6",
  "#efefef",
  "#333333",
  "#555555",
  "#666666",
  "#cccccc",
  "#bbbbbb",
  "#aaaaaa",
  "#999999",
  "#cdcdcd",
  "#f3f3f3",
  "#e2e2e2",
  "#d6d6d6",
  "#dcdcdc",
  "#2dc214",
  "#ce1515",
  "#1d68cd",
  "#1c6cc2",
  "#2975dd",
  "#3876ca",
  "#9abcea",
  "#769bcc",
  "#dd0000",
  "#590000",
  "#dd0000",
];

test.describe("/workspace record grid theme", () => {
  for (const mode of ["light", "dark"] as const) {
    test(`paints the grid in the site's tokens in ${mode} mode`, async ({
      page,
    }, testInfo) => {
      await openGrid(page, mode);

      const card = await token(page, "var(--color-fd-card)");
      const muted = await token(page, "var(--color-fd-muted)");
      const popover = await token(page, "var(--color-fd-popover)");
      const stripe = await token(
        page,
        "color-mix(in srgb, var(--color-fd-foreground) 4%, var(--color-fd-card))",
      );

      // The surfaces are the site's, not Tabulator's. A row is addressed by the
      // class that decides its shading rather than by its position, because
      // which row Tabulator counts as even is Tabulator's business.
      const plainRow = page.locator(".tabulator-row-odd").first();
      const shadedRow = page.locator(".tabulator-row-even").first();
      const header = page.locator(".tabulator-header").first();
      const container = page.locator(".tabulator").first();
      const tick = page.locator(".tabulator-tick").first();
      const cross = page.locator(".tabulator-cross").first();

      expect(await surfaceOf(page, plainRow)).toBe(card);
      expect(await surfaceOf(page, shadedRow)).toBe(stripe);
      expect(await paintedStyle(page, header, "background-color")).toBe(muted);
      expect(await paintedStyle(page, container, "background-color")).toBe(
        card,
      );

      // A boolean column reads as a value rather than as a verdict.
      expect(await paintedStyle(page, tick, "fill")).toBe(
        await token(page, "var(--color-fd-card-foreground)"),
      );
      expect(await paintedStyle(page, cross, "fill")).toBe(
        await token(page, "var(--color-fd-muted-foreground)"),
      );

      // The foreign-key picker, which Tabulator renders on the document body.
      const popup = pickerOf(page);
      await openPicker(page, popup);
      // Read while it is open: the picker is gone by the time the surfaces are
      // gathered below, and a locator for it would wait for a lifetime.
      const pickerSurface = await paintedStyle(page, popup, "background-color");
      expect(pickerSurface).toBe(popover);

      const measurements: Measurement[] = [];
      const measure = (
        what: string,
        ratio: number,
        floor: number,
        ink: string,
        surface: string,
      ): void => {
        measurements.push({
          what: `${what} (${ink} on ${surface})`,
          ratio,
          floor,
        });
        expect(
          ratio,
          `${what} in ${mode} mode: ${ink} on ${surface} is ${ratio.toFixed(2)} to 1, floor ${floor}`,
        ).toBeGreaterThanOrEqual(floor);
      };
      const record = async (
        what: string,
        text: Locator,
        floor: number,
      ): Promise<void> => {
        const { ratio, ink, surface } = await contrastOf(page, text);
        measure(what, ratio, floor, ink, surface);
      };

      // The option the cell already holds. Nothing but its own colour says it
      // is the chosen one, so the surface answers to the indicator mark as well
      // as the text on it answering to the text mark.
      const chosen = popup.locator(".tabulator-edit-list-item.active").first();
      await expect(chosen).toBeVisible();
      const chosenSurface = await surfaceOf(page, chosen);
      measure(
        "the chosen option against the picker surface",
        contrastRatio(chosenSurface, pickerSurface),
        3,
        chosenSurface,
        pickerSurface,
      );
      await record("picker text on the chosen option", chosen, 4.5);
      await record(
        "picker option text",
        popup.locator(".tabulator-edit-list-item:not(.active)").first(),
        4.5,
      );

      // The focus ring on each kind of option, measured against the surface it
      // is drawn on, which is the option's own fill: the ring that a visitor
      // sees is an inset one, because the popup clips the sides of an outline
      // (see the layer's own note on it). The two options are filled
      // differently, so each is measured against its own.
      //
      // The list opens with the option the cell holds already noted as focused
      // while no element carries the class yet, so the first press steps past
      // it onto the next option and the second press comes back to it. That is
      // also why the fixture gives the first customer the first region: were it
      // the last option in the list, the first press would have nowhere to go,
      // and the count below would fail rather than quietly measure nothing.
      const focused = popup.locator(".tabulator-edit-list-item.focused");
      const focusedAndChosen = popup.locator(
        ".tabulator-edit-list-item.active.focused",
      );
      for (const [what, key, isChosen] of [
        ["an option the cell does not hold", "ArrowDown", false],
        ["the option the cell holds", "ArrowUp", true],
      ] as const) {
        await page.keyboard.press(key);
        await expect(focused).toHaveCount(1);
        await expect(focusedAndChosen).toHaveCount(isChosen ? 1 : 0);

        const fill = await surfaceOf(page, focused);
        const shadow = await styleOf(focused, "box-shadow");
        expect(shadow, `${what} should carry a focus ring`).toMatch(/inset/);
        const ring = await paint(page, shadowColour(shadow));
        measure(
          `the focus ring on ${what}`,
          contrastRatio(ring, fill),
          3,
          ring,
          fill,
        );

        // Tabulator's own outline is bound too, and the horizontal edge of it
        // between two options is the part the popup does not clip. On the
        // option the cell holds it would be its own fill colour and say
        // nothing, so it is answered here, where it is a ring on the popup.
        if (!isChosen) {
          const outline = await paintedStyle(page, focused, "outline-color");
          measure(
            `the outline on ${what}, where the popup does not clip it`,
            contrastRatio(outline, pickerSurface),
            3,
            outline,
            pickerSurface,
          );
        }
      }

      await page.keyboard.press("Escape");
      await expect(popup).toBeHidden();

      // Out of the grid, so a row still under the pointer is not measured as
      // the plain surface it is not.
      await page.mouse.move(0, 0);
      await record(
        "cell text on the cell surface",
        plainRow.locator('[tabulator-field="name"]'),
        4.5,
      );
      await record(
        "cell text on the zebra stripe",
        shadedRow.locator('[tabulator-field="name"]'),
        4.5,
      );
      await plainRow.hover();
      await record(
        "cell text on a hovered row",
        plainRow.locator('[tabulator-field="name"]'),
        4.5,
      );
      await page.mouse.move(0, 0);
      await record(
        "header text on the header surface",
        page.locator(".tabulator-col-title").first(),
        4.5,
      );

      // The open editor: the field a visitor types in, and the border that says
      // which cell is open. The border is read off the open cell rather than
      // worked out from a token, so it is measured against the row it is drawn
      // on rather than against a surface assumed for it. The click that opened
      // the editor leaves the pointer on the row, so that surface is the hover
      // one, which is the least favourable a cell wears today.
      const editing = cellOf(page, "CUST-0001", "name");
      const input = editing.locator("input");
      await expect(async () => {
        await editing.click({ timeout: 2_000 });
        await expect(input).toBeVisible({ timeout: 2_000 });
      }).toPass({ timeout: 30_000 });
      await record("the open editor's own text", input, 4.5);
      const editingSurface = await surfaceOf(page, editing);
      const editingBorder = await paintedStyle(
        page,
        editing,
        "border-top-color",
      );
      measure(
        "the open-editor border",
        contrastRatio(editingBorder, editingSurface),
        3,
        editingBorder,
        editingSurface,
      );
      await page.keyboard.press("Escape");

      // The Record ID tooltip. It is the element that actually inherited the
      // page's text colour onto a light panel before this layer, so it is
      // measured rather than assumed.
      await page.locator(".tabulator-col-title").first().hover();
      const tooltip = page.locator(
        ".tabulator-popup-container.tabulator-tooltip",
      );
      await expect(tooltip).toBeVisible();
      await record("the header tooltip's text", tooltip, 4.5);
      // A tooltip Tabulator gives no elevation keeps none: the container rule
      // above outranks its own by a class, so the layer hands it back.
      expect(await styleOf(tooltip, "box-shadow")).toBe("none");
      // Tabulator hides the tooltip on its own terms, and when it does is not
      // this spec's business: the pointer is taken off the header and the next
      // step moves to a table of its own.
      await page.mouse.move(0, 0);

      // The empty-table placeholder, which Tabulator paints in a near-white
      // grey that this layer replaces with quiet ink.
      await page.getByTestId("workspace-table-select").selectOption("Note");
      const placeholder = page.locator(".tabulator-placeholder-contents");
      await expect(placeholder).toBeVisible();
      await record("the empty-table placeholder", placeholder, 4.5);
      await page.getByTestId("workspace-table-select").selectOption("Customer");
      await expect(cellOf(page, "CUST-0001", "name")).toBeVisible();

      // The bound-ahead indicators, of which the fill handle stands here for the
      // rest: the range borders and the row header wait on the interaction work
      // that draws them, the resize guide on an option Tabulator leaves off,
      // and the refused cell on a Tabulator validator the library never lets
      // run, because it turns a bad value away first. None can be put on
      // screen, so the handle is answered by its colour alone, and against both
      // of the surfaces it straddles rather than the kinder one: it is drawn on
      // the corner of a range, half over the selected cell and half over the
      // cell outside it, which may be hovered.
      const handle = await token(page, "var(--color-fd-primary)");
      const straddled = [
        [
          "a selected cell",
          "color-mix(in srgb, var(--color-fd-accent) 60%, var(--color-fd-card))",
        ],
        [
          "a hovered cell",
          "color-mix(in srgb, var(--color-fd-foreground) 8%, var(--color-fd-card))",
        ],
      ] as const;
      for (const [where, declaration] of straddled) {
        const surface = await token(page, declaration);
        measure(
          `the range handle, on ${where}`,
          contrastRatio(handle, surface),
          3,
          handle,
          surface,
        );
      }

      // Nothing Tabulator hardcoded survives on any surface the grid paints.
      const painted = [
        await surfaceOf(page, plainRow),
        await surfaceOf(page, shadedRow),
        await paintedStyle(page, header, "background-color"),
        pickerSurface,
        chosenSurface,
        await paintedStyle(page, tick, "fill"),
        await paintedStyle(page, cross, "fill"),
        await paintedStyle(
          page,
          plainRow.locator('[tabulator-field="name"]'),
          "color",
        ),
      ];
      for (const hex of painted) {
        expect(
          TABULATOR_HEXES,
          `${hex} is one of Tabulator's own`,
        ).not.toContain(hex);
      }

      const report = [
        `mode: ${mode}`,
        `cell surface: ${card}`,
        `zebra stripe: ${stripe}`,
        `header surface: ${muted}`,
        `popup surface: ${popover}`,
        ...measurements.map(
          (entry) =>
            `${entry.ratio.toFixed(2).padStart(6)} to 1 (floor ${entry.floor}): ${entry.what}`,
        ),
      ].join("\n");
      const file = testInfo.outputPath(`grid-theme-${mode}.txt`);
      await writeFile(file, report);
      await testInfo.attach(`grid-theme-${mode}.txt`, {
        path: file,
        contentType: "text/plain",
      });
    });
  }

  test("follows the theme with no re-render of the grid", async ({ page }) => {
    await openGrid(page, "light");

    const light = await token(page, "var(--color-fd-card)");
    const row = page.locator(".tabulator-row-odd").first();
    expect(await surfaceOf(page, row)).toBe(light);

    // Mark the elements on screen, then flip the class the site's own toggle
    // flips. A grid that had to be rebuilt to change colour would lose these.
    await page.evaluate(() => {
      document
        .querySelectorAll(".tabulator-row, .tabulator-cell")
        .forEach((element, index) => {
          (element as HTMLElement).dataset["themeProbe"] = String(index);
        });
      document.documentElement.classList.add("dark");
    });

    const dark = await token(page, "var(--color-fd-card)");
    expect(dark).not.toBe(light);
    expect(await surfaceOf(page, row)).toBe(dark);
    expect(await row.getAttribute("data-theme-probe")).not.toBeNull();
    expect(
      await cellOf(page, "CUST-0001", "name").getAttribute("data-theme-probe"),
    ).not.toBeNull();

    // And back again, with the same elements.
    await page.evaluate(() => {
      document.documentElement.classList.remove("dark");
    });
    expect(await surfaceOf(page, row)).toBe(light);
    expect(await row.getAttribute("data-theme-probe")).not.toBeNull();
  });

  for (const mode of ["light", "dark"] as const) {
    test(`shows the grid in ${mode} mode`, async ({ page }, testInfo) => {
      await openGrid(page, mode);
      await page.getByTestId("workspace-grid-section").scrollIntoViewIfNeeded();

      // The picker is open in the dark shot, with the option the cell holds
      // focused, so the ring that marks it is in the evidence rather than only
      // in the numbers.
      if (mode === "dark") {
        await openPicker(page, pickerOf(page));
        await page.keyboard.press("ArrowDown");
        await page.keyboard.press("ArrowUp");
        await expect(
          page.locator(".tabulator-edit-list-item.active.focused"),
        ).toHaveCount(1);
      }

      const file = testInfo.outputPath(`grid-${mode}.png`);
      await page.screenshot({ path: file });
      await testInfo.attach(`grid-${mode}.png`, {
        path: file,
        contentType: "image/png",
      });
    });
  }
});
