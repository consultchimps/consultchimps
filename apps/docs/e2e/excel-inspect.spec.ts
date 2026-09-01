import { expect, test, type Locator, type Page } from "@playwright/test";
import {
  createTextUpload,
  createWorkbookUpload,
  fileInput,
  resultsPanel,
  type WorksheetFixture,
} from "./fixtures";

/**
 * One workbook carrying every structure the report has to surface: a visible
 * worksheet with an Excel Table over it, a hidden worksheet that an operation
 * would skip by default, and a workbook-level named range. The amounts differ
 * per sheet so a test can tell the two apart by value alone.
 */
const REVIEW_LOG: readonly WorksheetFixture[] = [
  {
    name: "Clients",
    rows: [
      ["Region", "Owner", "Amount"],
      ["North", "Team A", 100],
      ["South", "Team B", 250],
    ],
    table: {
      columns: ["Region", "Owner", "Amount"],
      name: "ClientData",
      ref: "A1:C3",
    },
  },
  {
    name: "Archive",
    rows: [
      ["Region", "Amount"],
      ["East", 10],
    ],
    state: "hidden",
  },
];

const NAMED_RANGES = [
  { name: "ClientRange", reference: "'Clients'!$A$1:$C$3" },
] as const;

function reviewLog(name = "review-log.xlsx") {
  return createWorkbookUpload(name, REVIEW_LOG, {
    definedNames: NAMED_RANGES,
  });
}

/** The rendered description, which is the whole of this page's answer. */
function report(page: Page): Locator {
  return page.getByTestId("inspection-section");
}

test.describe("/tools/excel-inspect", () => {
  test("describes the worksheets, structures, and sample values", async ({
    page,
  }) => {
    await page.goto("/tools/excel-inspect");
    await expect(
      page.getByRole("heading", {
        level: 1,
        name: "Inspect an Excel workbook",
      }),
    ).toBeVisible();

    await fileInput(page).setInputFiles(await reviewLog());
    await expect(page.getByTestId("source-summary")).toContainText(
      "review-log.xlsx",
    );

    // There is no Run button here: choosing a workbook inspects it after the
    // usual debounce, and the report is the whole answer.
    await expect(report(page).getByTestId("inspection-error")).toHaveCount(0);
    await expect(report(page).getByTestId("inspection-worksheets")).toHaveText(
      "2",
    );
    await expect(
      report(page).getByTestId("inspection-excel-tables"),
    ).toHaveText("1");
    await expect(
      report(page).getByTestId("inspection-named-ranges"),
    ).toHaveText("1");

    // Both worksheets, in workbook order, with the hidden one called out on
    // its own row as well as in the headline above the list.
    await expect(report(page).getByTestId("worksheet-name")).toHaveText([
      "Clients",
      "Archive",
    ]);
    await expect(report(page).getByTestId("worksheet-visibility")).toHaveText([
      "Hidden",
    ]);
    await expect(
      report(page).getByTestId("hidden-worksheets-callout"),
    ).toContainText("1 worksheet is hidden");

    // The header row an operation would key on, and the rows below it.
    await expect(
      report(page).getByTestId("worksheet-item").first(),
    ).toContainText("header row 1");
    await expect(
      report(page).getByTestId("worksheet-item").first(),
    ).toContainText("2 data rows");

    const clients = report(page).getByTestId("worksheet-item").first();
    await expect(clients.getByTestId("column-header")).toHaveText([
      "Region",
      "Owner",
      "Amount",
    ]);
    // Stored values, in the order the rows carry them.
    await expect(clients.getByTestId("sample-value")).toContainText([
      "North",
      "South",
      "Team A",
      "Team B",
      "100",
      "250",
    ]);

    await expect(report(page).getByTestId("excel-table-item")).toContainText([
      "ClientData",
    ]);
    await expect(report(page).getByTestId("named-range-item")).toContainText([
      "ClientRange",
    ]);

    // An inspection creates nothing, so there is no Results panel to offer.
    await expect(resultsPanel(page)).toHaveCount(0);
  });

  test("turns the report back to what an operation would see", async ({
    page,
  }) => {
    await page.goto("/tools/excel-inspect");
    await fileInput(page).setInputFiles(await reviewLog());
    await expect(report(page).getByTestId("worksheet-item")).toHaveCount(2);

    // Hidden worksheets are described by default on this page. Turning that
    // off leaves the description with what an operation that skips them sees,
    // and the operation's own warning explains what went missing.
    await page.getByTestId("include-hidden-checkbox").uncheck();

    await expect(report(page).getByTestId("worksheet-name")).toHaveText([
      "Clients",
    ]);
    await expect(
      report(page).getByTestId("hidden-worksheets-callout"),
    ).toHaveCount(0);
    await expect(report(page).getByTestId("inspection-warning")).toContainText([
      "1 worksheet is hidden and was not described.",
    ]);
  });

  test("inspects a macro-enabled workbook", async ({ page }) => {
    await page.goto("/tools/excel-inspect");

    // The workbook is only ever read, and the package reader is
    // format-neutral, so the picker takes an .xlsm here too.
    await fileInput(page).setInputFiles(
      await createWorkbookUpload("review-log.xlsm", REVIEW_LOG, {
        definedNames: NAMED_RANGES,
        macroEnabled: true,
      }),
    );

    await expect(page.getByTestId("source-summary")).toContainText(
      "review-log.xlsm",
    );
    await expect(report(page).getByTestId("worksheet-name")).toHaveText([
      "Clients",
      "Archive",
    ]);
  });

  test("refuses a file that is not a workbook instead of failing", async ({
    page,
  }) => {
    const pageErrors: Error[] = [];
    page.on("pageerror", (error) => pageErrors.push(error));

    await page.goto("/tools/excel-inspect");
    await fileInput(page).setInputFiles(createTextUpload("notes.txt"));

    // Reading a file and inspecting it are both asynchronous, so give them
    // longer than the 250 ms debounce to produce something before asserting
    // that nothing appeared.
    await page.waitForTimeout(1_000);

    await expect(page.getByTestId("source-summary")).toHaveCount(0);
    await expect(page.getByTestId("source-rejected")).toContainText(
      "notes.txt",
    );
    await expect(report(page).getByTestId("worksheet-list")).toHaveCount(0);
    expect(pageErrors).toEqual([]);
  });

  test("describes the workbook that is actually chosen", async ({ page }) => {
    const pageErrors: Error[] = [];
    page.on("pageerror", (error) => pageErrors.push(error));

    await page.goto("/tools/excel-inspect");

    const [first, second] = await Promise.all([
      createWorkbookUpload("first.xlsx", [
        { name: "First", rows: [["Region"], ["North"]] },
      ]),
      createWorkbookUpload("second.xlsx", [
        { name: "Second", rows: [["Owner"], ["Team A"]] },
      ]),
    ]);

    // Replacing the workbook before the first report arrives withdraws that
    // inspection rather than letting it land: whichever cancellation path runs
    // (the debounce cleared, or the running scan aborted in the worker), the
    // page must never show one workbook's structure under another's name.
    await fileInput(page).setInputFiles(first);
    await fileInput(page).setInputFiles(second);

    await expect(page.getByTestId("source-summary")).toContainText(
      "second.xlsx",
    );
    await expect(report(page).getByTestId("worksheet-name")).toHaveText([
      "Second",
    ]);
    expect(pageErrors).toEqual([]);
  });
});
