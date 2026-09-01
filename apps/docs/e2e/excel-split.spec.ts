import { expect, test, type Page } from "@playwright/test";
import {
  createTextUpload,
  createWorkbookUpload,
  expectWorkbookDownload,
  fileInput,
  previewPanel,
  readWorkbookDownload,
  resultArtifacts,
  resultsPanel,
  VBA_PROJECT_PART,
} from "./fixtures";

/**
 * Two worksheets, because the split's default is the all-worksheet one: only a
 * workbook with a sheet that has no split column can show that such a sheet
 * still travels into every output. Amounts are distinct per sheet so a test can
 * tell the surviving rows apart by number alone.
 */
const CLIENTS = [
  {
    name: "Clients",
    rows: [
      ["Client", "Region", "Amount"],
      ["Acme", "North", 10],
      ["Bolt", "South", 20],
      ["Cog", "North", 30],
    ],
  },
  {
    name: "Reference",
    rows: [
      ["Measure", "Value"],
      ["Regions tracked", 101],
      ["Rows reviewed", 102],
    ],
  },
] as const;

const COLUMN_OPTIONS = [
  "Choose a column…",
  "Client",
  "Region",
  "Amount",
  "Other (type a name)…",
];

/**
 * The worksheet, table, and range fields sit behind a disclosure, and a closed
 * `<details>` hides them from Playwright. Opening it directly keeps the test
 * about the tool rather than about the disclosure widget.
 */
async function openAdvancedOptions(page: Page): Promise<void> {
  // Scoped to the column section: the page's other disclosure holds the
  // inspection report, which no test of the split options wants to run.
  for (const details of await page
    .getByTestId("column-section")
    .locator("details")
    .all()) {
    await details.evaluate((element: HTMLDetailsElement) => {
      element.open = true;
    });
  }
}

test.describe("/tools/excel-split", () => {
  test("writes one workbook per distinct column value", async ({ page }) => {
    await page.goto("/tools/excel-split");
    await expect(
      page.getByRole("heading", { level: 1, name: "Split an Excel workbook" }),
    ).toBeVisible();

    const source = await createWorkbookUpload("clients.xlsx", CLIENTS);
    await fileInput(page).setInputFiles(source);

    // The dropdown is populated by reading the workbook in the worker, so its
    // options are the signal that the lazily imported engine loaded. It lists
    // the first worksheet's headers, so the second worksheet adds none.
    const columns = page.getByTestId("column-select");
    await expect(columns.getByRole("option")).toContainText(COLUMN_OPTIONS);

    await columns.selectOption("Region");

    // Planning is debounced, and the planned names are exactly the files the
    // run must produce.
    const preview = previewPanel(page);
    await expect(preview).toContainText("clients-North.xlsx");
    await expect(preview).toContainText("clients-South.xlsx");

    await page.getByTestId("run-button").click();

    await expect(resultsPanel(page)).toBeVisible();
    const outputs = resultArtifacts(page);
    await expect(outputs).toHaveCount(2);
    await expect(outputs.nth(0)).toContainText("clients-North.xlsx");
    await expect(outputs.nth(1)).toContainText("clients-South.xlsx");
    await expect(page.getByTestId("result-message")).toContainText(
      "SUCCESS: ConsultChimps finished your task.",
    );

    await expectWorkbookDownload(
      page,
      () => outputs.nth(0).getByTestId("artifact-download").click(),
      "clients-North.xlsx",
    );
  });

  test("keeps every worksheet and removes only the other values' rows", async ({
    page,
  }) => {
    await page.goto("/tools/excel-split");
    await fileInput(page).setInputFiles(
      await createWorkbookUpload("clients.xlsx", CLIENTS),
    );

    const columns = page.getByTestId("column-select");
    await expect(columns.getByRole("option")).toContainText(COLUMN_OPTIONS);
    await columns.selectOption("Region");
    await expect(previewPanel(page)).toContainText("clients-North.xlsx");

    await page.getByTestId("run-button").click();
    await expect(resultsPanel(page)).toBeVisible();
    const outputs = resultArtifacts(page);
    await expect(outputs).toHaveCount(2);

    const north = await readWorkbookDownload(
      page,
      () => outputs.nth(0).getByTestId("artifact-download").click(),
      "clients-North.xlsx",
    );

    // Every worksheet of the source is present, in the source's order.
    expect(north.sheetNames).toEqual(["Clients", "Reference"]);

    // The worksheet carrying the column keeps its header and only the rows for
    // this output: Acme (10) and Cog (30) stay, Bolt (20) is gone. Row numbers
    // are not asserted, because retained rows keep their original numbers on a
    // plain worksheet and are compacted inside an Excel Table.
    const clients = north.sheet("Clients");
    expect(clients.rowNumbers).toHaveLength(3);
    expect(clients.rowNumbers[0]).toBe(1);
    expect(clients.numbers).toEqual([10, 30]);

    // The worksheet without the column travels through untouched.
    const reference = north.sheet("Reference");
    expect(reference.rowNumbers).toEqual([1, 2, 3]);
    expect(reference.numbers).toEqual([101, 102]);
  });

  test("ignores the options the all-worksheet split does not use", async ({
    page,
  }) => {
    await page.goto("/tools/excel-split");
    await fileInput(page).setInputFiles(
      await createWorkbookUpload("clients.xlsx", CLIENTS),
    );

    const columns = page.getByTestId("column-select");
    await expect(columns.getByRole("option")).toContainText(COLUMN_OPTIONS);
    await columns.selectOption("Region");

    // Strict matching is offered here exactly as it is on the command line.
    await expect(page.getByTestId("strict-checkbox")).toBeEnabled();

    // Blank split values never produce an output and hidden worksheets are
    // always filtered, so neither option can change an all-worksheet split.
    await expect(page.getByTestId("include-blank-checkbox")).toBeDisabled();
    await expect(page.getByTestId("include-hidden-checkbox")).toBeDisabled();
  });

  test("reports a column that is not in the workbook", async ({ page }) => {
    await page.goto("/tools/excel-split");
    await fileInput(page).setInputFiles(
      await createWorkbookUpload("clients.xlsx", CLIENTS),
    );

    const columns = page.getByTestId("column-select");
    await expect(columns.getByRole("option")).toContainText(COLUMN_OPTIONS);
    await columns.selectOption({ label: "Other (type a name)…" });
    await page.getByTestId("column-input").fill("Territory");

    // The preview surfaces the operation's own explanation rather than a
    // generic failure, and the run button stays available for another try.
    await expect(page.getByTestId("preview-error")).toContainText(
      "ERROR: ConsultChimps could not finish your task.",
    );
    await expect(page.getByTestId("run-button")).toBeEnabled();
  });

  test("splits a macro-enabled workbook into .xlsm outputs", async ({
    page,
  }) => {
    await page.goto("/tools/excel-split");

    // A genuine .xlsm arrives with its own media type, which is what the
    // page's picker and drop target have to recognise before any operation
    // can see the file at all.
    await fileInput(page).setInputFiles(
      await createWorkbookUpload("clients.xlsm", CLIENTS, {
        macroEnabled: true,
      }),
    );
    await expect(page.getByTestId("source-summary")).toContainText(
      "clients.xlsm",
    );

    const columns = page.getByTestId("column-select");
    await expect(columns.getByRole("option")).toContainText(COLUMN_OPTIONS);
    await columns.selectOption("Region");

    // The split takes its output extension from the source, so a macro
    // workbook must not be handed back named .xlsx.
    const preview = previewPanel(page);
    await expect(preview).toContainText("clients-North.xlsm");
    await expect(preview).toContainText("clients-South.xlsm");

    await page.getByTestId("run-button").click();
    await expect(resultsPanel(page)).toBeVisible();
    const outputs = resultArtifacts(page);
    await expect(outputs).toHaveCount(2);
    await expect(outputs.nth(0)).toContainText("clients-North.xlsm");

    const north = await readWorkbookDownload(
      page,
      () => outputs.nth(0).getByTestId("artifact-download").click(),
      "clients-North.xlsm",
    );
    // The package is preserved, so the macro project travels with it: an
    // output that lost the part would be an .xlsm in name only.
    expect(north.parts).toContain(VBA_PROJECT_PART);
    expect(north.sheetNames).toEqual(["Clients", "Reference"]);
    expect(north.sheet("Clients").numbers).toEqual([10, 30]);
  });

  test("rebuilds .xlsx files when a macro workbook is narrowed to one sheet", async ({
    page,
  }) => {
    await page.goto("/tools/excel-split");
    await fileInput(page).setInputFiles(
      await createWorkbookUpload("clients.xlsm", CLIENTS, {
        macroEnabled: true,
      }),
    );

    const columns = page.getByTestId("column-select");
    await expect(columns.getByRole("option")).toContainText(COLUMN_OPTIONS);
    await columns.selectOption("Region");
    await expect(previewPanel(page)).toContainText("clients-North.xlsm");

    // Naming a worksheet asks for the compact single-source rebuild, which
    // writes a new workbook rather than preserving the package. There is no
    // macro project in the result, so the outputs are .xlsx - which is what the
    // page's filename hint and the guide say.
    await openAdvancedOptions(page);
    await page.getByTestId("sheet-input").fill("Clients");
    await expect(previewPanel(page)).toContainText("clients-North.xlsx");
    await expect(previewPanel(page)).not.toContainText("clients-North.xlsm");
  });

  test("refuses a package whose type contradicts its name", async ({
    page,
  }) => {
    await page.goto("/tools/excel-split");

    // A macro-enabled package renamed .xlsx. The split names its outputs from
    // the source name and preserves the source package, so every output would
    // be an .xlsx carrying a macro project the name denies.
    await fileInput(page).setInputFiles(
      await createWorkbookUpload("renamed.xlsx", CLIENTS, {
        macroEnabled: true,
        mimeType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );

    const columns = page.getByTestId("column-select");
    await expect(columns.getByRole("option")).toContainText(COLUMN_OPTIONS);
    await columns.selectOption("Region");

    // The preview refuses before promising any output, and says which side to
    // correct rather than only that something failed.
    const previewError = page.getByTestId("preview-error");
    await expect(previewError).toContainText(
      "is a macro-enabled workbook but is named",
    );
    await expect(previewError).toContainText(
      "XLSX_SPLIT_PACKAGE_TYPE_MISMATCH",
    );
    await expect(previewPanel(page).getByTestId("planned-outputs")).toHaveCount(
      0,
    );

    // Running anyway fails the same way rather than producing mislabelled
    // files.
    await page.getByTestId("run-button").click();
    await expect(page.getByTestId("failure-message")).toContainText(
      "XLSX_SPLIT_PACKAGE_TYPE_MISMATCH",
    );
    await expect(resultArtifacts(page)).toHaveCount(0);
  });

  test("refuses an ordinary workbook named .xlsm", async ({ page }) => {
    await page.goto("/tools/excel-split");

    // The other direction: ordinary bytes under a macro-enabled name, which
    // the picker now accepts and the operation still refuses.
    await fileInput(page).setInputFiles(
      await createWorkbookUpload("renamed.xlsm", CLIENTS, {
        mimeType: "application/vnd.ms-excel.sheet.macroEnabled.12",
      }),
    );
    await expect(page.getByTestId("source-summary")).toContainText(
      "renamed.xlsm",
    );

    const columns = page.getByTestId("column-select");
    await expect(columns.getByRole("option")).toContainText(COLUMN_OPTIONS);
    await columns.selectOption("Region");

    const previewError = page.getByTestId("preview-error");
    await expect(previewError).toContainText(
      "is an ordinary Excel workbook with no macro project",
    );
    await expect(previewError).toContainText(
      "XLSX_SPLIT_PACKAGE_TYPE_MISMATCH",
    );
  });

  test("looks inside the chosen workbook without leaving the page", async ({
    page,
  }) => {
    await page.goto("/tools/excel-split");
    await fileInput(page).setInputFiles(
      await createWorkbookUpload("clients.xlsx", [
        ...CLIENTS,
        { name: "Old", rows: [["Client"], ["Dyne"]], state: "hidden" },
      ]),
    );

    // The report is folded away, and not mounted, until someone asks for it.
    await expect(page.getByTestId("inspection-section")).toHaveCount(0);
    await page.getByTestId("inspector-disclosure").locator("summary").click();

    await expect(page.getByTestId("inspection-section")).toBeVisible();
    await expect(page.getByTestId("worksheet-item")).toHaveCount(3);
    await expect(page.getByTestId("worksheet-name")).toHaveText([
      "Clients",
      "Reference",
      "Old",
    ]);
    // Hidden worksheets are described here, because the default split filters
    // them too, and each one carries its visibility.
    await expect(page.getByTestId("worksheet-visibility")).toHaveText([
      "Hidden",
    ]);
    await expect(
      page.getByTestId("worksheet-item").first().getByTestId("column-header"),
    ).toHaveText(["Client", "Region", "Amount"]);
  });

  test("refuses a file that is not a workbook instead of failing", async ({
    page,
  }) => {
    const pageErrors: Error[] = [];
    page.on("pageerror", (error) => pageErrors.push(error));

    await page.goto("/tools/excel-split");
    await fileInput(page).setInputFiles(createTextUpload("notes.txt"));

    // Reading a file and re-planning the preview are both asynchronous, so
    // give them longer than the 250 ms preview debounce to produce something
    // before asserting that nothing appeared.
    await page.waitForTimeout(1_000);

    await expect(page.getByTestId("run-button")).toBeDisabled();
    await expect(page.getByTestId("source-summary")).toHaveCount(0);
    await expect(page.getByTestId("column-select")).toHaveCount(0);
    await expect(resultsPanel(page)).toHaveCount(0);
    expect(pageErrors).toEqual([]);
  });
});
