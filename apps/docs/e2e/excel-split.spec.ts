import { expect, test } from "@playwright/test";
import {
  createTextUpload,
  createWorkbookUpload,
  expectWorkbookDownload,
  fileInput,
  previewPanel,
  readWorkbookDownload,
  resultArtifacts,
  resultsPanel,
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

  test("gives back a complete copy of the workbook", async ({ page }) => {
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
