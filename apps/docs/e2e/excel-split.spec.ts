import { expect, test } from "@playwright/test";
import {
  createTextUpload,
  createWorkbookUpload,
  expectWorkbookDownload,
  fileInput,
  previewPanel,
  resultArtifacts,
  resultsPanel,
} from "./fixtures";

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
] as const;

test.describe("/tools/excel-split", () => {
  test("writes one workbook per distinct column value", async ({ page }) => {
    await page.goto("/tools/excel-split");
    await expect(
      page.getByRole("heading", { level: 1, name: "Split an Excel workbook" }),
    ).toBeVisible();

    const source = await createWorkbookUpload("clients.xlsx", CLIENTS);
    await fileInput(page).setInputFiles(source);

    // The dropdown is populated by reading the workbook in the worker, so its
    // options are the signal that the lazily imported engine loaded.
    const columns = page.getByTestId("column-select");
    await expect(columns.getByRole("option")).toContainText([
      "Choose a column…",
      "Client",
      "Region",
      "Amount",
      "Other (type a name)…",
    ]);

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

  test("reports a column that is not in the workbook", async ({ page }) => {
    await page.goto("/tools/excel-split");
    await fileInput(page).setInputFiles(
      await createWorkbookUpload("clients.xlsx", CLIENTS),
    );

    const columns = page.getByTestId("column-select");
    await expect(columns.getByRole("option")).toContainText([
      "Choose a column…",
      "Client",
      "Region",
      "Amount",
      "Other (type a name)…",
    ]);
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
