import { expect, test } from "@playwright/test";
import {
  createWorkbookUpload,
  expectWorkbookDownload,
  fileInput,
  readWorkbookDownload,
  resultArtifacts,
  resultsPanel,
  VBA_PROJECT_PART,
} from "./fixtures";

const NORTH = [
  {
    name: "North",
    rows: [
      ["Client", "Amount"],
      ["Acme", 10],
    ],
  },
] as const;

test.describe("/tools/excel-merge", () => {
  test("combines two workbooks into one", async ({ page }) => {
    await page.goto("/tools/excel-merge");
    await expect(
      page.getByRole("heading", { level: 1, name: "Merge Excel workbooks" }),
    ).toBeVisible();

    const inputs = await Promise.all([
      createWorkbookUpload("north.xlsx", [
        {
          name: "North",
          rows: [
            ["Client", "Amount"],
            ["Acme", 10],
          ],
        },
      ]),
      createWorkbookUpload("south.xlsx", [
        {
          name: "South",
          rows: [
            ["Client", "Amount"],
            ["Bolt", 20],
          ],
        },
      ]),
    ]);
    await fileInput(page).setInputFiles(inputs);

    // Added files keep the order they were picked in; that order is the order
    // the merged workbook's tabs follow.
    const sources = page.getByTestId("source-item");
    await expect(sources).toHaveCount(2);
    await expect(sources.nth(0)).toContainText("north.xlsx");
    await expect(sources.nth(1)).toContainText("south.xlsx");

    await page.getByTestId("output-name-input").fill("all-sheets");
    await page.getByTestId("run-button").click();

    await expect(resultsPanel(page)).toBeVisible();
    const outputs = resultArtifacts(page);
    await expect(outputs).toHaveCount(1);
    await expect(outputs.first()).toContainText("all-sheets.xlsx");
    await expect(page.getByTestId("result-message")).toContainText(
      "SUCCESS: ConsultChimps finished your task.",
    );

    await expectWorkbookDownload(
      page,
      () => outputs.first().getByTestId("artifact-download").click(),
      "all-sheets.xlsx",
    );
  });

  test("takes a macro-enabled workbook and names the output as asked", async ({
    page,
  }) => {
    await page.goto("/tools/excel-merge");

    // The picker has to take the .xlsm before anything else can be true of
    // it, and the merge's own rule decides the output: the extension follows
    // the name the visitor types, not the inputs.
    await fileInput(page).setInputFiles(
      await createWorkbookUpload("macros.xlsm", NORTH, { macroEnabled: true }),
    );
    await expect(page.getByTestId("source-item")).toContainText("macros.xlsm");

    // Named without an extension, the merge writes .xlsx and says it dropped
    // the macro project rather than losing it quietly.
    await page.getByTestId("output-name-input").fill("all-sheets");
    await page.getByTestId("run-button").click();
    await expect(resultArtifacts(page).first()).toContainText(
      "all-sheets.xlsx",
    );
    await expect(page.getByTestId("result-message")).toContainText(
      "Removed the macro project",
    );

    // Named .xlsm, with exactly one input carrying macros, the project
    // travels into an output that says so.
    await page.getByTestId("output-name-input").fill("all-sheets.xlsm");
    await page.getByTestId("run-button").click();
    const outputs = resultArtifacts(page);
    await expect(outputs).toHaveCount(1);
    await expect(outputs.first()).toContainText("all-sheets.xlsm");

    const merged = await readWorkbookDownload(
      page,
      () => outputs.first().getByTestId("artifact-download").click(),
      "all-sheets.xlsm",
    );
    expect(merged.parts).toContain(VBA_PROJECT_PART);
  });

  test("merges in the order the list is arranged", async ({ page }) => {
    await page.goto("/tools/excel-merge");
    await fileInput(page).setInputFiles(
      await Promise.all([
        createWorkbookUpload("north.xlsx", [
          {
            name: "North",
            rows: [
              ["Client", "Amount"],
              ["Acme", 10],
            ],
          },
        ]),
        createWorkbookUpload("south.xlsx", [
          {
            name: "South",
            rows: [
              ["Client", "Amount"],
              ["Bolt", 20],
            ],
          },
        ]),
      ]),
    );

    const sources = page.getByTestId("source-item");
    await page.getByRole("button", { name: "Move south.xlsx earlier" }).click();
    await expect(sources.nth(0)).toContainText("south.xlsx");
    await expect(sources.nth(1)).toContainText("north.xlsx");

    await page.getByRole("button", { name: "Remove north.xlsx" }).click();
    await expect(sources).toHaveCount(1);

    await page.getByTestId("run-button").click();
    await expect(resultArtifacts(page)).toHaveCount(1);
    await expect(resultArtifacts(page).first()).toContainText("merged.xlsx");
  });
});
