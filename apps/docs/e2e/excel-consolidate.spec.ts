import { expect, test } from "@playwright/test";
import {
  createWorkbookUpload,
  expectWorkbookDownload,
  fileInput,
  resultArtifacts,
  resultsPanel,
  type UploadFile,
} from "./fixtures";

/**
 * Two exports of one review log whose headers drifted apart: "Failed Checks"
 * in one workbook, "Failed_Checks" in the other. Consolidation is the
 * operation that reconciles them, and the "Normalize headers" checkbox is what
 * decides whether it does.
 */
function driftedReviewLogs(): Promise<UploadFile[]> {
  return Promise.all([
    createWorkbookUpload("north.xlsx", [
      {
        name: "Review Log",
        rows: [
          ["Case_ID", "Failed Checks"],
          ["R-1", 5],
        ],
      },
    ]),
    createWorkbookUpload("south.xlsx", [
      {
        name: "vF",
        rows: [
          ["Case_ID", "Failed_Checks"],
          ["R-2", 7],
        ],
      },
    ]),
  ]);
}

test.describe("/tools/excel-consolidate", () => {
  test("takes a macro-enabled workbook and still writes an .xlsx", async ({
    page,
  }) => {
    await page.goto("/tools/excel-consolidate");

    await fileInput(page).setInputFiles(
      await createWorkbookUpload(
        "north.xlsm",
        [
          {
            name: "Review Log",
            rows: [
              ["Case_ID", "Failed Checks"],
              ["R-1", 5],
            ],
          },
        ],
        { macroEnabled: true },
      ),
    );
    await expect(page.getByTestId("source-item")).toContainText("north.xlsm");

    await page.getByTestId("run-button").click();

    // Consolidation writes a fresh table workbook rather than preserving a
    // source package, so there is nothing macro-enabled to carry: the output
    // is an .xlsx whatever the inputs were named.
    const outputs = resultArtifacts(page);
    await expect(outputs).toHaveCount(1);
    await expect(outputs.first()).toContainText("consolidated.xlsx");
    await expectWorkbookDownload(
      page,
      () => outputs.first().getByTestId("artifact-download").click(),
      "consolidated.xlsx",
    );
  });

  test("stacks rows from both workbooks into one table", async ({ page }) => {
    await page.goto("/tools/excel-consolidate");
    await expect(
      page.getByRole("heading", {
        level: 1,
        name: "Consolidate Excel workbooks",
      }),
    ).toBeVisible();

    await fileInput(page).setInputFiles(await driftedReviewLogs());

    const sources = page.getByTestId("source-item");
    await expect(sources).toHaveCount(2);
    await expect(sources.nth(0)).toContainText("north.xlsx");
    await expect(sources.nth(1)).toContainText("south.xlsx");

    await page.getByTestId("output-name-input").fill("review-log");
    await page.getByTestId("run-button").click();

    await expect(resultsPanel(page)).toBeVisible();
    const outputs = resultArtifacts(page);
    await expect(outputs).toHaveCount(1);
    await expect(outputs.first()).toContainText("review-log.xlsx");

    // Left alone, the two spellings stay separate columns, and the three
    // provenance columns are added: two data rows across six columns.
    const message = page.getByTestId("result-message");
    await expect(message).toContainText("SUCCESS:");
    await expect(message).toContainText("combined 2 visible worksheets");
    await expect(message).toContainText(
      "2 data rows arranged across 6 columns",
    );

    // One output file, so the zip bundle is not offered.
    await expect(page.getByTestId("archive-download")).toHaveCount(0);

    await expectWorkbookDownload(
      page,
      () => outputs.first().getByTestId("artifact-download").click(),
      "review-log.xlsx",
    );
  });

  test("normalizes drifted headers and can drop the source columns", async ({
    page,
  }) => {
    await page.goto("/tools/excel-consolidate");
    await fileInput(page).setInputFiles(await driftedReviewLogs());

    await page.getByTestId("normalize-headers-checkbox").check();
    await page.getByTestId("source-columns-checkbox").uncheck();
    await page.getByTestId("run-button").click();

    await expect(resultArtifacts(page)).toHaveCount(1);
    await expect(resultArtifacts(page).first()).toContainText(
      "consolidated.xlsx",
    );
    // "Failed Checks" and "Failed_Checks" collapse into one column, and
    // without the provenance columns only the two data columns remain.
    await expect(page.getByTestId("result-message")).toContainText(
      "2 data rows arranged across 2 columns",
    );
  });
});
