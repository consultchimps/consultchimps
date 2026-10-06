/**
 * The pages hand each chosen file to the worker without reading it. A page
 * that read a workbook itself would hold a whole copy in the tab, so every way
 * the page could read file bytes is made to fail here, and the operations must
 * still run and download.
 */
import { expect, test, type Page } from "@playwright/test";
import {
  createWorkbookUpload,
  expectWorkbookDownload,
  fileInput,
  resultArtifacts,
  sectionFileInput,
} from "./fixtures";

/**
 * Make the page's own file reads throw. Init scripts run in page frames, not
 * in workers, so the operation worker's reads are untouched.
 */
async function forbidPageFileReads(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const refuse = (what: string) => () => {
      throw new Error(`The page read file bytes through ${what}`);
    };
    Blob.prototype.arrayBuffer = refuse("Blob.arrayBuffer");
    Blob.prototype.bytes = refuse("Blob.bytes");
    Blob.prototype.stream = refuse("Blob.stream");
    FileReader.prototype.readAsArrayBuffer = refuse("FileReader");
    FileReader.prototype.readAsBinaryString = refuse("FileReader");
    FileReader.prototype.readAsDataURL = refuse("FileReader");
  });
}

const sheets = (prefix: string) => [
  {
    name: "Log",
    rows: [
      ["Case_ID", "Region"],
      [`${prefix}-1`, "North"],
      [`${prefix}-2`, "South"],
    ],
  },
];

test.describe("tool pages leave reading files to the worker", () => {
  test("consolidate", async ({ page }) => {
    await forbidPageFileReads(page);
    await page.goto("/tools/excel-consolidate");
    await sectionFileInput(page, "source-section").setInputFiles(
      await Promise.all([
        createWorkbookUpload("north.xlsx", sheets("N")),
        createWorkbookUpload("south.xlsx", sheets("S")),
      ]),
    );
    await page.getByTestId("run-button").click();
    await expect(resultArtifacts(page)).toHaveCount(1);
    await expectWorkbookDownload(
      page,
      () =>
        resultArtifacts(page).first().getByTestId("artifact-download").click(),
      "consolidated.xlsx",
    );
  });

  test("split", async ({ page }) => {
    await forbidPageFileReads(page);
    await page.goto("/tools/excel-split");
    await fileInput(page).setInputFiles(
      await createWorkbookUpload("log.xlsx", sheets("L")),
    );
    const columns = page.getByTestId("column-select");
    await expect(columns.getByRole("option", { name: "Region" })).toHaveCount(
      1,
    );
    await columns.selectOption("Region");
    await page.getByTestId("run-button").click();
    await expect(resultArtifacts(page)).toHaveCount(2);
    await expectWorkbookDownload(
      page,
      () =>
        resultArtifacts(page).first().getByTestId("artifact-download").click(),
      "log-North.xlsx",
    );
  });
});
