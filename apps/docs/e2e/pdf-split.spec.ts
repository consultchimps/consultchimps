import { expect, test } from "@playwright/test";
import {
  createPdfUpload,
  createTextUpload,
  expectPdfDownload,
  fileInput,
  previewPanel,
  resultArtifacts,
  resultsPanel,
} from "./fixtures";

test.describe("/tools/pdf-split", () => {
  test("splits a two-page PDF into zero-padded page files", async ({
    page,
  }) => {
    await page.goto("/tools/pdf-split");
    await expect(
      page.getByRole("heading", { level: 1, name: "Split a PDF" }),
    ).toBeVisible();

    const source = await createPdfUpload("sample.pdf", 2);
    await fileInput(page).setInputFiles(source);

    // The preview re-parses the PDF in the worker behind a debounce, so its
    // planned output names are the signal that the lazily imported PDF engine
    // loaded.
    const preview = previewPanel(page);
    await expect(preview).toContainText("sample-page-001.pdf");
    await expect(preview).toContainText("sample-page-002.pdf");

    await page.getByTestId("run-button").click();

    await expect(resultsPanel(page)).toBeVisible();
    const outputs = resultArtifacts(page);
    await expect(outputs).toHaveCount(2);
    await expect(outputs.nth(0)).toContainText("sample-page-001.pdf");
    await expect(outputs.nth(1)).toContainText("sample-page-002.pdf");
    await expect(page.getByTestId("result-message")).toContainText(
      "SUCCESS: ConsultChimps finished your task.",
    );

    await expectPdfDownload(
      page,
      () => outputs.nth(0).getByTestId("artifact-download").click(),
      "sample-page-001.pdf",
    );
  });

  test("refuses a file that is not a PDF instead of failing", async ({
    page,
  }) => {
    const pageErrors: Error[] = [];
    page.on("pageerror", (error) => pageErrors.push(error));

    await page.goto("/tools/pdf-split");
    await fileInput(page).setInputFiles(createTextUpload("notes.txt"));

    // Reading a file and re-planning the preview are both asynchronous, so
    // give them longer than the 250 ms preview debounce to produce something
    // before asserting that nothing appeared.
    await page.waitForTimeout(1_000);

    // The picker drops non-PDF files, so the page stays in its initial state:
    // no source summary, nothing to run, and no results panel.
    await expect(page.getByTestId("run-button")).toBeDisabled();
    await expect(page.getByTestId("source-summary")).toHaveCount(0);
    await expect(page.getByTestId("planned-outputs")).toHaveCount(0);
    await expect(resultsPanel(page)).toHaveCount(0);
    expect(pageErrors).toEqual([]);
  });
});
