import { expect, test } from "@playwright/test";
import {
  createPdfUpload,
  expectPdfDownload,
  fileInput,
  resultArtifacts,
  resultsPanel,
} from "./fixtures";

test.describe("/tools/pdf-merge", () => {
  test("merges two single-page PDFs into one document", async ({ page }) => {
    await page.goto("/tools/pdf-merge");
    await expect(
      page.getByRole("heading", { level: 1, name: "Merge PDFs" }),
    ).toBeVisible();

    const inputs = await Promise.all([
      createPdfUpload("cover.pdf", 1),
      createPdfUpload("appendix.pdf", 1),
    ]);
    await fileInput(page).setInputFiles(inputs);

    // Added files keep the order they were picked in; that order is what the
    // merge copies.
    const sources = page.getByTestId("source-item");
    await expect(sources).toHaveCount(2);
    await expect(sources.nth(0)).toContainText("cover.pdf");
    await expect(sources.nth(1)).toContainText("appendix.pdf");

    await page.getByTestId("run-button").click();

    await expect(resultsPanel(page)).toBeVisible();
    const outputs = resultArtifacts(page);
    await expect(outputs).toHaveCount(1);
    await expect(outputs.first()).toContainText("combined.pdf");
    await expect(page.getByTestId("result-message")).toContainText(
      "SUCCESS: ConsultChimps finished your task.",
    );

    await expectPdfDownload(
      page,
      () => outputs.first().getByTestId("artifact-download").click(),
      "combined.pdf",
    );
  });
});
