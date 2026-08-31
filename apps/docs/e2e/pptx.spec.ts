import { expect, test, type Locator, type Page } from "@playwright/test";
import {
  createPresentationUpload,
  createTextUpload,
  createWorkbookUpload,
  expectPresentationDownload,
  previewPanel,
  resultArtifacts,
  resultsPanel,
} from "./fixtures";

/**
 * One template slide whose `{{title}}` straddles a text-run boundary, the way
 * PowerPoint splits a line the moment part of it is formatted differently. An
 * engine that read each run on its own would see unbalanced braces and refuse
 * the template, so reaching a plan at all proves the runs were stitched back
 * together before `{{field}}` was looked for.
 */
const REVIEW_TEMPLATE = [["{{ti", "tle}}, ", "{{region}}"]] as const;

/** The same slide asking for a field the records below do not carry. */
const AMOUNT_TEMPLATE = [["{{title}}, ", "{{amount}}"]] as const;

const RECORDS = [
  {
    name: "Records",
    rows: [
      ["title", "region"],
      ["Quarterly review", "North"],
      ["Quarterly review", "South"],
    ],
  },
] as const;

/**
 * The populate page renders two file pickers, so every upload is addressed
 * through its own section rather than through the shared `file-input` helper.
 */
function templateInput(page: Page): Locator {
  return page.getByTestId("template-section").getByTestId("file-input");
}

function recordsInput(page: Page): Locator {
  return page.getByTestId("records-section").getByTestId("file-input");
}

/**
 * The advanced record options sit behind a disclosure, and a closed
 * `<details>` hides its fields from Playwright. Opening it directly keeps the
 * test about the tool rather than about the disclosure widget, and does
 * nothing when the fields are already on the page.
 */
async function openAdvancedOptions(section: Locator): Promise<void> {
  for (const details of await section.locator("details").all()) {
    await details.evaluate((element: HTMLDetailsElement) => {
      element.open = true;
    });
  }
}

test.describe("/tools/pptx-populate", () => {
  test("writes one populated slide per record", async ({ page }) => {
    await page.goto("/tools/pptx-populate");
    await expect(
      page.getByRole("heading", {
        level: 1,
        name: "Populate a PowerPoint template",
      }),
    ).toBeVisible();

    const [template, records] = await Promise.all([
      createPresentationUpload("review-template.pptx", REVIEW_TEMPLATE),
      createWorkbookUpload("records.xlsx", RECORDS),
    ]);
    await templateInput(page).setInputFiles(template);
    await recordsInput(page).setInputFiles(records);

    // Both summaries appearing is the signal that the worker read each file
    // and the lazily imported engine loaded.
    await expect(page.getByTestId("template-summary")).toContainText(
      "review-template.pptx",
    );
    await expect(page.getByTestId("records-summary")).toContainText(
      "records.xlsx",
    );

    // Planning is debounced. A population always produces exactly one deck,
    // and its default name is derived from the template's.
    const preview = previewPanel(page);
    await expect(preview.getByTestId("planned-outputs")).toContainText(
      "review-template-populated.pptx",
    );
    await expect(
      preview.getByTestId("planned-outputs").getByRole("listitem"),
    ).toHaveCount(1);
    await expect(preview.getByTestId("preview-error")).toHaveCount(0);

    const recordsSection = page.getByTestId("records-section");
    await openAdvancedOptions(recordsSection);
    await recordsSection
      .getByTestId("output-name-input")
      .fill("quarterly-review");

    await page.getByTestId("run-button").click();

    await expect(resultsPanel(page)).toBeVisible();
    const outputs = resultArtifacts(page);
    await expect(outputs).toHaveCount(1);
    await expect(outputs.first().getByTestId("artifact-name")).toHaveText(
      "quarterly-review.pptx",
    );

    // The outcome text carries the operation's own counts: two records in,
    // two generated slides out.
    const message = page.getByTestId("result-message");
    await expect(message).toContainText(
      "SUCCESS: ConsultChimps finished your task.",
    );
    await expect(message).toContainText(
      "ConsultChimps read 2 nonempty Excel records and created 2 populated slides in worksheet order.",
    );

    await expectPresentationDownload(
      page,
      () => outputs.first().getByTestId("artifact-download").click(),
      "quarterly-review.pptx",
    );
  });

  test("reports a placeholder the workbook has no column for", async ({
    page,
  }) => {
    await page.goto("/tools/pptx-populate");

    const [template, records] = await Promise.all([
      createPresentationUpload("amount-template.pptx", AMOUNT_TEMPLATE),
      createWorkbookUpload("records.xlsx", RECORDS),
    ]);
    await templateInput(page).setInputFiles(template);
    await recordsInput(page).setInputFiles(records);

    // The preview surfaces the operation's own explanation, which names the
    // placeholder that has nowhere to read a value from, rather than a
    // generic failure.
    const previewError = previewPanel(page).getByTestId("preview-error");
    await expect(previewError).toBeVisible();
    await expect(previewError).toContainText(
      "ERROR: ConsultChimps could not finish your task.",
    );
    await expect(previewError).toContainText(
      'Template placeholder "amount" does not match any Excel column.',
    );

    // Nothing was produced, so the Results panel stays away entirely.
    await expect(resultsPanel(page)).toHaveCount(0);
  });

  // A number input accepts "0", "1.5", and "1e2". Reading any of them as
  // "not supplied" would populate a different row or slide than the one that
  // was typed, so each has to stop the task and say why.
  for (const { typed, why } of [
    { typed: "0", why: "is counted from 1" },
    { typed: "1.5", why: "must be a whole number" },
    { typed: "1e2", why: "must be a whole number" },
  ] as const) {
    test(`refuses the template slide "${typed}" instead of defaulting`, async ({
      page,
    }) => {
      await page.goto("/tools/pptx-populate");

      const [template, records] = await Promise.all([
        createPresentationUpload("review-template.pptx", REVIEW_TEMPLATE),
        createWorkbookUpload("records.xlsx", RECORDS),
      ]);
      await templateInput(page).setInputFiles(template);
      await recordsInput(page).setInputFiles(records);

      // A plan appears first, so the assertions below prove the bad value
      // withdrew it rather than merely never producing one.
      await expect(
        previewPanel(page).getByTestId("planned-outputs"),
      ).toBeVisible();

      const recordsSection = page.getByTestId("records-section");
      await openAdvancedOptions(recordsSection);
      await recordsSection.getByTestId("template-slide-input").fill(typed);

      await expect(
        recordsSection.getByTestId("template-slide-input-error"),
      ).toContainText(why);
      await expect(
        previewPanel(page).getByTestId("preview-invalid-options"),
      ).toContainText(why);
      await expect(
        previewPanel(page).getByTestId("planned-outputs"),
      ).toHaveCount(0);
      await expect(page.getByTestId("run-button")).toBeDisabled();
    });
  }

  test("reads records from a macro-enabled workbook", async ({ page }) => {
    await page.goto("/tools/pptx-populate");

    // The records workbook is only read, never rewritten, and the reader is
    // format-neutral - so the picker takes an .xlsm here too.
    const [template, records] = await Promise.all([
      createPresentationUpload("review-template.pptx", REVIEW_TEMPLATE),
      createWorkbookUpload("records.xlsm", RECORDS, { macroEnabled: true }),
    ]);
    await templateInput(page).setInputFiles(template);
    await recordsInput(page).setInputFiles(records);

    await expect(page.getByTestId("records-summary")).toContainText(
      "records.xlsm",
    );
    await expect(
      previewPanel(page).getByTestId("planned-outputs"),
    ).toBeVisible();

    await page.getByTestId("run-button").click();
    const outputs = resultArtifacts(page);
    await expect(outputs).toHaveCount(1);
    // The output is the presentation; the workbook only supplied the records.
    await expect(outputs.first()).toContainText(
      "review-template-populated.pptx",
    );
  });

  test("refuses a template that is not a presentation instead of failing", async ({
    page,
  }) => {
    const pageErrors: Error[] = [];
    page.on("pageerror", (error) => pageErrors.push(error));

    await page.goto("/tools/pptx-populate");
    await templateInput(page).setInputFiles(createTextUpload("notes.txt"));

    // Reading a file and re-planning the preview are both asynchronous, so
    // give them longer than the 250 ms preview debounce to produce something
    // before asserting that nothing appeared.
    await page.waitForTimeout(1_000);

    await expect(page.getByTestId("template-summary")).toHaveCount(0);
    await expect(page.getByTestId("template-rejected")).toContainText(
      "notes.txt",
    );
    await expect(page.getByTestId("run-button")).toBeDisabled();
    await expect(resultsPanel(page)).toHaveCount(0);
    expect(pageErrors).toEqual([]);
  });

  test("clears a chosen template when its replacement is rejected", async ({
    page,
  }) => {
    await page.goto("/tools/pptx-populate");

    const [template, records] = await Promise.all([
      createPresentationUpload("review-template.pptx", REVIEW_TEMPLATE),
      createWorkbookUpload("records.xlsx", RECORDS),
    ]);
    await templateInput(page).setInputFiles(template);
    await recordsInput(page).setInputFiles(records);
    await expect(
      previewPanel(page).getByTestId("planned-outputs"),
    ).toBeVisible();

    // Someone replacing the template with the wrong file must not be left
    // able to populate the old one.
    await templateInput(page).setInputFiles(createTextUpload("notes.txt"));

    await expect(page.getByTestId("template-rejected")).toContainText(
      "notes.txt",
    );
    await expect(page.getByTestId("template-summary")).toHaveCount(0);
    await expect(page.getByTestId("run-button")).toBeDisabled();
    await expect(previewPanel(page).getByTestId("planned-outputs")).toHaveCount(
      0,
    );
  });

  test("withdraws a finished deck when an option changes", async ({ page }) => {
    await page.goto("/tools/pptx-populate");

    const [template, records] = await Promise.all([
      createPresentationUpload("review-template.pptx", REVIEW_TEMPLATE),
      createWorkbookUpload("records.xlsx", RECORDS),
    ]);
    await templateInput(page).setInputFiles(template);
    await recordsInput(page).setInputFiles(records);
    await expect(
      previewPanel(page).getByTestId("planned-outputs"),
    ).toBeVisible();

    await page.getByTestId("run-button").click();
    await expect(resultArtifacts(page).first()).toContainText(
      "review-template-populated.pptx",
    );

    // A finished deck belongs to the options that made it. Changing the
    // worksheet would produce a deck with this very filename from different
    // rows, so leaving the old one downloadable beside the new preview is how
    // someone ends up shipping the wrong data under the right name.
    const recordsSection = page.getByTestId("records-section");
    await openAdvancedOptions(recordsSection);
    await recordsSection.getByTestId("worksheet-input").fill("Records");

    await expect(resultsPanel(page)).toHaveCount(0);
  });

  test("withdraws the preview while a changed option is replanned", async ({
    page,
  }) => {
    await page.goto("/tools/pptx-populate");

    const [template, records] = await Promise.all([
      createPresentationUpload("review-template.pptx", REVIEW_TEMPLATE),
      createWorkbookUpload("records.xlsx", RECORDS),
    ]);
    await templateInput(page).setInputFiles(template);
    await recordsInput(page).setInputFiles(records);

    const preview = previewPanel(page);
    await expect(preview.getByTestId("planned-outputs")).toContainText(
      "review-template-populated.pptx",
    );

    // Run would apply the new name at once, so the old plan must not stay on
    // screen describing a file the button would no longer produce.
    await page.getByTestId("output-name-input").fill("quarterly-review");
    await expect(preview.getByTestId("preview-pending")).toBeVisible();

    await expect(preview.getByTestId("planned-outputs")).toContainText(
      "quarterly-review.pptx",
    );
    await expect(preview.getByTestId("preview-pending")).toHaveCount(0);
  });
});

test.describe("/tools/pptx-inspect", () => {
  test("lists every placeholder the template slide uses", async ({ page }) => {
    await page.goto("/tools/pptx-inspect");
    await expect(
      page.getByRole("heading", {
        level: 1,
        name: "Inspect a PowerPoint template",
      }),
    ).toBeVisible();

    // `{{title}}` twice and `{{region}}` once, so the report has to count
    // occurrences rather than distinct names.
    const template = await createPresentationUpload("review-template.pptx", [
      ["{{title}}, ", "{{region}}, ", "{{title}}"],
    ]);
    await page
      .getByTestId("source-section")
      .getByTestId("file-input")
      .setInputFiles(template);
    await expect(page.getByTestId("source-summary")).toContainText(
      "review-template.pptx",
    );

    // There is no Run button here: choosing a template inspects it after the
    // usual debounce.
    const report = page.getByTestId("inspection-section");
    await expect(report.getByTestId("inspection-error")).toHaveCount(0);

    const placeholders = report.getByTestId("placeholder-item");
    await expect(placeholders).toHaveCount(2);
    await expect(placeholders.getByTestId("placeholder-name")).toContainText([
      "title",
      "region",
    ]);
    await expect(placeholders.nth(0)).toContainText("2");
    await expect(placeholders.nth(1)).toContainText("1");
  });

  test("reports what would make a populate refuse the template", async ({
    page,
  }) => {
    await page.goto("/tools/pptx-inspect");

    // One unbalanced brace, which also leaves the slide with no usable
    // placeholder: the operation's own result carries both warnings.
    const template = await createPresentationUpload("malformed.pptx", [
      ["{{title}"],
    ]);
    await page
      .getByTestId("source-section")
      .getByTestId("file-input")
      .setInputFiles(template);

    const report = page.getByTestId("inspection-section");
    await expect(report.getByTestId("inspection-malformed")).toHaveText("1");
    await expect(report.getByTestId("placeholder-item")).toHaveCount(0);

    const warnings = report.getByTestId("inspection-warning");
    await expect(warnings).toHaveCount(2);
    await expect(warnings.first()).toContainText(
      "malformed placeholder braces",
    );
    await expect(warnings.last()).toContainText(
      "does not contain any valid {{field_name}} placeholders",
    );
  });

  test("refuses a slide number that is not a whole number counted from 1", async ({
    page,
  }) => {
    await page.goto("/tools/pptx-inspect");

    const template = await createPresentationUpload("two-slides.pptx", [
      ["{{title}}"],
      ["{{region}}"],
    ]);
    await page
      .getByTestId("source-section")
      .getByTestId("file-input")
      .setInputFiles(template);

    const report = page.getByTestId("inspection-section");
    await expect(report.getByTestId("placeholder-item")).toHaveCount(1);

    await page.getByTestId("template-slide-input").fill("0");

    // Slide 1 would be a plausible-looking answer to a request nobody made,
    // so the report is withdrawn rather than quietly defaulted.
    await expect(page.getByTestId("template-slide-input-error")).toContainText(
      "is counted from 1",
    );
    await expect(report.getByTestId("inspection-invalid-slide")).toBeVisible();
    await expect(report.getByTestId("placeholder-item")).toHaveCount(0);
  });

  test("inspects the chosen template slide, not always the first", async ({
    page,
  }) => {
    await page.goto("/tools/pptx-inspect");

    const template = await createPresentationUpload("two-slides.pptx", [
      ["{{title}}"],
      ["{{region}}, ", "{{amount}}"],
    ]);
    await page
      .getByTestId("source-section")
      .getByTestId("file-input")
      .setInputFiles(template);

    const report = page.getByTestId("inspection-section");
    await expect(report.getByTestId("placeholder-item")).toHaveCount(1);

    await page.getByTestId("template-slide-input").fill("2");

    // Slide 1's placeholders must not stay on screen under a heading that now
    // says slide 2.
    await expect(report.getByTestId("inspection-pending")).toBeVisible();
    await expect(report.getByTestId("placeholder-item")).toHaveCount(0);

    await expect(report.getByTestId("placeholder-item")).toHaveCount(2);
    await expect(
      report.getByTestId("placeholder-item").getByTestId("placeholder-name"),
    ).toContainText(["region", "amount"]);
  });
});
