import { expect, test, type Page } from "@playwright/test";
import {
  createMappingUpload,
  createWorkbookUpload,
  expectWorkbookDownload,
  readTextDownload,
  readWorkbookDownload,
  resultArtifacts,
  resultsPanel,
  sectionFileInput,
  type UploadFile,
} from "./fixtures";

/** The page takes workbooks and a mapping, so each picker is addressed here. */
function workbookInput(page: Page) {
  return sectionFileInput(page, "source-section");
}

function mappingInput(page: Page) {
  return sectionFileInput(page, "mapping-section");
}

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

/**
 * Two exports whose columns carry genuinely different names for one field,
 * which is what a mapping is for: normalized matching would leave "Reference"
 * and "Case_ID" as two columns however hard it squinted. "Notes" is claimed by
 * nothing, so it exercises the unmapped passthrough and its warning.
 */
function differentlyNamedExports(): Promise<UploadFile[]> {
  return Promise.all([
    createWorkbookUpload("north.xlsx", [
      {
        name: "Log",
        rows: [
          ["Case_ID", "Amount", "Notes"],
          ["R-1", 5, "first"],
        ],
      },
    ]),
    createWorkbookUpload("south.xlsx", [
      {
        name: "vF",
        rows: [
          ["Reference", "Total"],
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

    await workbookInput(page).setInputFiles(
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

    await workbookInput(page).setInputFiles(await driftedReviewLogs());

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
    await workbookInput(page).setInputFiles(await driftedReviewLogs());

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

  test("inspects any one of the added workbooks", async ({ page }) => {
    await page.goto("/tools/excel-consolidate");
    await workbookInput(page).setInputFiles(await driftedReviewLogs());

    // Folded away until asked for, and not even mounted before that: the
    // report is the longest thing this page can show.
    await expect(page.getByTestId("inspection-section")).toHaveCount(0);
    await page.getByTestId("inspector-disclosure").locator("summary").click();
    await expect(page.getByTestId("inspection-section")).toBeVisible();

    await page.getByTestId("inspect-select").selectOption({ index: 2 });
    await expect(page.getByTestId("worksheet-item")).toHaveCount(1);
    await expect(page.getByTestId("worksheet-name")).toHaveText("vF");
    await expect(page.getByTestId("column-header")).toHaveText([
      "Case_ID",
      "Failed_Checks",
    ]);
  });

  test("applies a column mapping and warns about the columns it does not claim", async ({
    page,
  }) => {
    await page.goto("/tools/excel-consolidate");
    await workbookInput(page).setInputFiles(await differentlyNamedExports());
    await mappingInput(page).setInputFiles(
      createMappingUpload("mapping.json", {
        version: 1,
        columns: [
          { name: "Case_ID", aliases: ["Reference"] },
          { name: "Amount", aliases: ["Total"] },
        ],
      }),
    );

    await expect(page.getByTestId("mapping-summary")).toContainText(
      "mapping.json",
    );
    await expect(page.getByTestId("mapping-columns")).toHaveText(
      "2 canonical columns, 2 aliases",
    );

    await page.getByTestId("run-button").click();
    await expect(resultsPanel(page)).toBeVisible();

    const message = page.getByTestId("result-message");
    // "Reference" folded into "Case_ID" and "Total" into "Amount", so the two
    // exports share two columns instead of four; "Notes" passed through under
    // its own name and is named in a warning.
    await expect(message).toContainText(
      "2 data rows arranged across 6 columns",
    );
    await expect(message).toContainText(
      "1 column did not match the column mapping and kept its own name",
    );
    await expect(message).toContainText('"Notes"');

    const workbook = await readWorkbookDownload(
      page,
      () =>
        resultArtifacts(page).first().getByTestId("artifact-download").click(),
      "consolidated.xlsx",
    );
    expect(workbook.sheet("Consolidated").headers).toEqual([
      "Case_ID",
      "Amount",
      "Notes",
      "_source_file",
      "_source_sheet",
      "_source_row",
    ]);
  });

  test("refuses an unusable mapping before any workbook is read", async ({
    page,
  }) => {
    await page.goto("/tools/excel-consolidate");
    await workbookInput(page).setInputFiles(await driftedReviewLogs());
    await mappingInput(page).setInputFiles(
      createMappingUpload("ambiguous.json", {
        version: 1,
        columns: [
          { name: "Case ID", aliases: [] },
          { name: "case_id", aliases: [] },
        ],
      }),
    );

    // Matching is normalized, so those two canonical columns claim the same
    // source headers. The refusal arrives on selection, with its stable code.
    const failure = page.getByTestId("mapping-error");
    await expect(failure).toContainText("TABLE_MAPPING_INVALID");
    await expect(failure).toContainText("Keep one of them");
    await expect(page.getByTestId("mapping-columns")).toHaveCount(0);
    // A mapping that cannot be read is not one the run may quietly ignore.
    await expect(page.getByTestId("run-button")).toBeDisabled();

    await page.getByTestId("mapping-remove").click();
    await expect(page.getByTestId("mapping-error")).toHaveCount(0);
    await expect(page.getByTestId("run-button")).toBeEnabled();
  });

  test("drafts a mapping for review and hands back the reviewed document", async ({
    page,
  }) => {
    await page.goto("/tools/excel-consolidate");
    await workbookInput(page).setInputFiles(await driftedReviewLogs());

    await page.getByTestId("suggest-button").click();

    // Only the headers that already match once case, spacing, and punctuation
    // are set aside: "Case_ID" is spelled the same way in both workbooks, so
    // it is not a group.
    const groups = page.getByTestId("suggestion-group");
    await expect(groups).toHaveCount(1);
    await expect(page.getByTestId("suggestion-spellings")).toHaveText(
      "Failed Checks, Failed_Checks",
    );
    await expect(page.getByTestId("suggestion-evidence")).toHaveText(
      "Seen in 2 worksheets across 2 workbooks",
    );
    // Nothing is applied by drafting: the run is still the unmapped run.
    await expect(page.getByTestId("mapping-summary")).toHaveCount(0);

    const drafted = JSON.parse(
      await readTextDownload(
        page,
        () => page.getByTestId("suggestion-download").click(),
        "mapping-draft.json",
      ),
    ) as unknown;
    expect(drafted).toEqual({
      version: 1,
      columns: [{ name: "Failed Checks", aliases: [] }],
    });

    // Renaming the canonical column keeps the group reachable: the new name
    // normalizes differently, so one of the folded spellings travels with it.
    await page.getByTestId("suggestion-canonical").fill("Checks Failed");
    const renamed = JSON.parse(
      await readTextDownload(
        page,
        () => page.getByTestId("suggestion-download").click(),
        "mapping-draft.json",
      ),
    ) as unknown;
    expect(renamed).toEqual({
      version: 1,
      columns: [{ name: "Checks Failed", aliases: ["Failed Checks"] }],
    });
  });

  test("applies a drafted mapping only once it is added back", async ({
    page,
  }) => {
    await page.goto("/tools/excel-consolidate");
    await workbookInput(page).setInputFiles(await driftedReviewLogs());

    await page.getByTestId("suggest-button").click();
    await expect(page.getByTestId("suggestion-group")).toHaveCount(1);
    const drafted = await readTextDownload(
      page,
      () => page.getByTestId("suggestion-download").click(),
      "mapping-draft.json",
    );

    await mappingInput(page).setInputFiles({
      name: "mapping-draft.json",
      mimeType: "application/json",
      buffer: Buffer.from(drafted, "utf8"),
    });
    await expect(page.getByTestId("mapping-columns")).toHaveText(
      "1 canonical column, 0 aliases",
    );

    await page.getByTestId("run-button").click();
    // The two spellings the draft grouped now share one output column: five
    // columns instead of the six an unmapped run produced.
    await expect(page.getByTestId("result-message")).toContainText(
      "2 data rows arranged across 5 columns",
    );
    const workbook = await readWorkbookDownload(
      page,
      () =>
        resultArtifacts(page).first().getByTestId("artifact-download").click(),
      "consolidated.xlsx",
    );
    expect(workbook.sheet("Consolidated").headers).toEqual([
      "Case_ID",
      "Failed Checks",
      "_source_file",
      "_source_sheet",
      "_source_row",
    ]);
  });
});
