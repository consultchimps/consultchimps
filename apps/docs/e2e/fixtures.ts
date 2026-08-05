/**
 * Test fixtures for the in-browser tool suite.
 *
 * PDFs are generated in-process with pdf-lib and handed to the page as
 * in-memory uploads, so nothing binary is checked in and no temporary files
 * are left behind.
 */
import type { Page } from "@playwright/test";
import { expect } from "@playwright/test";
import { readFile } from "node:fs/promises";
import { PDFDocument } from "pdf-lib";

/** An in-memory upload in the shape Playwright's setInputFiles accepts. */
export interface UploadFile {
  readonly name: string;
  readonly mimeType: string;
  readonly buffer: Buffer;
}

/**
 * Builds the smallest useful PDF: `pageCount` blank Letter pages, no fonts
 * and no content streams, which keeps every fixture under a kilobyte.
 */
export async function createPdfUpload(
  name: string,
  pageCount: number,
): Promise<UploadFile> {
  const document = await PDFDocument.create();
  for (let index = 0; index < pageCount; index += 1) {
    document.addPage([612, 792]);
  }
  return {
    name,
    mimeType: "application/pdf",
    buffer: Buffer.from(await document.save()),
  };
}

/** A file the tools must refuse: right shape, wrong media type. */
export function createTextUpload(name: string): UploadFile {
  return {
    name,
    mimeType: "text/plain",
    buffer: Buffer.from("This is not a PDF.\n", "utf8"),
  };
}

/**
 * The Results panel, scoped by its heading. Scoping matters on the split page,
 * where the preview lists the same output names as the results.
 */
export function resultsPanel(page: Page) {
  return page
    .locator("section")
    .filter({ has: page.getByRole("heading", { name: "Results" }) });
}

/**
 * Clicks a Download button and asserts the saved bytes are a non-empty PDF.
 * Downloads arrive from a blob: URL created in the page, so this also covers
 * the artifact-to-Blob copy in the tool components.
 */
export async function expectPdfDownload(
  page: Page,
  trigger: () => Promise<void>,
  expectedFileName: string,
): Promise<void> {
  const downloadPromise = page.waitForEvent("download");
  await trigger();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe(expectedFileName);

  const bytes = await readFile(await download.path());
  expect(bytes.byteLength).toBeGreaterThan(0);
  expect(bytes.subarray(0, 5).toString("latin1")).toBe("%PDF-");
}
