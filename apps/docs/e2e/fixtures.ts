/**
 * Test fixtures and shared locators for the in-browser tool suite.
 *
 * PDFs are generated in-process with pdf-lib and workbooks are assembled with
 * jszip, so nothing binary is checked in and no temporary files are left
 * behind. Both are handed to the pages as in-memory uploads.
 *
 * Pages are addressed through the `data-testid` attributes the tool shell
 * renders rather than through heading text, so wording changes do not break
 * the suite. `e2e/README.md` lists the identifiers.
 */
import type { Locator, Page } from "@playwright/test";
import { expect } from "@playwright/test";
import JSZip from "jszip";
import { readFile } from "node:fs/promises";
import { PDFDocument } from "pdf-lib";

const WORKBOOK_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/** An in-memory upload in the shape Playwright's setInputFiles accepts. */
export interface UploadFile {
  readonly name: string;
  readonly mimeType: string;
  readonly buffer: Buffer;
}

/** One worksheet: a name and its rows, top-left aligned at A1. */
export interface WorksheetFixture {
  readonly name: string;
  readonly rows: ReadonlyArray<ReadonlyArray<number | string>>;
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

const XML_ESCAPES: Record<string, string> = {
  '"': "&quot;",
  "&": "&amp;",
  "'": "&apos;",
  "<": "&lt;",
  ">": "&gt;",
};

function escapeXml(value: string): string {
  return value.replaceAll(/["&'<>]/gu, (character) => XML_ESCAPES[character]!);
}

/** Spreadsheet column letters for the fixture widths this suite uses. */
function columnLetter(index: number): string {
  let remaining = index;
  let letters = "";
  do {
    letters = String.fromCodePoint(65 + (remaining % 26)) + letters;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  return letters;
}

function worksheetXml(rows: WorksheetFixture["rows"]): string {
  const body = rows
    .map((cells, rowIndex) => {
      const reference = rowIndex + 1;
      const cellXml = cells
        .map((value, columnIndex) => {
          const address = `${columnLetter(columnIndex)}${reference}`;
          return typeof value === "number"
            ? `<c r="${address}"><v>${value}</v></c>`
            : `<c r="${address}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
        })
        .join("");
      return `<row r="${reference}">${cellXml}</row>`;
    })
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`;
}

/**
 * Assembles a minimal but valid `.xlsx` package: one part per worksheet, cell
 * text stored inline so no shared-string table is needed. This is the smallest
 * thing the workbook reader accepts, which keeps every fixture a few hundred
 * bytes and every test fast.
 */
export async function createWorkbookUpload(
  name: string,
  sheets: readonly WorksheetFixture[],
): Promise<UploadFile> {
  const archive = new JSZip();

  const overrides = sheets
    .map(
      (_sheet, index) =>
        `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    )
    .join("");
  archive.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${overrides}</Types>`,
  );

  archive.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );

  const sheetEntries = sheets
    .map(
      (sheet, index) =>
        `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`,
    )
    .join("");
  archive.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheetEntries}</sheets></workbook>`,
  );

  const relationships = sheets
    .map(
      (_sheet, index) =>
        `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`,
    )
    .join("");
  archive.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`,
  );

  for (const [index, sheet] of sheets.entries()) {
    archive.file(
      `xl/worksheets/sheet${index + 1}.xml`,
      worksheetXml(sheet.rows),
    );
  }

  return {
    name,
    mimeType: WORKBOOK_MEDIA_TYPE,
    buffer: await archive.generateAsync({
      compression: "DEFLATE",
      type: "nodebuffer",
    }),
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

/** The file input a tool page renders, whichever tool it is. */
export function fileInput(page: Page): Locator {
  return page.getByTestId("file-input");
}

/**
 * The Results panel. Scoping matters on the split pages, where the preview
 * lists the same output names as the results.
 */
export function resultsPanel(page: Page): Locator {
  return page.getByTestId("results-section");
}

/** The produced outputs, in the order the operation reported them. */
export function resultArtifacts(page: Page): Locator {
  return resultsPanel(page).getByTestId("artifact-item");
}

/** The preview panel, which renders the plan for a split before it runs. */
export function previewPanel(page: Page): Locator {
  return page.getByTestId("preview-section");
}

async function downloadedBytes(
  page: Page,
  trigger: () => Promise<void>,
  expectedFileName: string,
): Promise<Buffer> {
  const downloadPromise = page.waitForEvent("download");
  await trigger();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe(expectedFileName);

  const bytes = await readFile(await download.path());
  expect(bytes.byteLength).toBeGreaterThan(0);
  return bytes;
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
  const bytes = await downloadedBytes(page, trigger, expectedFileName);
  expect(bytes.subarray(0, 5).toString("latin1")).toBe("%PDF-");
}

/**
 * Clicks a Download button and asserts the saved bytes are a ZIP package,
 * which every `.xlsx` workbook is. A tool that "finishes" while producing
 * empty or corrupt bytes fails here.
 */
export async function expectWorkbookDownload(
  page: Page,
  trigger: () => Promise<void>,
  expectedFileName: string,
): Promise<void> {
  const bytes = await downloadedBytes(page, trigger, expectedFileName);
  expect(bytes.subarray(0, 2).toString("latin1")).toBe("PK");
}
