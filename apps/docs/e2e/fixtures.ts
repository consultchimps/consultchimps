/**
 * Test fixtures and shared locators for the in-browser tool suite.
 *
 * PDFs are generated in-process with pdf-lib, and workbooks and presentations
 * are assembled with jszip, so nothing binary is checked in and no temporary
 * files are left behind. All three are handed to the pages as in-memory
 * uploads.
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

const PRESENTATION_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";

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

function slideXml(runs: ReadonlyArray<string>): string {
  const shape = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${runs
    .map(
      (text) =>
        `<a:r><a:rPr lang="en-US"/><a:t xml:space="preserve">${escapeXml(text)}</a:t></a:r>`,
    )
    .join("")}</a:p></p:txBody></p:sp>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shape}</p:spTree></p:cSld></p:sld>`;
}

/**
 * Assembles a minimal but valid `.pptx` package: one text shape per slide,
 * whose single paragraph carries one `<a:r>` run per string given for that
 * slide. Splitting the placeholder text across runs the way PowerPoint does is
 * the point — the populate engine has to stitch runs back together before it
 * can see a `{{field}}`, so a checked-in deck would hide that behaviour behind
 * an opaque binary. Generating the package here keeps every fixture a few
 * hundred bytes, keeps the run layout visible in the test that depends on it,
 * and leaves nothing binary in the repository.
 *
 * Parts are written with `createFolders: false`, which is how PowerPoint
 * writes a package: entries for parts only, never for directories.
 */
export async function createPresentationUpload(
  name: string,
  slides: ReadonlyArray<ReadonlyArray<string>>,
): Promise<UploadFile> {
  const archive = new JSZip();
  const write = (partPath: string, content: string): void => {
    archive.file(partPath, content, { createFolders: false });
  };

  const overrides = slides
    .map(
      (_slide, index) =>
        `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
    )
    .join("");
  write(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${overrides}</Types>`,
  );

  write(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
  );

  const slideReferences = slides
    .map(
      (_slide, index) =>
        `<p:sldId id="${256 + index}" r:id="rId${index + 1}"/>`,
    )
    .join("");
  write(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst>${slideReferences}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`,
  );

  const relationships = slides
    .map(
      (_slide, index) =>
        `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index + 1}.xml"/>`,
    )
    .join("");
  write(
    "ppt/_rels/presentation.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`,
  );

  for (const [index, runs] of slides.entries()) {
    write(`ppt/slides/slide${index + 1}.xml`, slideXml(runs));
  }

  return {
    name,
    mimeType: PRESENTATION_MEDIA_TYPE,
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
 * Clicks a Download button and asserts the saved bytes are a ZIP package. Both
 * Office formats the tools produce are ZIP containers, so a `.xlsx` workbook
 * and a `.pptx` presentation share one header check: a tool that "finishes"
 * while producing empty or corrupt bytes fails here.
 */
async function expectOfficePackageDownload(
  page: Page,
  trigger: () => Promise<void>,
  expectedFileName: string,
): Promise<void> {
  const bytes = await downloadedBytes(page, trigger, expectedFileName);
  expect(bytes.subarray(0, 2).toString("latin1")).toBe("PK");
}

/** Asserts a downloaded `.xlsx` workbook is a non-empty ZIP package. */
export const expectWorkbookDownload = expectOfficePackageDownload;

/** Asserts a downloaded `.pptx` presentation is a non-empty ZIP package. */
export const expectPresentationDownload = expectOfficePackageDownload;

/** One worksheet of a downloaded workbook, read back from its own part. */
export interface DownloadedWorksheet {
  /** The worksheet's tab name. */
  readonly name: string;
  /** The part's raw XML, for assertions this helper does not cover. */
  readonly xml: string;
  /** The Excel row numbers the part still declares, in document order. */
  readonly rowNumbers: readonly number[];
  /** Every numeric cell value in the part, in document order. */
  readonly numbers: readonly number[];
}

/** A downloaded workbook, addressed by worksheet tab name. */
export interface DownloadedWorkbook {
  /** Every worksheet tab name, in workbook order. */
  readonly sheetNames: readonly string[];
  /** The named worksheet; fails the test when the workbook has no such tab. */
  sheet(name: string): DownloadedWorksheet;
}

const XML_UNESCAPES: Record<string, string> = {
  "&amp;": "&",
  "&apos;": "'",
  "&gt;": ">",
  "&lt;": "<",
  "&quot;": '"',
};

function unescapeXml(value: string): string {
  return value.replaceAll(
    /&(?:amp|apos|gt|lt|quot);/gu,
    (entity) => XML_UNESCAPES[entity]!,
  );
}

function attribute(element: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`, "u").exec(element)?.[1];
}

function rowNumbersOf(xml: string): number[] {
  return [...xml.matchAll(/<row\b[^>]*>/gu)].flatMap((match) => {
    const reference = attribute(match[0], "r");
    return reference === undefined ? [] : [Number(reference)];
  });
}

/**
 * Numeric cell values only. Cells carrying a `t` attribute hold text — either
 * inline or through the shared-string table — so skipping them keeps these
 * assertions independent of how the writer chose to store strings.
 */
function numbersOf(xml: string): number[] {
  return [...xml.matchAll(/<c\b([^>]*)>(.*?)<\/c>/gsu)].flatMap((match) => {
    const type = attribute(match[1]!, "t");
    if (type !== undefined && type !== "n") {
      return [];
    }
    const value = /<v>([^<]*)<\/v>/u.exec(match[2]!)?.[1];
    return value === undefined || value.trim() === "" ? [] : [Number(value)];
  });
}

async function readPart(archive: JSZip, path: string): Promise<string> {
  const part = archive.file(path);
  expect(part, `the workbook is missing ${path}`).not.toBeNull();
  return part!.async("string");
}

/**
 * Clicks a Download button and reads the saved workbook back, so a test can
 * assert what actually reached the user rather than only that bytes arrived.
 * Worksheets are resolved through the workbook's own relationships, so the
 * assertions do not depend on which part file a worksheet landed in.
 */
export async function readWorkbookDownload(
  page: Page,
  trigger: () => Promise<void>,
  expectedFileName: string,
): Promise<DownloadedWorkbook> {
  const bytes = await downloadedBytes(page, trigger, expectedFileName);
  expect(bytes.subarray(0, 2).toString("latin1")).toBe("PK");

  const archive = await JSZip.loadAsync(bytes);
  const relationships = new Map<string, string>();
  for (const match of (
    await readPart(archive, "xl/_rels/workbook.xml.rels")
  ).matchAll(/<Relationship\b[^>]*>/gu)) {
    const id = attribute(match[0], "Id");
    const target = attribute(match[0], "Target");
    if (id !== undefined && target !== undefined) {
      relationships.set(id, target);
    }
  }

  const worksheets = new Map<string, DownloadedWorksheet>();
  const sheetNames: string[] = [];
  for (const match of (await readPart(archive, "xl/workbook.xml")).matchAll(
    /<sheet\b[^>]*>/gu,
  )) {
    const name = unescapeXml(attribute(match[0], "name") ?? "");
    const target = relationships.get(attribute(match[0], "r:id") ?? "");
    expect(target, `no worksheet part is related to "${name}"`).toBeDefined();
    const path = target!.startsWith("/")
      ? target!.slice(1)
      : `xl/${target!.replace(/^\.\//u, "")}`;
    const xml = await readPart(archive, path);
    sheetNames.push(name);
    worksheets.set(name, {
      name,
      numbers: numbersOf(xml),
      rowNumbers: rowNumbersOf(xml),
      xml,
    });
  }

  return {
    sheet(name: string): DownloadedWorksheet {
      const worksheet = worksheets.get(name);
      expect(
        worksheet,
        `the workbook has no worksheet named "${name}"`,
      ).toBeDefined();
      return worksheet!;
    },
    sheetNames,
  };
}
