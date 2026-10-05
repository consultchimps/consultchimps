/**
 * A test-only builder that writes a minimal valid .xlsx from a small spec.
 * It imports nothing but JSZip, so the cli and pptx tests can import it by
 * relative path. The output is deterministic: no timestamps, fixed part order.
 */
import JSZip from "jszip";

/** A cell that needs more than a plain value. */
export interface FixtureCell {
  /** The stored value, or the cached result of `formula`. */
  readonly value?: string | number | boolean | null;
  /** A number format code, or a built-in format id such as 14. */
  readonly format?: string | number;
  /** A formula without the leading "=". */
  readonly formula?: string;
  /** An error value such as "#DIV/0!", stored or cached. */
  readonly error?: string;
}

export type FixtureValue = string | number | boolean | null | FixtureCell;

export interface FixtureSheet {
  readonly name: string;
  readonly rows: readonly (readonly FixtureValue[])[];
  /** Merged ranges such as "A1:C1". */
  readonly merges?: readonly string[];
  readonly state?: "hidden" | "veryHidden";
  /** Column widths in characters, from column A. */
  readonly widths?: readonly number[];
}

export interface FixtureDefinedName {
  readonly name: string;
  /** The formula the name stands for, such as "Data!$A$1:$B$3". */
  readonly reference: string;
  /** The zero-based sheet the name is local to; omit for a workbook name. */
  readonly localSheet?: number;
  readonly hidden?: boolean;
}

export interface FixtureWorkbook {
  readonly sheets: readonly FixtureSheet[];
  readonly names?: readonly FixtureDefinedName[];
  /** Count dates from 1904 rather than 1900. */
  readonly date1904?: boolean;
}

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT = "application/vnd.openxmlformats-officedocument.spreadsheetml";
const HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;

function escapeXml(text: string): string {
  return text
    .replace(/&/gu, "&amp;")
    .replace(/</gu, "&lt;")
    .replace(/>/gu, "&gt;")
    .replace(/"/gu, "&quot;");
}

function columnName(index: number): string {
  let name = "";
  for (let rest = index + 1; rest > 0; rest = Math.floor((rest - 1) / 26)) {
    name = String.fromCharCode(65 + ((rest - 1) % 26)) + name;
  }
  return name;
}

/** Number formats in first-use order; custom codes get ids from 164. */
class Styles {
  readonly #formats: (string | number)[] = [];

  /** The cellXfs index for a format; 0 is the default style. */
  index(format: string | number): number {
    let position = this.#formats.indexOf(format);
    if (position < 0) position = this.#formats.push(format) - 1;
    return position + 1;
  }

  xml(): string {
    const customs = this.#formats.filter(
      (format): format is string => typeof format === "string",
    );
    const id = (format: string | number): number =>
      typeof format === "number" ? format : 164 + customs.indexOf(format);
    const numFmts =
      customs.length === 0
        ? ""
        : `<numFmts count="${customs.length}">${customs
            .map(
              (code) =>
                `<numFmt numFmtId="${id(code)}" formatCode="${escapeXml(code)}"/>`,
            )
            .join("")}</numFmts>`;
    const xfs = [
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`,
      ...this.#formats.map(
        (format) =>
          `<xf numFmtId="${id(format)}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`,
      ),
    ];
    return `${HEADER}<styleSheet xmlns="${MAIN}">${numFmts}<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
  }
}

/** Shared strings in first-use order. */
class Strings {
  readonly #index = new Map<string, number>();
  #count = 0;

  index(text: string): number {
    this.#count += 1;
    let position = this.#index.get(text);
    if (position === undefined) {
      position = this.#index.size;
      this.#index.set(text, position);
    }
    return position;
  }

  get size(): number {
    return this.#index.size;
  }

  xml(): string {
    const items = [...this.#index.keys()].map(
      (text) => `<si><t xml:space="preserve">${escapeXml(text)}</t></si>`,
    );
    return `${HEADER}<sst xmlns="${MAIN}" count="${this.#count}" uniqueCount="${items.length}">${items.join("")}</sst>`;
  }
}

function cellXml(
  ref: string,
  input: FixtureValue,
  styles: Styles,
  strings: Strings,
): string {
  const cell: FixtureCell =
    input !== null && typeof input === "object" ? input : { value: input };
  const style =
    cell.format === undefined ? "" : ` s="${styles.index(cell.format)}"`;
  const formula =
    cell.formula === undefined ? "" : `<f>${escapeXml(cell.formula)}</f>`;
  const value = cell.value ?? null;
  if (cell.error !== undefined) {
    return `<c r="${ref}"${style} t="e">${formula}<v>${escapeXml(cell.error)}</v></c>`;
  }
  if (typeof value === "boolean") {
    return `<c r="${ref}"${style} t="b">${formula}<v>${value ? 1 : 0}</v></c>`;
  }
  if (typeof value === "number") {
    return `<c r="${ref}"${style}>${formula}<v>${value}</v></c>`;
  }
  if (typeof value === "string") {
    // A formula caches text inline; a plain string goes to the shared table.
    return formula === ""
      ? `<c r="${ref}"${style} t="s"><v>${strings.index(value)}</v></c>`
      : `<c r="${ref}"${style} t="str">${formula}<v>${escapeXml(value)}</v></c>`;
  }
  return formula === "" && style === ""
    ? ""
    : `<c r="${ref}"${style}>${formula}</c>`;
}

function worksheetXml(
  sheet: FixtureSheet,
  styles: Styles,
  strings: Strings,
): string {
  let lastRow = -1;
  let lastColumn = -1;
  const rows: string[] = [];
  sheet.rows.forEach((row, rowIndex) => {
    const cells: string[] = [];
    row.forEach((value, columnIndex) => {
      const xml = cellXml(
        `${columnName(columnIndex)}${rowIndex + 1}`,
        value,
        styles,
        strings,
      );
      if (xml === "") return;
      cells.push(xml);
      lastRow = Math.max(lastRow, rowIndex);
      lastColumn = Math.max(lastColumn, columnIndex);
    });
    if (cells.length > 0) {
      rows.push(`<row r="${rowIndex + 1}">${cells.join("")}</row>`);
    }
  });
  const dimension =
    lastRow < 0 ? "A1" : `A1:${columnName(lastColumn)}${lastRow + 1}`;
  const merges =
    sheet.merges === undefined || sheet.merges.length === 0
      ? ""
      : `<mergeCells count="${sheet.merges.length}">${sheet.merges
          .map((ref) => `<mergeCell ref="${ref}"/>`)
          .join("")}</mergeCells>`;
  const cols =
    sheet.widths === undefined || sheet.widths.length === 0
      ? ""
      : `<cols>${sheet.widths
          .map(
            (width, index) =>
              `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`,
          )
          .join("")}</cols>`;
  return `${HEADER}<worksheet xmlns="${MAIN}" xmlns:r="${REL}"><dimension ref="${dimension}"/>${cols}<sheetData>${rows.join("")}</sheetData>${merges}</worksheet>`;
}

/** Write the workbook a spec describes, as .xlsx bytes. */
export async function buildWorkbookFixture(
  spec: FixtureWorkbook,
): Promise<Uint8Array> {
  const styles = new Styles();
  const strings = new Strings();
  const worksheets = spec.sheets.map((sheet) =>
    worksheetXml(sheet, styles, strings),
  );
  const zip = new JSZip();
  const date = new Date(Date.UTC(2000, 0, 1));
  const add = (name: string, text: string): void => {
    zip.file(name, text, { date, createFolders: false });
  };
  const hasStrings = strings.size > 0;
  add(
    "[Content_Types].xml",
    `${HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${CT}.sheet.main+xml"/>${spec.sheets
      .map(
        (_, index) =>
          `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="${CT}.worksheet+xml"/>`,
      )
      .join(
        "",
      )}<Override PartName="/xl/styles.xml" ContentType="${CT}.styles+xml"/>${hasStrings ? `<Override PartName="/xl/sharedStrings.xml" ContentType="${CT}.sharedStrings+xml"/>` : ""}</Types>`,
  );
  add(
    "_rels/.rels",
    `${HEADER}<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  const names =
    spec.names === undefined || spec.names.length === 0
      ? ""
      : `<definedNames>${spec.names
          .map(
            (name) =>
              `<definedName name="${escapeXml(name.name)}"${name.localSheet === undefined ? "" : ` localSheetId="${name.localSheet}"`}${name.hidden === true ? ` hidden="1"` : ""}>${escapeXml(name.reference)}</definedName>`,
          )
          .join("")}</definedNames>`;
  add(
    "xl/workbook.xml",
    `${HEADER}<workbook xmlns="${MAIN}" xmlns:r="${REL}">${spec.date1904 === true ? `<workbookPr date1904="1"/>` : "<workbookPr/>"}<sheets>${spec.sheets
      .map(
        (sheet, index) =>
          `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}"${sheet.state === undefined ? "" : ` state="${sheet.state}"`} r:id="rId${index + 1}"/>`,
      )
      .join("")}</sheets>${names}</workbook>`,
  );
  const count = spec.sheets.length;
  add(
    "xl/_rels/workbook.xml.rels",
    `${HEADER}<Relationships xmlns="${PKG_REL}">${spec.sheets
      .map(
        (_, index) =>
          `<Relationship Id="rId${index + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`,
      )
      .join(
        "",
      )}<Relationship Id="rId${count + 1}" Type="${REL}/styles" Target="styles.xml"/>${hasStrings ? `<Relationship Id="rId${count + 2}" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>` : ""}</Relationships>`,
  );
  worksheets.forEach((xml, index) => {
    add(`xl/worksheets/sheet${index + 1}.xml`, xml);
  });
  add("xl/styles.xml", styles.xml());
  if (hasStrings) add("xl/sharedStrings.xml", strings.xml());
  return zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
  });
}

/** Shorthand for a one-sheet workbook. */
export function buildSheetFixture(
  name: string,
  rows: FixtureSheet["rows"],
): Promise<Uint8Array> {
  return buildWorkbookFixture({ sheets: [{ name, rows }] });
}
