import type { CellValue } from "@consultchimps/tabular";
import { Zip, ZipDeflate, strToU8 } from "fflate";

import { CellDate } from "./cell-date.js";
import { CellError } from "./cell-error.js";

/** A value the writer can store: a table value, an error cell, or a date. */
export type WritableCellValue = CellValue | CellError | CellDate;

/**
 * A single-worksheet workbook written row by row (ADR 0006). The worksheet XML
 * is deflated as it is produced, so memory holds one batch of rows and the
 * deflate window, never the whole sheet or a cell object per value.
 *
 * Text is stored inline in each cell rather than in a shared-strings table: a
 * shared table cannot be written until every string is known, which a stream
 * does not know. Output depends only on the rows given: every zip entry carries
 * the same fixed timestamp, so identical rows give identical bytes on any
 * machine, in any time zone, in Node and in the browser.
 */

// fflate writes an entry's DOS timestamp from a date's local fields in the time
// zone current when it writes, so the date is built at that moment from local
// fields: 1980-01-01 12:00 then encodes identically in every zone. A date built
// once would shift with a later zone change, and midnight could fall into the
// year before, which a zip cannot store. Noon also clears daylight-saving gaps.
function fixedModificationTime(): Date {
  return new Date(1980, 0, 1, 12, 0, 0);
}
const FLUSH_CHARS = 64 * 1024;
const CORE_DATE = "1970-01-01T00:00:00Z";

// Characters XML 1.0 cannot carry, and the carriage return, which XML parsers
// normalize to a line feed on read. OOXML writes them as _xHHHH_ (ECMA-376
// ST_Xstring), and a literal _xHHHH_ in the text is itself escaped as
// _x005F_xHHHH_ so a reader does not decode it into a different character.
const INVALID_XML =
  // eslint-disable-next-line no-control-regex -- XML 1.0 cannot carry these characters; the pattern finds them so they can be escaped.
  /[\u0000-\u0008\u000B\u000C\u000D\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/gu;
const LITERAL_ESCAPE = /_(x[0-9A-Fa-f]{4}_)/gu;
// Any character one of the replacements below would touch.
const NEEDS_ESCAPE =
  // eslint-disable-next-line no-control-regex -- XML 1.0 cannot carry these characters; the pattern finds them so they can be escaped.
  /[&<>_\u0000-\u0008\u000B\u000C\u000D\u000E-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/u;

export function escapeCellText(text: string): string {
  // Most text has nothing to escape; checking once is far cheaper than five
  // replacements over every string in a large table.
  if (!NEEDS_ESCAPE.test(text)) return text;
  return text
    .replace(LITERAL_ESCAPE, "_x005F_$1")
    .replace(/&/gu, "&amp;")
    .replace(/</gu, "&lt;")
    .replace(/>/gu, "&gt;")
    .replace(
      INVALID_XML,
      (char) =>
        `_x${char.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")}_`,
    );
}

function escapeAttribute(text: string): string {
  return escapeCellText(text).replace(/"/gu, "&quot;");
}

export function columnLetters(index: number): string {
  let letters = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  }
  return letters;
}

/** The column width today's tables use: the longest value plus two, between 10 and 60 characters. */
export function tableColumnWidth(longest: number): number {
  return Math.min(Math.max(longest + 2, 10), 60);
}

/** The length a value contributes to its column's width. */
export function cellWidthLength(value: WritableCellValue): number {
  return value === null ? 0 : String(value).length;
}

export interface TableWorkbookWriterOptions {
  sheetName: string;
  columns: readonly string[];
  /** Width of each column in characters, as tableColumnWidth gives it. */
  widths: readonly number[];
  /** Data rows that will be written, not counting the header. */
  rowCount: number;
  /** Receives the workbook's bytes in order as they are produced. */
  onChunk: (chunk: Uint8Array) => void;
}

/**
 * Writes the header row at construction, then each data row as it arrives.
 * finish() must follow exactly `rowCount` calls to writeRow.
 */
export class TableWorkbookWriter {
  readonly #columns: number;
  readonly #rowCount: number;
  readonly #letters: readonly string[];
  readonly #zip: Zip;
  readonly #sheet: ZipDeflate;
  readonly #lastRef: string;
  #pending = "";
  #written = 0;
  #error: Error | undefined;

  constructor(options: TableWorkbookWriterOptions) {
    const { columns, widths, rowCount, sheetName } = options;
    if (columns.length === 0) {
      throw new Error("A table workbook needs at least one column.");
    }
    if (widths.length !== columns.length) {
      throw new Error("Every column needs a width.");
    }
    this.#columns = columns.length;
    this.#rowCount = rowCount;
    this.#letters = columns.map((_, index) => columnLetters(index));
    this.#lastRef = `${this.#letters.at(-1)}${rowCount + 1}`;

    this.#zip = new Zip((error, data) => {
      if (error) this.#error = error;
      else options.onChunk(data);
    });
    this.#addPart("[Content_Types].xml", contentTypesXml());
    this.#addPart("_rels/.rels", packageRelsXml());
    this.#addPart("docProps/core.xml", corePropertiesXml());
    this.#addPart("docProps/app.xml", appPropertiesXml(sheetName));
    this.#addPart(
      "xl/workbook.xml",
      workbookXml(sheetName, this.#letters.at(-1)!, rowCount + 1),
    );
    this.#addPart("xl/_rels/workbook.xml.rels", workbookRelsXml());
    this.#addPart("xl/styles.xml", STYLES_XML);

    this.#sheet = new ZipDeflate("xl/worksheets/sheet1.xml", { level: 6 });
    this.#sheet.mtime = fixedModificationTime();
    this.#zip.add(this.#sheet);
    this.#pending =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">` +
      `<dimension ref="A1:${this.#lastRef}"/>` +
      `<sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
      `<cols>${widths
        .map(
          (width, index) =>
            `<col min="${index + 1}" max="${index + 1}" width="${width + 0.83203125}" customWidth="1"/>`,
        )
        .join("")}</cols><sheetData>`;
    this.#appendRow(1, columns);
  }

  writeRow(values: readonly WritableCellValue[]): void {
    if (this.#written >= this.#rowCount) {
      throw new Error(
        `More rows were written than the ${this.#rowCount} declared.`,
      );
    }
    this.#written += 1;
    this.#appendRow(this.#written + 1, values);
  }

  finish(): void {
    if (this.#written !== this.#rowCount) {
      throw new Error(
        `${this.#written} rows were written but ${this.#rowCount} were declared.`,
      );
    }
    const ref = `A1:${this.#lastRef}`;
    this.#pending +=
      `</sheetData><autoFilter ref="${ref}"/>` +
      `<ignoredErrors><ignoredError numberStoredAsText="1" sqref="${ref}"/></ignoredErrors>` +
      `</worksheet>`;
    this.#sheet.push(strToU8(this.#pending), true);
    this.#pending = "";
    this.#zip.end();
    if (this.#error) throw this.#error;
  }

  #appendRow(rowNumber: number, values: readonly WritableCellValue[]): void {
    let xml = `<row r="${rowNumber}">`;
    for (let index = 0; index < this.#columns; index += 1) {
      const value = values[index] ?? null;
      if (value === null) continue;
      const ref = `${this.#letters[index]}${rowNumber}`;
      if (typeof value === "number") {
        // A number that is not finite has no cell form; it stays blank, as the
        // readers treat one.
        if (Number.isFinite(value))
          xml += `<c r="${ref}"><v>${String(value)}</v></c>`;
      } else if (value instanceof CellDate) {
        xml += `<c r="${ref}" s="${value.time ? DATE_TIME_STYLE : DATE_STYLE}"><v>${String(value.serial)}</v></c>`;
      } else if (value instanceof CellError) {
        xml += `<c r="${ref}" t="e"><v>${escapeCellText(value.text)}</v></c>`;
      } else if (typeof value === "boolean") {
        xml += `<c r="${ref}" t="b"><v>${value ? 1 : 0}</v></c>`;
      } else {
        const space = /^\s|\s$/u.test(value) ? ' xml:space="preserve"' : "";
        xml += `<c r="${ref}" t="inlineStr"><is><t${space}>${escapeCellText(value)}</t></is></c>`;
      }
    }
    this.#pending += `${xml}</row>`;
    if (this.#pending.length >= FLUSH_CHARS) {
      this.#sheet.push(strToU8(this.#pending));
      this.#pending = "";
      if (this.#error) throw this.#error;
    }
  }

  #addPart(name: string, xml: string): void {
    const part = new ZipDeflate(name, { level: 6 });
    part.mtime = fixedModificationTime();
    this.#zip.add(part);
    part.push(strToU8(xml), true);
  }
}

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const XML_HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;

function contentTypesXml(): string {
  return (
    XML_HEADER +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
    `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` +
    `</Types>`
  );
}

function packageRelsXml(): string {
  return (
    XML_HEADER +
    `<Relationships xmlns="${PKG_REL_NS}">` +
    `<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
    `<Relationship Id="rId3" Type="${REL_NS}/extended-properties" Target="docProps/app.xml"/>` +
    `</Relationships>`
  );
}

function corePropertiesXml(): string {
  return (
    XML_HEADER +
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    `<dc:creator>ConsultChimps</dc:creator><cp:lastModifiedBy>ConsultChimps</cp:lastModifiedBy>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${CORE_DATE}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${CORE_DATE}</dcterms:modified>` +
    `</cp:coreProperties>`
  );
}

function appPropertiesXml(sheetName: string): string {
  return (
    XML_HEADER +
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
    `<Application>ConsultChimps</Application>` +
    `<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>` +
    `<TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>${escapeCellText(sheetName)}</vt:lpstr></vt:vector></TitlesOfParts>` +
    `</Properties>`
  );
}

function workbookXml(
  sheetName: string,
  lastColumn: string,
  lastRow: number,
): string {
  // A defined name refers to its sheet in quotes, with any quote doubled, so a
  // name with spaces or punctuation still resolves.
  const quoted = `'${sheetName.replace(/'/gu, "''")}'`;
  const absolute = `$A$1:$${lastColumn}$${lastRow}`;
  return (
    XML_HEADER +
    `<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">` +
    `<sheets><sheet name="${escapeAttribute(sheetName)}" sheetId="1" r:id="rId1"/></sheets>` +
    `<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">${escapeCellText(`${quoted}!${absolute}`)}</definedName></definedNames>` +
    `</workbook>`
  );
}

function workbookRelsXml(): string {
  return (
    XML_HEADER +
    `<Relationships xmlns="${PKG_REL_NS}">` +
    `<Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>` +
    `<Relationship Id="rId2" Type="${REL_NS}/styles" Target="styles.xml"/>` +
    `</Relationships>`
  );
}

// A date is written as its serial with one of two formats: the date alone at
// midnight, the date and time otherwise. Both are ISO order, which reads the
// same in every locale, where the built-in date formats follow the viewer's
// regional settings.
const DATE_STYLE = 1;
const DATE_TIME_STYLE = 2;

const STYLES_XML =
  XML_HEADER +
  `<styleSheet xmlns="${MAIN_NS}">` +
  `<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts>` +
  `<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
  `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
  `<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
  `<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>` +
  `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
  `</styleSheet>`;
