import JSZip from "jszip";
import type { PbiReasonCode } from "./errors.js";
import type { PbiColumnType, PbiValue } from "./model.js";
import { toCell } from "./values.js";

/**
 * A byte-deterministic workbook writer. The same input bytes and options
 * produce the same workbook bytes on every platform and every run: fixed ZIP
 * entry dates, one compression method and level, a fixed entry order, no
 * `streamFiles`, `platform: "UNIX"`, and no locale-sensitive comparison.
 *
 * It lives here rather than in `@consultchimps/xlsx` because everything ADR
 * Decision 5 and Decision 6 need is net-new: the SpreadsheetML escaping, the
 * sanitizing name allocator, typed cells, number formats and the multi-sheet
 * split. It is internal and is not part of this package's public surface.
 */

const FIXED_DATE = new Date(Date.UTC(1980, 0, 1, 0, 0, 0));
const LITERAL_ESCAPE = /^_x[0-9a-fA-F]{4}_/;

function encodeUnit(unit: number): string {
  return `_x${unit.toString(16).toUpperCase().padStart(4, "0")}_`;
}

/**
 * SpreadsheetML string escaping, lossless and adding no manifest reason. Every
 * XML 1.0-forbidden code unit, every carriage return, U+FFFE, U+FFFF and every
 * unpaired surrogate becomes `_xHHHH_`; a literal `_xHHHH_` has its leading
 * underscore encoded so it is recovered exactly. Ordinary entity escaping is
 * applied on top, and leading and trailing whitespace is preserved.
 */
export function escapeSpreadsheetText(
  value: string,
  attribute = false,
): string {
  let out = "";
  for (let index = 0; index < value.length; index++) {
    if (
      value.charCodeAt(index) === 0x5f &&
      LITERAL_ESCAPE.test(value.slice(index, index + 7))
    ) {
      // Protect the original sequence, then continue from the character after
      // the underscore so nothing this pass produced is re-escaped.
      out += encodeUnit(0x5f);
      continue;
    }
    const unit = value.charCodeAt(index);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const low = value.charCodeAt(index + 1);
      if (low >= 0xdc00 && low <= 0xdfff) {
        out += value.slice(index, index + 2);
        index++;
        continue;
      }
      out += encodeUnit(unit);
      continue;
    }
    if (unit >= 0xdc00 && unit <= 0xdfff) {
      out += encodeUnit(unit);
      continue;
    }
    if (
      unit <= 0x08 ||
      unit === 0x0b ||
      unit === 0x0c ||
      unit === 0x0d ||
      (unit >= 0x0e && unit <= 0x1f) ||
      unit === 0xfffe ||
      unit === 0xffff
    ) {
      out += encodeUnit(unit);
      continue;
    }
    if (unit === 0x26) out += "&amp;";
    else if (unit === 0x3c) out += "&lt;";
    else if (unit === 0x3e) out += "&gt;";
    else if (attribute && unit === 0x22) out += "&quot;";
    else out += value[index];
  }
  return out;
}

export function columnReference(index: number): string {
  let out = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26))
    out = String.fromCharCode(65 + ((n - 1) % 26)) + out;
  return out;
}

export interface WorksheetColumn {
  /** Manifest key, `tableId:columnId`, used to attribute conversion counts. */
  readonly key: string;
  readonly header: string;
  readonly type: PbiColumnType;
  readonly values: readonly PbiValue[];
}

export interface WorksheetPlan {
  readonly sheetName: string;
  readonly columns: readonly WorksheetColumn[];
  /** Zero-based half-open source row range this part carries. */
  readonly start: number;
  readonly end: number;
}

export type ConversionCounts = Map<string, Map<PbiReasonCode, number>>;

function tally(
  counts: ConversionCounts,
  key: string,
  reason: PbiReasonCode,
): void {
  let bucket = counts.get(key);
  if (bucket === undefined) {
    bucket = new Map<PbiReasonCode, number>();
    counts.set(key, bucket);
  }
  bucket.set(reason, (bucket.get(reason) ?? 0) + 1);
}

function worksheetXml(plan: WorksheetPlan, counts: ConversionCounts): string {
  const parts: string[] = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>',
    '<row r="1">',
  ];
  const references = plan.columns.map((_, index) => columnReference(index));
  for (let index = 0; index < plan.columns.length; index++) {
    parts.push(
      `<c r="${references[index]!}1" t="inlineStr"><is><t xml:space="preserve">${escapeSpreadsheetText(plan.columns[index]!.header)}</t></is></c>`,
    );
  }
  parts.push("</row>");
  for (let row = plan.start; row < plan.end; row++) {
    const number = row - plan.start + 2;
    let line = `<row r="${number}">`;
    for (let index = 0; index < plan.columns.length; index++) {
      const column = plan.columns[index]!;
      const { cell, reason } = toCell(column.values[row] ?? null, column.type);
      if (reason !== undefined) tally(counts, column.key, reason);
      const reference = references[index]! + number;
      if (cell.kind === "blank") continue;
      if (cell.kind === "number") {
        line +=
          cell.style === 0
            ? `<c r="${reference}"><v>${cell.text}</v></c>`
            : `<c r="${reference}" s="${cell.style}"><v>${cell.text}</v></c>`;
      } else if (cell.kind === "boolean") {
        line += `<c r="${reference}" t="b"><v>${cell.value ? 1 : 0}</v></c>`;
      } else {
        line += `<c r="${reference}" t="inlineStr"><is><t xml:space="preserve">${escapeSpreadsheetText(cell.text)}</t></is></c>`;
      }
    }
    parts.push(`${line}</row>`);
  }
  parts.push("</sheetData></worksheet>");
  return parts.join("");
}

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd\\ hh:mm:ss"/><numFmt numFmtId="165" formatCode="#,##0.0000"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

export interface WorkbookResult {
  readonly bytes: Uint8Array;
  readonly counts: ConversionCounts;
}

export async function writeWorkbook(
  plans: readonly WorksheetPlan[],
): Promise<WorkbookResult> {
  const counts: ConversionCounts = new Map();
  const zip = new JSZip();
  const add = (path: string, content: string): void => {
    zip.file(path, content, {
      date: FIXED_DATE,
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
      createFolders: false,
    });
  };
  add(
    "[Content_Types].xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      plans
        .map(
          (_, index) =>
            `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
        )
        .join("") +
      "</Types>",
  );
  add(
    "_rels/.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      "</Relationships>",
  );
  add(
    "xl/workbook.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
      plans
        .map(
          (plan, index) =>
            `<sheet name="${escapeSpreadsheetText(plan.sheetName, true)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`,
        )
        .join("") +
      "</sheets></workbook>",
  );
  add(
    "xl/_rels/workbook.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      plans
        .map(
          (_, index) =>
            `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`,
        )
        .join("") +
      `<Relationship Id="rId${plans.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      "</Relationships>",
  );
  add("xl/styles.xml", STYLES);
  for (let index = 0; index < plans.length; index++)
    add(
      `xl/worksheets/sheet${index + 1}.xml`,
      worksheetXml(plans[index]!, counts),
    );
  const bytes = await zip.generateAsync({
    type: "uint8array",
    platform: "UNIX",
    streamFiles: false,
  });
  return { bytes, counts };
}
