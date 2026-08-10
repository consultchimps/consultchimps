/**
 * L1 - just enough of `xl/styles.xml` to answer one question: is this cell
 * formatted as a date?
 *
 * Excel stores dates as numbers; only the number format distinguishes
 * 45000 the quantity from 45000 the day. Grouping keys and output filenames
 * depend on telling them apart, so the model reads the format id a cell's
 * style points at. It reads nothing else: styles.xml is never rewritten, which
 * is what lets it travel through an edit byte-identical.
 */
import { decodeXmlText, editElements, getAttribute } from "./xml.js";

/**
 * The built-in number formats ECMA-376 defines as dates or times. Ids outside
 * these ranges are either non-date built-ins or custom formats declared in
 * `<numFmts>`.
 */
const BUILTIN_DATE_FORMAT_IDS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47,
]);

/** Milliseconds in one day, the unit an Excel serial counts. */
const MILLISECONDS_PER_DAY = 86_400_000;

/**
 * Whether a custom format code describes a date or a time. Quoted literals and
 * bracketed sections are removed first so `0.00" months"` and `[$-409]` do not
 * read as date tokens.
 */
export function isDateFormatCode(formatCode: string): boolean {
  const bare = formatCode
    .replace(/"[^"]*"/gu, "")
    .replace(/\[[^\]]*\]/gu, "")
    .replace(/\\./gu, "");
  return /[ymdhs]/iu.test(bare);
}

/** The number formats a workbook's cell styles point at. */
export class StyleTable {
  /** Number-format id per `cellXfs` entry, indexed by a cell's `s` attribute. */
  readonly #formatIdByStyle: readonly number[];
  /** Format codes of the workbook's custom (`numFmtId >= 164`) formats. */
  readonly #customFormatCodes: ReadonlyMap<number, string>;

  private constructor(
    formatIdByStyle: readonly number[],
    customFormatCodes: ReadonlyMap<number, string>,
  ) {
    this.#formatIdByStyle = formatIdByStyle;
    this.#customFormatCodes = customFormatCodes;
  }

  static parse(stylesXml: string | undefined): StyleTable {
    if (stylesXml === undefined) {
      return new StyleTable([], new Map());
    }

    const customFormatCodes = new Map<number, string>();
    editElements(stylesXml, "numFmt", (element, text) => {
      const id = Number(getAttribute(element.openTag, "numFmtId"));
      const code = getAttribute(element.openTag, "formatCode");
      if (Number.isInteger(id) && code !== undefined) {
        // A format code's own quotes arrive escaped inside the attribute.
        customFormatCodes.set(id, decodeXmlText(code));
      }
      return text;
    });

    // Only `cellXfs` is indexed by a cell's `s`; `cellStyleXfs` is not.
    const formatIdByStyle: number[] = [];
    editElements(stylesXml, "cellXfs", (container, containerText) => {
      const inner = containerText.slice(
        container.innerStart - container.start,
        container.innerEnd - container.start,
      );
      editElements(inner, "xf", (element, text) => {
        formatIdByStyle.push(Number(getAttribute(element.openTag, "numFmtId")));
        return text;
      });
      return containerText;
    });

    return new StyleTable(formatIdByStyle, customFormatCodes);
  }

  /** Whether the cell style at `styleIndex` formats its value as a date. */
  isDateStyle(styleIndex: number | undefined): boolean {
    const formatId = this.#formatIdByStyle[styleIndex ?? 0];
    if (formatId === undefined || !Number.isInteger(formatId)) {
      return false;
    }
    if (BUILTIN_DATE_FORMAT_IDS.has(formatId)) {
      return true;
    }
    const custom = this.#customFormatCodes.get(formatId);
    return custom === undefined ? false : isDateFormatCode(custom);
  }
}

/**
 * Convert an Excel date serial to a `Date`.
 *
 * The components are assembled in local time, matching what the spreadsheet
 * reader the split engine used to call handed back. Group keys derive from
 * `toISOString()`, so changing the construction here would silently rename
 * every date-valued output workbook.
 */
export function excelSerialToDate(serial: number, date1904: boolean): Date {
  const epoch = date1904
    ? Date.UTC(1904, 0, 1)
    : // The 1900 system counts a day that never existed (29 February 1900),
      // which an epoch of 30 December 1899 absorbs for every later serial.
      Date.UTC(1899, 11, 30);
  const instant = new Date(epoch + Math.round(serial * MILLISECONDS_PER_DAY));

  return new Date(
    instant.getUTCFullYear(),
    instant.getUTCMonth(),
    instant.getUTCDate(),
    instant.getUTCHours(),
    instant.getUTCMinutes(),
    instant.getUTCSeconds(),
    instant.getUTCMilliseconds(),
  );
}
