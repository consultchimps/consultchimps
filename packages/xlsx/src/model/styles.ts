/**
 * L1 - just enough of `xl/styles.xml` to answer two questions: is this cell
 * formatted as a date, and what is its number format?
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

/**
 * The codes of the built-in number formats a workbook may point at without
 * declaring, as Excel shows them in an English (United States) locale. Ids
 * 5 to 8, 23 to 36 and 41 to 44 are locale-dependent and the file format
 * leaves them unwritten; these are the defaults Excel applies, as SheetJS did. Ids
 * 50 to 81 are East Asian built-ins, read as their Western counterparts.
 */
const BUILTIN_FORMAT_CODES: Readonly<Record<number, string>> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  5: '"$"#,##0_);\\("$"#,##0\\)',
  6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
  7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
  8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  23: "General",
  24: "General",
  25: "General",
  26: "General",
  27: "m/d/yy",
  28: "m/d/yy",
  29: "m/d/yy",
  30: "m/d/yy",
  31: "m/d/yy",
  32: "h:mm:ss",
  33: "h:mm:ss",
  34: "h:mm:ss",
  35: "h:mm:ss",
  36: "m/d/yy",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
  50: "m/d/yy",
  51: "m/d/yy",
  52: "m/d/yy",
  53: "m/d/yy",
  54: "m/d/yy",
  55: "m/d/yy",
  56: "m/d/yy",
  57: "m/d/yy",
  58: "m/d/yy",
  59: "0",
  60: "0.00",
  61: "#,##0",
  62: "#,##0.00",
  63: '"$"#,##0_);\\("$"#,##0\\)',
  64: '"$"#,##0_);[Red]\\("$"#,##0\\)',
  65: '"$"#,##0.00_);\\("$"#,##0.00\\)',
  66: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  67: "0%",
  68: "0.00%",
  69: "# ?/?",
  70: "# ??/??",
  71: "m/d/yy",
  72: "m/d/yy",
  73: "d-mmm-yy",
  74: "d-mmm",
  75: "mmm-yy",
  76: "h:mm",
  77: "h:mm:ss",
  78: "m/d/yy h:mm",
  79: "mm:ss",
  80: "[h]:mm:ss",
  81: "mmss.0",
};

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

  /**
   * The number format code of the cell style at `styleIndex`: the workbook's
   * own code for the id, else the built-in one, else General.
   */
  formatCode(styleIndex: number | undefined): string {
    const formatId = this.#formatIdByStyle[styleIndex ?? 0];
    if (
      formatId === undefined ||
      !Number.isInteger(formatId) ||
      formatId === 0
    ) {
      return "General";
    }
    return (
      this.#customFormatCodes.get(formatId) ??
      BUILTIN_FORMAT_CODES[formatId] ??
      "General"
    );
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
