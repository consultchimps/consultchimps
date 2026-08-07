/**
 * One style table for the merged workbook.
 *
 * A cell does not carry its formatting; it carries an index into `cellXfs`,
 * and that entry carries indexes into `fonts`, `fills`, `borders` and the
 * number formats. Transplanting a worksheet without remapping those indexes is
 * how a merged workbook ends up with the right numbers wearing another
 * workbook's clothes - or, in the engine this replaces, with no clothes at all.
 *
 * The merge copies only the entries the transplanted worksheets actually
 * reference, resolves each one down to the font/fill/border/number-format it
 * names, and deduplicates on the resolved text. Two inputs produced by the
 * same template therefore contribute one set of styles, not two.
 */
import {
  addAttribute,
  decodeXmlText,
  findElement,
  findElements,
  getAttribute,
  setAttribute,
} from "../model/xml.js";
import { escapeXmlAttribute } from "./ooxml.js";

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NAMESPACE_MAIN =
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** The lowest id a workbook may use for a number format it defines itself. */
const FIRST_CUSTOM_NUMBER_FORMAT_ID = 164;

/**
 * `styleSheet` children in schema order. A section that has to be created goes
 * in before the first section that already exists and sorts after it.
 */
const SECTION_ORDER = [
  "numFmts",
  "fonts",
  "fills",
  "borders",
  "cellStyleXfs",
  "cellXfs",
  "cellStyles",
  "dxfs",
  "tableStyles",
  "colors",
  "extLst",
] as const;

/** A minimal but complete styles part, for an input that ships without one. */
export const MINIMAL_STYLES_XML: string =
  `${XML_DECLARATION}<styleSheet xmlns="${NAMESPACE_MAIN}">` +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  "</styleSheet>";

/** The source text of every `child` inside the `container` section. */
function sectionItems(xml: string, container: string, child: string): string[] {
  const element = findElement(xml, container);
  if (!element || element.selfClosing) {
    return [];
  }
  const inner = xml.slice(element.innerStart, element.innerEnd);
  return findElements(inner, child).map((item) =>
    inner.slice(item.start, item.end),
  );
}

/** Set an attribute, adding it when the element does not declare one. */
function withAttribute(tag: string, name: string, value: string): string {
  return getAttribute(tag, name) === undefined
    ? addAttribute(tag, name, value)
    : setAttribute(tag, name, value);
}

/** Replace an element's opening tag, keeping its body and closing tag. */
function withOpenTag(elementXml: string, openTag: string): string {
  const original = /^<[^>]*>/u.exec(elementXml)?.[0].length ?? 0;
  return `${openTag}${elementXml.slice(original)}`;
}

/** One input's style part, read far enough to copy entries out of it. */
export class SourceStyles {
  readonly fonts: readonly string[];
  readonly fills: readonly string[];
  readonly borders: readonly string[];
  readonly cellStyleXfs: readonly string[];
  readonly cellXfs: readonly string[];
  readonly dxfs: readonly string[];
  /** Custom number-format codes by the id the source gave them. */
  readonly numberFormats: ReadonlyMap<number, string>;
  /** `cellStyle` elements by the `cellStyleXfs` entry they name. */
  readonly cellStyles: ReadonlyMap<number, string>;

  constructor(xml: string | undefined) {
    const source = xml ?? "";
    this.fonts = sectionItems(source, "fonts", "font");
    this.fills = sectionItems(source, "fills", "fill");
    this.borders = sectionItems(source, "borders", "border");
    this.cellStyleXfs = sectionItems(source, "cellStyleXfs", "xf");
    this.cellXfs = sectionItems(source, "cellXfs", "xf");
    this.dxfs = sectionItems(source, "dxfs", "dxf");

    const numberFormats = new Map<number, string>();
    for (const element of sectionItems(source, "numFmts", "numFmt")) {
      const id = Number(getAttribute(element, "numFmtId"));
      const code = getAttribute(element, "formatCode");
      if (Number.isInteger(id) && code !== undefined) {
        numberFormats.set(id, decodeXmlText(code));
      }
    }
    this.numberFormats = numberFormats;

    const cellStyles = new Map<number, string>();
    for (const element of sectionItems(source, "cellStyles", "cellStyle")) {
      const xfId = Number(getAttribute(element, "xfId"));
      if (Number.isInteger(xfId)) {
        cellStyles.set(xfId, element);
      }
    }
    this.cellStyles = cellStyles;
  }
}

/** A list that appends only entries it does not already hold. */
class DedupedList {
  readonly items: string[];
  readonly #indexByItem = new Map<string, number>();

  constructor(items: readonly string[]) {
    this.items = [...items];
    this.items.forEach((item, index) => {
      if (!this.#indexByItem.has(item)) {
        this.#indexByItem.set(item, index);
      }
    });
  }

  /** The index of `item`, appending it when it is new. */
  intern(item: string): number {
    const existing = this.#indexByItem.get(item);
    if (existing !== undefined) {
      return existing;
    }
    const index = this.items.length;
    this.items.push(item);
    this.#indexByItem.set(item, index);
    return index;
  }
}

/**
 * The merged style table. Seeded from the first input's styles part, then
 * grown by the entries later inputs' worksheets reference.
 */
export class MergedStyles {
  readonly #xml: string;
  readonly #fonts: DedupedList;
  readonly #fills: DedupedList;
  readonly #borders: DedupedList;
  readonly #cellStyleXfs: DedupedList;
  readonly #cellXfs: DedupedList;
  readonly #dxfs: DedupedList;
  readonly #numberFormatIdByCode = new Map<string, number>();
  readonly #numberFormatCodeById = new Map<number, string>();
  readonly #cellStyles: string[];
  readonly #cellStyleNames = new Set<string>();
  #nextNumberFormatId = FIRST_CUSTOM_NUMBER_FORMAT_ID;
  #changed = false;

  constructor(stylesXml: string | undefined) {
    const xml = stylesXml ?? MINIMAL_STYLES_XML;
    this.#xml = xml;
    const seed = new SourceStyles(xml);
    this.#fonts = new DedupedList(seed.fonts);
    this.#fills = new DedupedList(seed.fills);
    this.#borders = new DedupedList(seed.borders);
    this.#cellStyleXfs = new DedupedList(seed.cellStyleXfs);
    this.#cellXfs = new DedupedList(seed.cellXfs);
    this.#dxfs = new DedupedList(seed.dxfs);
    this.#cellStyles = sectionItems(xml, "cellStyles", "cellStyle");

    for (const [id, code] of seed.numberFormats) {
      this.#numberFormatIdByCode.set(code, id);
      this.#numberFormatCodeById.set(id, code);
      this.#nextNumberFormatId = Math.max(this.#nextNumberFormatId, id + 1);
    }
    for (const element of this.#cellStyles) {
      const name = getAttribute(element, "name");
      if (name !== undefined) {
        this.#cellStyleNames.add(decodeXmlText(name).toLocaleLowerCase());
      }
    }
  }

  /** Whether the seeded part has to be rewritten. */
  get changed(): boolean {
    return this.#changed;
  }

  /**
   * The merged-table index for one source `cellXfs` entry.
   *
   * Index 0 always maps to index 0: an unstyled cell carries no `s`
   * attribute at all, so entry 0 is the workbook's default format by
   * construction and cannot be relocated without rewriting every cell that
   * omits the attribute. Every producer writes the same default there.
   */
  remapCellXf(source: SourceStyles, index: number): number {
    if (index === 0) {
      return 0;
    }
    const element = source.cellXfs[index];
    if (element === undefined) {
      return 0;
    }
    return this.#internXf(source, element, this.#cellXfs, true);
  }

  /** The merged-table index for one source `dxfs` entry. */
  remapDxf(source: SourceStyles, index: number): number {
    const element = source.dxfs[index];
    if (element === undefined) {
      return index;
    }
    return this.#intern(this.#dxfs, element);
  }

  #intern(list: DedupedList, element: string): number {
    const before = list.items.length;
    const index = list.intern(element);
    if (list.items.length !== before) {
      this.#changed = true;
    }
    return index;
  }

  /**
   * Copy one `xf` entry, resolving the entries it points at first so the copy
   * describes the same formatting inside the merged table.
   */
  #internXf(
    source: SourceStyles,
    element: string,
    list: DedupedList,
    isCellXf: boolean,
  ): number {
    const openTag = /^<[^>]*>/u.exec(element)?.[0] ?? element;
    let rewritten = openTag;

    const numberFormatId = Number(getAttribute(openTag, "numFmtId") ?? "0");
    if (Number.isInteger(numberFormatId) && numberFormatId !== 0) {
      rewritten = withAttribute(
        rewritten,
        "numFmtId",
        String(this.#remapNumberFormat(source, numberFormatId)),
      );
    }
    rewritten = this.#remapIndexAttribute(
      rewritten,
      "fontId",
      source.fonts,
      this.#fonts,
    );
    rewritten = this.#remapIndexAttribute(
      rewritten,
      "fillId",
      source.fills,
      this.#fills,
    );
    rewritten = this.#remapIndexAttribute(
      rewritten,
      "borderId",
      source.borders,
      this.#borders,
    );

    if (isCellXf) {
      const xfId = Number(getAttribute(openTag, "xfId") ?? "0");
      if (Number.isInteger(xfId) && xfId !== 0) {
        rewritten = withAttribute(
          rewritten,
          "xfId",
          String(this.#remapCellStyleXf(source, xfId)),
        );
      }
    }

    return this.#intern(list, withOpenTag(element, rewritten));
  }

  #remapIndexAttribute(
    openTag: string,
    attribute: string,
    sourceItems: readonly string[],
    target: DedupedList,
  ): string {
    const declared = getAttribute(openTag, attribute);
    if (declared === undefined) {
      return openTag;
    }
    const index = Number(declared);
    const item = sourceItems[index];
    if (!Number.isInteger(index) || item === undefined) {
      return openTag;
    }
    return setAttribute(openTag, attribute, String(this.#intern(target, item)));
  }

  /**
   * A built-in number format keeps its id; a custom one is looked up by its
   * format code, so two inputs that both define `#,##0.00" kg"` share one id.
   */
  #remapNumberFormat(source: SourceStyles, id: number): number {
    if (id < FIRST_CUSTOM_NUMBER_FORMAT_ID) {
      return id;
    }
    const code = source.numberFormats.get(id);
    if (code === undefined) {
      return id;
    }
    const existing = this.#numberFormatIdByCode.get(code);
    if (existing !== undefined) {
      return existing;
    }
    const allocated = this.#nextNumberFormatId;
    this.#nextNumberFormatId += 1;
    this.#numberFormatIdByCode.set(code, allocated);
    this.#numberFormatCodeById.set(allocated, code);
    this.#changed = true;
    return allocated;
  }

  /** Copy the named cell style an `xf` points at, with its `cellStyle` entry. */
  #remapCellStyleXf(source: SourceStyles, index: number): number {
    const element = source.cellStyleXfs[index];
    if (element === undefined) {
      return 0;
    }
    const target = this.#internXf(source, element, this.#cellStyleXfs, false);
    const declaration = source.cellStyles.get(index);
    if (declaration !== undefined) {
      const name = decodeXmlText(
        getAttribute(declaration, "name") ?? "",
      ).toLocaleLowerCase();
      if (name !== "" && !this.#cellStyleNames.has(name)) {
        this.#cellStyles.push(
          setAttribute(declaration, "xfId", String(target)),
        );
        this.#cellStyleNames.add(name);
        this.#changed = true;
      }
    }
    return target;
  }

  /** Serialize the merged style table, preserving sections it does not own. */
  toXml(): string {
    let xml = this.#xml;
    const numberFormats = [...this.#numberFormatCodeById.entries()]
      .sort((left, right) => left[0] - right[0])
      .map(
        ([id, code]) =>
          `<numFmt numFmtId="${id}" formatCode="${escapeXmlAttribute(code)}"/>`,
      );

    xml = writeSection(xml, "numFmts", numberFormats);
    xml = writeSection(xml, "fonts", this.#fonts.items);
    xml = writeSection(xml, "fills", this.#fills.items);
    xml = writeSection(xml, "borders", this.#borders.items);
    xml = writeSection(xml, "cellStyleXfs", this.#cellStyleXfs.items);
    xml = writeSection(xml, "cellXfs", this.#cellXfs.items);
    xml = writeSection(xml, "cellStyles", this.#cellStyles);
    xml = writeSection(xml, "dxfs", this.#dxfs.items);
    return xml;
  }
}

/**
 * Replace one `styleSheet` section with the merged entries, creating it in
 * schema order when the seeded part had none.
 */
function writeSection(
  xml: string,
  container: (typeof SECTION_ORDER)[number],
  items: readonly string[],
): string {
  const replacement =
    items.length === 0
      ? ""
      : `<${container} count="${items.length}">${items.join("")}</${container}>`;
  const existing = findElement(xml, container);

  if (existing) {
    return `${xml.slice(0, existing.start)}${replacement}${xml.slice(existing.end)}`;
  }
  if (replacement === "") {
    return xml;
  }

  const position = SECTION_ORDER.indexOf(container);
  for (const later of SECTION_ORDER.slice(position + 1)) {
    const element = findElement(xml, later);
    if (element) {
      return `${xml.slice(0, element.start)}${replacement}${xml.slice(element.start)}`;
    }
  }
  const closing = xml.lastIndexOf("</styleSheet>");
  return closing < 0
    ? xml
    : `${xml.slice(0, closing)}${replacement}${xml.slice(closing)}`;
}
