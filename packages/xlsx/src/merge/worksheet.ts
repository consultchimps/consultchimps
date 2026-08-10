/**
 * Re-indexing a transplanted part.
 *
 * A worksheet part is almost self-contained: its cells, merges, conditional
 * formatting, data validation and hyperlink anchors all describe positions
 * that a transplant does not move. What it is NOT self-contained about are the
 * three index spaces it borrows from the workbook around it - the shared
 * string table, the cell formats, and the differential formats - plus the
 * relationship ids that name its dependent parts. Those are exactly what this
 * module rewrites, and nothing else, so everything the old merge dropped
 * survives by simply not being touched.
 */
import { editElements, getAttribute, setAttribute } from "../model/xml.js";

/** The index translations one transplanted worksheet needs. */
export interface IndexRemaps {
  /** Source `cellXfs` index to its index in the merged style table. */
  readonly style: (index: number) => number;
  /** Source `sst` index to its index in the merged string table. */
  readonly sharedString: (index: number) => number;
  /** Source `dxfs` index to its index in the merged style table. */
  readonly differentialFormat: (index: number) => number;
}

/** A cached-value element holding a shared-string index. */
const SHARED_STRING_VALUE =
  /(<(?:[A-Za-z_][\w.-]*:)?v>)(\d+)(<\/(?:[A-Za-z_][\w.-]*:)?v>)/u;

function remapAttribute(
  openTag: string,
  attribute: string,
  remap: (index: number) => number,
): string {
  const declared = getAttribute(openTag, attribute);
  if (declared === undefined) {
    return openTag;
  }
  const index = Number(declared);
  if (!Number.isInteger(index) || index < 0) {
    return openTag;
  }
  return setAttribute(openTag, attribute, String(remap(index)));
}

/** Rewrite the style, string and differential-format indexes of a worksheet. */
export function rewriteWorksheetIndexes(
  xml: string,
  remaps: IndexRemaps,
): string {
  let rewritten = editElements(xml, "c", (element, text) => {
    const openTag = remapAttribute(element.openTag, "s", remaps.style);
    const body = text.slice(element.openTag.length);
    if (getAttribute(openTag, "t") !== "s") {
      return `${openTag}${body}`;
    }
    return `${openTag}${body.replace(
      SHARED_STRING_VALUE,
      (_match, open: string, digits: string, close: string) =>
        `${open}${remaps.sharedString(Number(digits))}${close}`,
    )}`;
  });

  rewritten = editElements(rewritten, "row", (element, text) => {
    const openTag = remapAttribute(element.openTag, "s", remaps.style);
    return `${openTag}${text.slice(element.openTag.length)}`;
  });

  rewritten = editElements(rewritten, "col", (element, text) => {
    const openTag = remapAttribute(element.openTag, "style", remaps.style);
    return `${openTag}${text.slice(element.openTag.length)}`;
  });

  return rewriteDifferentialFormatIds(rewritten, remaps.differentialFormat);
}

/**
 * Rewrite every attribute that indexes `dxfs`. Conditional-formatting rules
 * name one; an Excel Table part names up to four (header, data, totals and per
 * column), which is why the attributes are matched by suffix rather than by a
 * list that would have to grow with the schema.
 */
export function rewriteDifferentialFormatIds(
  xml: string,
  remap: (index: number) => number,
): string {
  return xml.replace(
    /\b([A-Za-z]*[Dd]xfId)="(\d+)"/gu,
    (_match, attribute: string, index: string) =>
      `${attribute}="${remap(Number(index))}"`,
  );
}

/**
 * Point a copied part's relationship references at the ids its new
 * relationships part declares. Relationship ids only ever appear as whole
 * attribute values, so the rewrite is anchored on the `="rIdN"` shape.
 */
export function rewriteRelationshipIds(
  xml: string,
  idMap: ReadonlyMap<string, string>,
): string {
  if (idMap.size === 0) {
    return xml;
  }
  return xml.replace(
    /=(["'])(rId\d+)\1/gu,
    (match, quote: string, id: string) => {
      const mapped = idMap.get(id);
      return mapped === undefined ? match : `=${quote}${mapped}${quote}`;
    },
  );
}
