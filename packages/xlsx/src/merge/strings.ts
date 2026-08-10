/**
 * One shared string table for the merged workbook.
 *
 * Two strategies were available for the strings a transplanted worksheet
 * points at: convert every `t="s"` cell to an inline string, or merge the
 * source tables and remap the indexes. This engine merges and remaps, applied
 * uniformly to every input, because:
 *
 * - remapping edits one number inside cells the transplant already rewrites
 *   for their style index, while inlining rewrites every string cell's body;
 * - a shared table keeps rich-text runs interned, so a workbook whose strings
 *   repeat (every workbook worth merging) does not grow by the product of its
 *   rows and its columns; and
 * - identical `<si>` items across inputs collapse, which is what makes two
 *   copies of the same workbook merge into a table the size of one.
 *
 * Items are compared by their exact source XML, so two identical plain strings
 * dedupe and a plain string never collapses into a rich-text item that happens
 * to render the same.
 */
import { findElements } from "../model/xml.js";

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NAMESPACE_MAIN =
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** The `<si>` items of one shared string table, as their source XML. */
export function parseSharedStringItems(xml: string | undefined): string[] {
  if (xml === undefined) {
    return [];
  }
  return findElements(xml, "si").map((element) =>
    xml.slice(element.start, element.end),
  );
}

/** The merged table, plus the per-input index remaps that feed a transplant. */
export class SharedStringTable {
  readonly #items: string[];
  readonly #indexByItem = new Map<string, number>();
  #changed = false;

  private constructor(items: readonly string[]) {
    this.#items = [...items];
    this.#items.forEach((item, index) => {
      if (!this.#indexByItem.has(item)) {
        this.#indexByItem.set(item, index);
      }
    });
  }

  /** Seed the merged table from the first input's table, if it has one. */
  static from(xml: string | undefined): SharedStringTable {
    return new SharedStringTable(parseSharedStringItems(xml));
  }

  /** Whether an item was appended since the table was seeded. */
  get changed(): boolean {
    return this.#changed;
  }

  get size(): number {
    return this.#items.length;
  }

  /**
   * Fold one source table into the merged one. The returned array maps a
   * source index to its index in the merged table; it is empty when the source
   * declared no table, in which case no cell in that workbook can point at one.
   */
  absorb(xml: string | undefined): readonly number[] {
    return parseSharedStringItems(xml).map((item) => {
      const existing = this.#indexByItem.get(item);
      if (existing !== undefined) {
        return existing;
      }
      const index = this.#items.length;
      this.#items.push(item);
      this.#indexByItem.set(item, index);
      this.#changed = true;
      return index;
    });
  }

  /**
   * Serialize the merged table. `count` counts references and `uniqueCount`
   * counts items; Excel treats both as hints and recomputes them, so the item
   * count answers for both rather than tracking every cell that points here.
   */
  toXml(): string {
    return `${XML_DECLARATION}<sst xmlns="${NAMESPACE_MAIN}" count="${this.#items.length}" uniqueCount="${this.#items.length}">${this.#items.join(
      "",
    )}</sst>`;
  }
}
