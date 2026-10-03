import { SaxesParser, type SaxesTag } from "saxes";

import type { ZipReader } from "./zip-reader.js";

/**
 * Worksheet and shared-string parts read as a stream of XML events (ADR 0006).
 * This layer reports what the markup says, cell by cell, and interprets none of
 * it: which cells hold values, what type a value has and where a row starts are
 * the reader's questions, answered in `src/operations/consolidate/`.
 */

/** One `<c>` element as written. */
export interface RawCell {
  /** The `r` attribute as written, or undefined. */
  readonly ref: string | undefined;
  /** The `t` attribute, or undefined. */
  readonly type: string | undefined;
  /** The `s` attribute, or undefined. */
  readonly style: string | undefined;
  /** The text of `<v>` with entities resolved, or undefined when absent or empty. */
  readonly value: string | undefined;
  /**
   * The text of `<is>` read as a string item (see `StringItem`), with entities
   * and `_xHHHH_` escapes resolved; null when the cell has no `<is>`.
   */
  readonly inline: string | undefined | null;
  /** The cell's `<f>` element as written, or undefined when it has none. */
  readonly formula: RawFormula | undefined;
}

/** A `<f>` element: its text and the attributes that say what it shares. */
export interface RawFormula {
  readonly text: string;
  readonly type: string | undefined;
  readonly sharedIndex: string | undefined;
  readonly ref: string | undefined;
}

export interface WorksheetEvents {
  /** The `ref` of the first `<dimension>` before the cells, as written. */
  dimension(ref: string): void;
  /**
   * One `<row>`, with its `r` attribute as written. A row that holds no cells
   * reports an empty list.
   */
  row(ref: string | undefined, selfClosing: boolean, cells: RawCell[]): void;
  /** The end of `<sheetData>`. */
  endRows(): void;
  /** A `<mergeCell>` after the cells, with its attributes. */
  mergeCell(attributes: Readonly<Record<string, string>>): void;
}

const ESCAPE = /_x([\da-fA-F]{4})_/giu;

/**
 * Resolve the `_xHHHH_` escapes ECMA-376 writes characters XML cannot carry
 * with (ST_Xstring), in one pass, so `_x005F_x0041_` reads `_x0041_`.
 */
export function decodeEscapes(text: string): string {
  return text.includes("_")
    ? text.replace(ESCAPE, (_, code: string) =>
        String.fromCharCode(Number.parseInt(code, 16)),
      )
    : text;
}

function localName(name: string): string {
  const colon = name.indexOf(":");
  return colon < 0 ? name : name.slice(colon + 1);
}

/**
 * The text of an `<si>` or `<is>` string item. A plain item is the text of its
 * first `<t>`; a rich one, whose runs are `<r>` elements, is every `<t>` joined,
 * phonetic runs (`<rPh>`) left out. An item that is neither holds no text.
 */
class StringItem {
  readonly #trimmed: boolean;
  #sawText = false;
  #first: string | undefined;
  #joined = "";
  #plain = false;
  #rich = false;
  #seenElement = false;
  #inText = false;
  #phonetic = 0;
  #textCount = 0;
  #current = "";

  /**
   * `trimmed` says whether the item's content is trimmed before it is read,
   * as a shared-string item's is and an inline one's is not: an item holding
   * only spaces is then empty text rather than no text.
   */
  constructor(trimmed: boolean) {
    this.#trimmed = trimmed;
  }

  open(name: string): void {
    if (!this.#seenElement) {
      this.#seenElement = true;
      this.#plain = name === "t";
    }
    if (name === "r") this.#rich = true;
    else if (name === "rPh") this.#phonetic += 1;
    else if (name === "t" && this.#phonetic === 0) {
      this.#inText = true;
      this.#current = "";
    }
  }

  close(name: string): void {
    if (name === "rPh") this.#phonetic -= 1;
    else if (name === "t" && this.#inText) {
      this.#inText = false;
      if (this.#textCount === 0) this.#first = this.#current;
      this.#textCount += 1;
      this.#joined += this.#current;
    }
  }

  text(text: string, escaped: boolean): void {
    if (!this.#seenElement && text !== "") this.#sawText = true;
    if (this.#inText) this.#current += escaped ? decodeEscapes(text) : text;
  }

  value(): string | undefined {
    // An item with no elements is empty text when it holds nothing at all.
    if (!this.#seenElement) {
      return this.#trimmed || !this.#sawText ? "" : undefined;
    }
    if (this.#plain) return this.#first ?? "";
    return this.#rich ? this.#joined : undefined;
  }
}

function createParser(
  part: string,
  handlers: {
    open(name: string, tag: SaxesTag): void;
    close(name: string): void;
    text(text: string, escaped: boolean): void;
  },
): SaxesParser {
  const parser = new SaxesParser({ fileName: part, position: false });
  parser.on("doctype", () => {
    throw new Error(`DOCTYPE declarations are not allowed in ${part}.`);
  });
  parser.on("error", (error) => {
    throw error;
  });
  parser.on("opentag", (tag) => {
    handlers.open(localName(tag.name), tag);
  });
  parser.on("closetag", (tag) => {
    handlers.close(localName(tag.name));
  });
  parser.on("text", (text) => {
    handlers.text(text, true);
  });
  parser.on("cdata", (text) => {
    handlers.text(text, false);
  });
  return parser;
}

async function parsePart(
  zip: ZipReader,
  part: string,
  parser: SaxesParser,
  between?: () => Promise<void>,
): Promise<void> {
  const decoder = new TextDecoder();
  await zip.stream(
    part,
    (chunk) => {
      parser.write(decoder.decode(chunk, { stream: true }));
    },
    between,
  );
  parser.write(decoder.decode());
  parser.close();
}

function attribute(tag: SaxesTag, name: string): string | undefined {
  const value = tag.attributes[name];
  return typeof value === "string" ? value : undefined;
}

/**
 * Stream a worksheet part's rows and merged ranges. `between` runs after each
 * chunk of the part is read.
 */
export async function readWorksheetEvents(
  zip: ZipReader,
  part: string,
  events: WorksheetEvents,
  between?: () => Promise<void>,
): Promise<void> {
  let state: "before" | "rows" | "after" = "before";
  let seenDimension = false;
  let rowRef: string | undefined;
  let rowSelfClosing = false;
  let cells: RawCell[] | undefined;
  let cell:
    | {
        ref: string | undefined;
        type: string | undefined;
        style: string | undefined;
      }
    | undefined;
  let value: string | undefined;
  let inValue = false;
  let item: StringItem | undefined;
  let itemDepth = 0;
  let inlineText: string | undefined | null = null;
  let formula: RawFormula | undefined;
  let inFormula = false;

  const parser = createParser(part, {
    open(name, tag) {
      if (state === "before") {
        if (name === "sheetData") {
          state = tag.isSelfClosing ? "after" : "rows";
          if (tag.isSelfClosing) events.endRows();
        } else if (name === "dimension" && !seenDimension) {
          seenDimension = true;
          const ref = attribute(tag, "ref");
          if (ref !== undefined) events.dimension(ref);
        }
        return;
      }
      if (state === "after") {
        if (name === "mergeCell") {
          events.mergeCell(tag.attributes as Record<string, string>);
        }
        return;
      }
      if (item !== undefined) {
        itemDepth += 1;
        item.open(name);
        return;
      }
      if (cell !== undefined) {
        if (name === "v") {
          inValue = true;
          value = "";
        } else if (name === "is" && !tag.isSelfClosing) {
          item = new StringItem(false);
          itemDepth = 0;
        } else if (name === "f") {
          inFormula = true;
          formula = {
            text: "",
            type: attribute(tag, "t"),
            sharedIndex: attribute(tag, "si"),
            ref: attribute(tag, "ref"),
          };
        }
        return;
      }
      if (name === "row") {
        rowRef = attribute(tag, "r");
        rowSelfClosing = tag.isSelfClosing;
        cells = [];
      } else if (name === "c" && cells !== undefined) {
        cell = {
          ref: attribute(tag, "r"),
          type: attribute(tag, "t"),
          style: attribute(tag, "s"),
        };
        value = undefined;
        inValue = false;
        inlineText = null;
        formula = undefined;
        inFormula = false;
      }
    },
    close(name) {
      if (state !== "rows") return;
      if (item !== undefined) {
        if (itemDepth === 0 && name === "is") {
          inlineText = item.value();
          item = undefined;
        } else {
          itemDepth -= 1;
          item.close(name);
        }
        return;
      }
      if (cell !== undefined) {
        if (name === "v") {
          inValue = false;
        } else if (name === "f") {
          inFormula = false;
        } else if (name === "c") {
          cells?.push({
            ref: cell.ref,
            type: cell.type,
            style: cell.style,
            value: value === "" ? undefined : value,
            inline: inlineText,
            formula,
          });
          cell = undefined;
        }
        return;
      }
      if (name === "row" && cells !== undefined) {
        events.row(rowRef, rowSelfClosing, cells);
        cells = undefined;
      } else if (name === "sheetData") {
        state = "after";
        events.endRows();
      }
    },
    text(text, escaped) {
      if (item !== undefined) item.text(text, escaped);
      else if (inValue) value = (value ?? "") + text;
      else if (inFormula && formula !== undefined) {
        formula = { ...formula, text: formula.text + text };
      }
    },
  });
  await parsePart(zip, part, parser, between);
  // Closures above move the state on, which the checker cannot follow.
  if ((state as string) !== "after") events.endRows();
}

/**
 * The shared-string table, one entry per `<si>`, undefined for an item that
 * holds no text. Like the reader it replaces, the table ends with one extra
 * empty string after the last item, so an index one past the end reads as
 * empty text rather than failing.
 */
export async function readSharedStrings(
  zip: ZipReader,
  part: string,
): Promise<Array<string | undefined>> {
  const strings: Array<string | undefined> = [];
  let item: StringItem | undefined;
  let depth = 0;
  let inTable = false;
  const parser = createParser(part, {
    open(name, tag) {
      if (item !== undefined) {
        depth += 1;
        item.open(name);
      } else if (name === "sst") {
        inTable = true;
      } else if (inTable && name === "si") {
        if (tag.isSelfClosing) strings.push("");
        else {
          item = new StringItem(true);
          depth = 0;
        }
      }
    },
    close(name) {
      if (item === undefined) return;
      if (depth === 0 && name === "si") {
        strings.push(item.value());
        item = undefined;
      } else {
        depth -= 1;
        item.close(name);
      }
    },
    text(text, escaped) {
      item?.text(text, escaped);
    },
  });
  await parsePart(zip, part, parser);
  if (inTable) strings.push("");
  return strings;
}
