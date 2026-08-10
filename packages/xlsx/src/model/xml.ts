/**
 * Structural XML scanning for the document model.
 *
 * These helpers locate elements and attributes as spans in the source text.
 * Callers get parsed structure to decide with and exact source text to
 * serialize from, so a part that the model does not change round-trips
 * byte-identical. Nothing here is an operation-level pass over a whole
 * document: every edit is applied to one element the caller located first.
 */

export interface XmlAttributeSpan {
  /** The attribute name exactly as written, including any prefix. */
  readonly name: string;
  /** The name without its namespace prefix, lower-cased for comparison. */
  readonly localName: string;
  /** The attribute value, still XML-escaped. */
  readonly value: string;
  /** Offset of the value inside the tag text, excluding the quotes. */
  readonly valueStart: number;
  readonly valueEnd: number;
}

export interface XmlElementSpan {
  /** Offset of the opening `<`. */
  readonly start: number;
  /** Offset just past the element's final `>`. */
  readonly end: number;
  /** The element name exactly as written, including any prefix. */
  readonly name: string;
  /** The opening tag text, from `<` through its `>`. */
  readonly openTag: string;
  readonly selfClosing: boolean;
  /** Offsets of the element's content; equal when the element is empty. */
  readonly innerStart: number;
  readonly innerEnd: number;
}

const ENTITIES: Readonly<Record<string, string>> = {
  amp: "&",
  apos: "'",
  gt: ">",
  lt: "<",
  quot: '"',
};

/**
 * Resolve XML entities in attribute values and text content. Callers that
 * compare against text a person typed - header labels, number-format codes -
 * need the decoded form; callers that re-serialize keep the source text.
 */
export function decodeXmlText(text: string): string {
  if (!text.includes("&")) {
    return text;
  }
  return text.replace(
    /&(?:#(\d+)|#[xX]([0-9A-Fa-f]+)|([A-Za-z]+));/gu,
    (
      entity,
      decimal: string | undefined,
      hex: string | undefined,
      name: string | undefined,
    ) => {
      if (decimal !== undefined) {
        return String.fromCodePoint(Number(decimal));
      }
      if (hex !== undefined) {
        return String.fromCodePoint(Number.parseInt(hex, 16));
      }
      return (name !== undefined && ENTITIES[name]) || entity;
    },
  );
}

function localNameOf(name: string): string {
  const colon = name.indexOf(":");
  return (colon < 0 ? name : name.slice(colon + 1)).toLocaleLowerCase();
}

/** Parse the attributes of a single opening tag's text. */
export function parseAttributes(tagText: string): XmlAttributeSpan[] {
  const attributes: XmlAttributeSpan[] = [];
  const pattern = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/gu;
  // Skip "<" and the element name so an element name is never read as one.
  const nameEnd = /^<[^\s/>]+/u.exec(tagText)?.[0].length ?? 1;
  pattern.lastIndex = nameEnd;

  for (
    let match = pattern.exec(tagText);
    match !== null;
    match = pattern.exec(tagText)
  ) {
    const name = match[1] ?? "";
    const doubleQuoted = match[2];
    const value = doubleQuoted ?? match[3] ?? "";
    // The value starts one character past the opening quote.
    const valueStart = match.index + match[0].length - value.length - 1;
    attributes.push({
      name,
      localName: localNameOf(name),
      value,
      valueStart,
      valueEnd: valueStart + value.length,
    });
  }

  return attributes;
}

/** One attribute's value, matched on local name, or undefined. */
export function getAttribute(
  tagText: string,
  localName: string,
): string | undefined {
  const requested = localName.toLocaleLowerCase();
  return parseAttributes(tagText).find(
    (attribute) => attribute.localName === requested,
  )?.value;
}

/**
 * Replace one attribute's value in an opening tag, keeping every other
 * attribute, its quoting and the surrounding whitespace exactly as written.
 * Returns the tag unchanged when it has no such attribute.
 */
export function setAttribute(
  tagText: string,
  localName: string,
  value: string,
): string {
  const requested = localName.toLocaleLowerCase();
  const attribute = parseAttributes(tagText).find(
    (candidate) => candidate.localName === requested,
  );
  if (!attribute || attribute.value === value) {
    return tagText;
  }
  return `${tagText.slice(0, attribute.valueStart)}${value}${tagText.slice(
    attribute.valueEnd,
  )}`;
}

/** Add an attribute just before the tag's closing bracket. */
export function addAttribute(
  tagText: string,
  name: string,
  value: string,
): string {
  const closing = tagText.endsWith("/>") ? 2 : 1;
  return `${tagText.slice(0, tagText.length - closing)} ${name}="${value}"${tagText.slice(
    tagText.length - closing,
  )}`;
}

/**
 * Read the opening tag that starts at `start`, skipping quoted attribute
 * values so a `>` inside one does not end the tag early.
 */
function readOpenTag(
  xml: string,
  start: number,
): { end: number; name: string; selfClosing: boolean } | undefined {
  const nameMatch = /^<([^\s/>!?][^\s/>]*)/u.exec(xml.slice(start));
  if (!nameMatch?.[1]) {
    return undefined;
  }

  let index = start + nameMatch[0].length;
  let quote: string | undefined;
  while (index < xml.length) {
    const character = xml[index];
    if (quote) {
      if (character === quote) {
        quote = undefined;
      }
    } else if (character === '"' || character === "'") {
      quote = character;
    } else if (character === ">") {
      return {
        end: index + 1,
        name: nameMatch[1],
        selfClosing: xml[index - 1] === "/",
      };
    }
    index += 1;
  }
  return undefined;
}

/** Offset just past the markup construct starting at `start`, if it is one. */
function skipMarkupConstruct(xml: string, start: number): number | undefined {
  if (xml.startsWith("<!--", start)) {
    const end = xml.indexOf("-->", start + 4);
    return end < 0 ? xml.length : end + 3;
  }
  if (xml.startsWith("<![CDATA[", start)) {
    const end = xml.indexOf("]]>", start + 9);
    return end < 0 ? xml.length : end + 3;
  }
  if (xml.startsWith("<?", start)) {
    const end = xml.indexOf("?>", start + 2);
    return end < 0 ? xml.length : end + 2;
  }
  if (xml.startsWith("<!", start)) {
    const end = xml.indexOf(">", start + 2);
    return end < 0 ? xml.length : end + 1;
  }
  return undefined;
}

/**
 * The first element with the given local name at or after `from`, with its
 * matching close tag found by depth counting rather than by a lazy pattern.
 */
export function findElement(
  xml: string,
  localName: string,
  from = 0,
): XmlElementSpan | undefined {
  const requested = localName.toLocaleLowerCase();

  for (let index = xml.indexOf("<", from); index >= 0;) {
    const skipped = skipMarkupConstruct(xml, index);
    if (skipped !== undefined) {
      index = xml.indexOf("<", skipped);
      continue;
    }

    const openTag = readOpenTag(xml, index);
    if (!openTag) {
      index = xml.indexOf("<", index + 1);
      continue;
    }
    if (localNameOf(openTag.name) !== requested) {
      index = xml.indexOf("<", openTag.end);
      continue;
    }

    if (openTag.selfClosing) {
      return {
        start: index,
        end: openTag.end,
        name: openTag.name,
        openTag: xml.slice(index, openTag.end),
        selfClosing: true,
        innerStart: openTag.end,
        innerEnd: openTag.end,
      };
    }

    const innerEnd = findMatchingClose(xml, openTag.end, openTag.name);
    if (innerEnd === undefined) {
      return undefined;
    }
    return {
      start: index,
      end: innerEnd + `</${openTag.name}>`.length,
      name: openTag.name,
      openTag: xml.slice(index, openTag.end),
      selfClosing: false,
      innerStart: openTag.end,
      innerEnd,
    };
  }

  return undefined;
}

/** Offset of the `<` of the close tag matching an already-opened element. */
function findMatchingClose(
  xml: string,
  from: number,
  name: string,
): number | undefined {
  const closeTag = `</${name}>`;
  let depth = 1;

  for (let index = xml.indexOf("<", from); index >= 0;) {
    const skipped = skipMarkupConstruct(xml, index);
    if (skipped !== undefined) {
      index = xml.indexOf("<", skipped);
      continue;
    }
    if (xml.startsWith(closeTag, index)) {
      depth -= 1;
      if (depth === 0) {
        return index;
      }
      index = xml.indexOf("<", index + closeTag.length);
      continue;
    }
    const openTag = readOpenTag(xml, index);
    if (openTag) {
      if (openTag.name === name && !openTag.selfClosing) {
        depth += 1;
      }
      index = xml.indexOf("<", openTag.end);
      continue;
    }
    index = xml.indexOf("<", index + 1);
  }

  return undefined;
}

/** Every element with the given local name, in document order, not nested. */
export function findElements(
  xml: string,
  localName: string,
  from = 0,
  to: number | undefined = undefined,
): XmlElementSpan[] {
  const elements: XmlElementSpan[] = [];
  const limit = to ?? xml.length;
  let cursor = from;

  for (
    let element = findElement(xml, localName, cursor);
    element !== undefined && element.start < limit;
    element = findElement(xml, localName, cursor)
  ) {
    elements.push(element);
    cursor = element.end;
  }

  return elements;
}

/**
 * Rewrite each element with the given local name through `edit`. Returning
 * undefined removes the element. The scan runs once, left to right, over the
 * elements it located structurally.
 */
export function editElements(
  xml: string,
  localName: string,
  edit: (element: XmlElementSpan, text: string) => string | undefined,
): string {
  const elements = findElements(xml, localName);
  if (elements.length === 0) {
    return xml;
  }

  const pieces: string[] = [];
  let cursor = 0;
  for (const element of elements) {
    pieces.push(xml.slice(cursor, element.start));
    const replacement = edit(element, xml.slice(element.start, element.end));
    if (replacement !== undefined) {
      pieces.push(replacement);
    }
    cursor = element.end;
  }
  pieces.push(xml.slice(cursor));
  return pieces.join("");
}
