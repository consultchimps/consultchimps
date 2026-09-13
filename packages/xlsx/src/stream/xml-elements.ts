import { SaxesParser, type SaxesTagNS } from "saxes";

const spreadsheetRelationships = new Map([
  [
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  ],
  [
    "http://purl.oclc.org/ooxml/spreadsheetml/main",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
  ],
]);

export const SPREADSHEET_NAMESPACES: ReadonlySet<string> = new Set(
  spreadsheetRelationships.keys(),
);
export const PACKAGE_RELATIONSHIP_NAMESPACES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/package/2006/relationships",
]);

export interface XmlElementPath {
  readonly namespace: string | undefined;
  readonly inDocument: boolean;
  is(...names: readonly string[]): boolean;
  within(...names: readonly string[]): boolean;
}

export function unqualifiedAttribute(
  tag: SaxesTagNS,
  name: string,
): string | undefined {
  const attribute = tag.attributes[name];
  return attribute?.uri === "" ? attribute.value : undefined;
}

export function relationshipAttribute(
  tag: SaxesTagNS,
  path: XmlElementPath,
): string | undefined {
  const namespace = spreadsheetRelationships.get(path.namespace ?? "");
  if (namespace === undefined) return undefined;
  return Object.values(tag.attributes).find(
    (attribute) => attribute.local === "id" && attribute.uri === namespace,
  )?.value;
}

export function createElementParser(options: {
  readonly root: string;
  readonly namespaces?: ReadonlySet<string>;
  readonly open?: (tag: SaxesTagNS, path: XmlElementPath) => void;
  readonly close?: (tag: SaxesTagNS, path: XmlElementPath) => void;
  readonly text?: (text: string, path: XmlElementPath) => void;
}): SaxesParser<{ xmlns: true }> {
  const parser = new SaxesParser({ xmlns: true });
  const elements: SaxesTagNS[] = [];
  const accepted = options.namespaces ?? SPREADSHEET_NAMESPACES;
  let namespace: string | undefined;
  let foreignDepth = 0;
  const path: XmlElementPath = {
    get namespace() {
      return namespace;
    },
    get inDocument() {
      return namespace !== undefined && foreignDepth === 0;
    },
    is(...names) {
      return elements.length === names.length && this.within(...names);
    },
    within(...names) {
      return (
        this.inDocument &&
        elements.length >= names.length &&
        names.every((name, index) => elements[index]?.local === name)
      );
    },
  };
  parser.on("doctype", () => {
    throw new Error("Document type declarations are not allowed in XLSX XML.");
  });
  parser.on("opentag", (tag) => {
    if (elements.length === 0) {
      namespace =
        tag.local === options.root && accepted.has(tag.uri)
          ? tag.uri
          : undefined;
    }
    elements.push(tag);
    if (tag.uri !== namespace) foreignDepth += 1;
    options.open?.(tag, path);
  });
  parser.on("closetag", (tag) => {
    options.close?.(tag, path);
    if (tag.uri !== namespace) foreignDepth -= 1;
    elements.pop();
  });
  const append = (text: string) => options.text?.(text, path);
  parser.on("text", append);
  parser.on("cdata", append);
  return parser;
}
