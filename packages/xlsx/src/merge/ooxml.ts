/**
 * Package-shaped edits the merge transplant needs on top of L0.
 *
 * `WorkbookPackage` owns loading, saving and removing parts. Growing a package
 * - claiming a free part path, declaring a content type, adding a relationship
 * - is only ever needed when parts are copied INTO a workbook, which is what
 * the merge does and nothing else does yet. The helpers live here rather than
 * in L0 so the layer keeps exactly one reason to change while Phase 1 lands.
 *
 * Nothing here interprets worksheet semantics; every function is about part
 * paths, `[Content_Types].xml` and `.rels` parts.
 */
import {
  packagePartDirectory,
  packagePartName,
  relationshipsPartPath,
  type PackageRelationship,
  type WorkbookPackage,
} from "../package/index.js";
import { getAttribute } from "../model/xml.js";

export const CONTENT_TYPES_PART = "[Content_Types].xml";
export const WORKBOOK_PART = "xl/workbook.xml";

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const RELATIONSHIPS_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/relationships";

/** Escape a value for use inside an XML attribute. */
export function escapeXmlAttribute(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

/** Escape a value for use as XML text content. */
export function escapeXmlText(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

/** The file extension of a part path, lower-cased, without the dot. */
export function partExtension(partPath: string): string {
  const name = packagePartName(partPath);
  const dot = name.lastIndexOf(".");
  return dot < 0 ? "" : name.slice(dot + 1).toLocaleLowerCase();
}

/**
 * The part name without its trailing digits and extension, so
 * `xl/worksheets/sheet12.xml` yields `sheet`. New parts are numbered from that
 * stem, which keeps a transplanted package readable by a human opening the zip.
 */
export function partStem(partPath: string): string {
  const name = packagePartName(partPath);
  const dot = name.lastIndexOf(".");
  const base = dot < 0 ? name : name.slice(0, dot);
  return base.replace(/\d+$/u, "") || "part";
}

/**
 * Claim the first free `<directory>/<stem><n>.<extension>` path, counting from
 * one. Deterministic: the same package and the same request always produce the
 * same path, which is what lets identical inputs merge to identical bytes.
 */
export function allocatePartPath(
  workbookPackage: WorkbookPackage,
  directory: string,
  stem: string,
  extension: string,
): string {
  const suffix = extension === "" ? "" : `.${extension}`;
  const prefix = directory === "" ? "" : `${directory}/`;
  for (let index = 1; ; index += 1) {
    const candidate = `${prefix}${stem}${index}${suffix}`;
    if (!workbookPackage.has(candidate)) {
      return candidate;
    }
  }
}

// -- Content types ---------------------------------------------------------

/** The content type declared for a part, by override first, then by extension. */
export function contentTypeOf(
  workbookPackage: WorkbookPackage,
  partPath: string,
): string | undefined {
  const xml = workbookPackage.readText(CONTENT_TYPES_PART);
  if (xml === undefined) {
    return undefined;
  }
  const target = `/${partPath}`.toLocaleLowerCase();
  for (const match of xml.matchAll(/<Override\b[^>]*\/>/gu)) {
    const element = match[0];
    if (
      getAttribute(element, "PartName")?.toLocaleLowerCase() === target &&
      getAttribute(element, "ContentType") !== undefined
    ) {
      return getAttribute(element, "ContentType");
    }
  }
  const extension = partExtension(partPath);
  for (const match of xml.matchAll(/<Default\b[^>]*\/>/gu)) {
    const element = match[0];
    if (
      getAttribute(element, "Extension")?.toLocaleLowerCase() === extension &&
      getAttribute(element, "ContentType") !== undefined
    ) {
      return getAttribute(element, "ContentType");
    }
  }
  return undefined;
}

function appendToContentTypes(
  workbookPackage: WorkbookPackage,
  element: string,
): void {
  const xml = workbookPackage.readText(CONTENT_TYPES_PART);
  if (xml === undefined) {
    return;
  }
  const closing = xml.lastIndexOf("</Types>");
  if (closing < 0) {
    return;
  }
  workbookPackage.writeText(
    CONTENT_TYPES_PART,
    `${xml.slice(0, closing)}${element}${xml.slice(closing)}`,
  );
}

/** Declare a part's content type, unless an override already covers it. */
export function addContentTypeOverride(
  workbookPackage: WorkbookPackage,
  partPath: string,
  contentType: string,
): void {
  const xml = workbookPackage.readText(CONTENT_TYPES_PART) ?? "";
  const target = `/${partPath}`;
  if (xml.includes(`PartName="${target}"`)) {
    return;
  }
  appendToContentTypes(
    workbookPackage,
    `<Override PartName="${escapeXmlAttribute(target)}" ContentType="${escapeXmlAttribute(
      contentType,
    )}"/>`,
  );
}

/**
 * Declare a part's content type, replacing any override it already has. Used
 * for the workbook part, whose type states whether the package is macro
 * enabled and therefore changes when the macro project does.
 */
export function setContentTypeOverride(
  workbookPackage: WorkbookPackage,
  partPath: string,
  contentType: string,
): void {
  const xml = workbookPackage.readText(CONTENT_TYPES_PART);
  if (xml === undefined) {
    return;
  }
  const target = `/${partPath}`.toLocaleLowerCase();
  const rewritten = xml.replace(/<Override\b[^>]*\/>/gu, (element) =>
    getAttribute(element, "PartName")?.toLocaleLowerCase() === target
      ? `<Override PartName="${escapeXmlAttribute(`/${partPath}`)}" ContentType="${escapeXmlAttribute(
          contentType,
        )}"/>`
      : element,
  );
  if (rewritten === xml) {
    addContentTypeOverride(workbookPackage, partPath, contentType);
    return;
  }
  workbookPackage.writeText(CONTENT_TYPES_PART, rewritten);
}

/** Declare a default content type for an extension, unless one already exists. */
export function addDefaultContentType(
  workbookPackage: WorkbookPackage,
  extension: string,
  contentType: string,
): void {
  const xml = workbookPackage.readText(CONTENT_TYPES_PART);
  if (xml === undefined) {
    return;
  }
  const wanted = extension.toLocaleLowerCase();
  for (const match of xml.matchAll(/<Default\b[^>]*\/>/gu)) {
    if (getAttribute(match[0], "Extension")?.toLocaleLowerCase() === wanted) {
      return;
    }
  }
  appendToContentTypes(
    workbookPackage,
    `<Default Extension="${escapeXmlAttribute(extension)}" ContentType="${escapeXmlAttribute(
      contentType,
    )}"/>`,
  );
}

/**
 * Carry a source part's content-type declaration into the output package: an
 * override for a part with one, a default for the extension otherwise.
 */
export function copyContentTypeDeclaration(
  source: WorkbookPackage,
  sourcePartPath: string,
  target: WorkbookPackage,
  targetPartPath: string,
): void {
  const declared = contentTypeOf(source, sourcePartPath);
  if (declared === undefined) {
    return;
  }
  const xml = source.readText(CONTENT_TYPES_PART) ?? "";
  const hasOverride = xml.includes(`PartName="/${sourcePartPath}"`);
  if (hasOverride) {
    addContentTypeOverride(target, targetPartPath, declared);
    return;
  }
  addDefaultContentType(target, partExtension(targetPartPath), declared);
}

// -- Relationships ---------------------------------------------------------

/** Serialize a relationship list as a `.rels` part. */
export function relationshipsXml(
  relationships: readonly PackageRelationship[],
): string {
  const entries = relationships
    .map(
      (relationship) =>
        `<Relationship Id="${escapeXmlAttribute(relationship.id)}" Type="${escapeXmlAttribute(
          relationship.type,
        )}" Target="${escapeXmlAttribute(relationship.target)}"${
          relationship.targetMode === "External" ? ' TargetMode="External"' : ""
        }/>`,
    )
    .join("");
  return `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">${entries}</Relationships>`;
}

/** Replace the relationships a part declares. */
export function writeRelationships(
  workbookPackage: WorkbookPackage,
  sourcePart: string,
  relationships: readonly PackageRelationship[],
): void {
  const partPath = relationshipsPartPath(sourcePart);
  if (relationships.length === 0) {
    workbookPackage.remove(partPath);
    return;
  }
  workbookPackage.writeText(partPath, relationshipsXml(relationships));
}

/**
 * Add one relationship to a part, returning the fresh id. Ids are unique
 * within the declaring part, so the counter restarts for every part.
 */
export function addRelationship(
  workbookPackage: WorkbookPackage,
  sourcePart: string,
  relationship: Omit<PackageRelationship, "id">,
): string {
  const existing = workbookPackage.relationshipsOf(sourcePart);
  const used = new Set(existing.map((entry) => entry.id));
  let index = existing.length + 1;
  while (used.has(`rId${index}`)) {
    index += 1;
  }
  const id = `rId${index}`;
  writeRelationships(workbookPackage, sourcePart, [
    ...existing,
    { ...relationship, id },
  ]);
  return id;
}

/**
 * A relationship target expressed relative to the part that declares it.
 * Package part paths are absolute; relationships are conventionally written
 * relative, which is what Excel emits and what keeps a package portable.
 */
export function relativeRelationshipTarget(
  sourcePart: string,
  targetPart: string,
): string {
  const from = packagePartDirectory(sourcePart).split("/").filter(Boolean);
  const to = targetPart.split("/");
  const name = to.pop() ?? "";

  let shared = 0;
  while (
    shared < from.length &&
    shared < to.length &&
    from[shared] === to[shared]
  ) {
    shared += 1;
  }
  const up = new Array(from.length - shared).fill("..");
  return [...up, ...to.slice(shared), name].join("/");
}
