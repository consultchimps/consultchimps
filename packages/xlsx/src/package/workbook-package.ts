/**
 * L0 - the package model.
 *
 * `WorkbookPackage` is the single owner of ZIP and part-path concerns for this
 * package. Nothing above this layer opens a workbook archive; nothing in this
 * layer interprets worksheet semantics.
 *
 * Serialization is deterministic by construction: every part this layer writes
 * carries a fixed DOS timestamp, folder entries are never created, and parts
 * are emitted in the order the source package declared them. Parts the caller
 * never wrote keep their source bytes and their source timestamp, so an
 * untouched part travels through a load/save cycle byte-identical.
 */
import JSZip from "jszip";
import { SaxesParser, type SaxesTagNS } from "saxes";

import { relationshipsPartPath, resolveRelationshipTarget } from "./paths.js";
import type {
  LoadWorkbookPackageOptions,
  PackagePart,
  RelationshipEntry,
  WorkbookPackage as WorkbookPackageContract,
} from "./types.js";

/** Every part this layer writes takes this timestamp, so outputs reproduce. */
const FIXED_PACKAGE_DATE = new Date("1980-01-01T00:00:00.000Z");

const CONTENT_TYPES_PART = "[Content_Types].xml";
/** The relationships of the package root live in a well-known part. */
const PACKAGE_ROOT = "";

/** The main workbook part, whose content type declares the package's type. */
export const WORKBOOK_MAIN_PART = "xl/workbook.xml";
/** Content type of the main workbook part in an ordinary `.xlsx` package. */
export const WORKBOOK_MAIN_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
/** Content type of the main workbook part in a macro-enabled `.xlsm` package. */
export const MACRO_WORKBOOK_MAIN_CONTENT_TYPE =
  "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
/** The opaque VBA project a macro-enabled workbook carries. */
export const VBA_PROJECT_PART = "xl/vbaProject.bin";

export type PackageRelationship = RelationshipEntry;

const decoder = new TextDecoder();
const encoder = new TextEncoder();

function attributeValue(
  tag: SaxesTagNS,
  localName: string,
): string | undefined {
  const requested = localName.toLocaleLowerCase();
  return Object.values(tag.attributes).find(
    (candidate) => candidate.local.toLocaleLowerCase() === requested,
  )?.value;
}

/**
 * Walk an XML part's opening tags with a real parser. Used for the structural
 * reads this layer owns (relationships, content types, the sheet list).
 */
export function forEachOpenTag(
  xml: string,
  fileName: string,
  onOpenTag: (tag: SaxesTagNS) => void,
): void {
  const parser = new SaxesParser({
    fileName,
    position: true,
    xmlns: true,
  } as const);
  parser.on("doctype", () => {
    throw new Error(`DOCTYPE declarations are not allowed in ${fileName}.`);
  });
  parser.on("error", (error) => {
    throw error;
  });
  parser.on("opentag", onOpenTag);
  parser.write(xml).close();
}

export { attributeValue as tagAttribute };

export class WorkbookPackage implements WorkbookPackageContract {
  /** Part bytes keyed by part path, in the order the source declared them. */
  readonly #parts = new Map<string, Uint8Array>();
  /** Source timestamps, so untouched parts round-trip unchanged. */
  readonly #dates = new Map<string, Date>();
  /** Cached relationship reads, keyed by the part that declares them. */
  readonly #relationships = new Map<string, PackageRelationship[]>();
  readonly #sourceLabel: string;

  private constructor(sourceLabel: string) {
    this.#sourceLabel = sourceLabel;
  }

  static async load(
    bytes: Uint8Array,
    options: LoadWorkbookPackageOptions = {},
  ): Promise<WorkbookPackage> {
    const archive = await JSZip.loadAsync(bytes);
    const workbookPackage = new WorkbookPackage(
      options.sourceLabel ?? "memory",
    );

    for (const entry of Object.values(archive.files)) {
      if (entry.dir) {
        continue;
      }
      workbookPackage.#parts.set(entry.name, await entry.async("uint8array"));
      workbookPackage.#dates.set(entry.name, entry.date);
    }

    return workbookPackage;
  }

  /** The label errors name this package by: a filename, or "memory". */
  get sourceLabel(): string {
    return this.#sourceLabel;
  }

  // -- The L0 seam ---------------------------------------------------------

  partPaths(): readonly string[] {
    return [...this.#parts.keys()];
  }

  part(partPath: string): PackagePart | undefined {
    const bytes = this.#parts.get(partPath);
    if (bytes === undefined) {
      return undefined;
    }
    return {
      path: partPath,
      bytes: () => Promise.resolve(bytes),
      text: () => Promise.resolve(decoder.decode(bytes)),
    };
  }

  setPartText(partPath: string, xml: string): void {
    this.setPartBytes(partPath, encoder.encode(xml));
  }

  setPartBytes(partPath: string, bytes: Uint8Array): void {
    this.#parts.set(partPath, bytes);
    this.#dates.set(partPath, FIXED_PACKAGE_DATE);
    this.#invalidateRelationships(partPath);
  }

  removePart(partPath: string): void {
    this.#parts.delete(partPath);
    this.#dates.delete(partPath);
    this.#invalidateRelationships(partPath);
  }

  /**
   * Remove every empty element named `localName` from a part and report how many
   * were removed. Both spellings of an empty element are matched, whatever
   * namespace prefix it carries: the self-closing `<sheetProtection/>` and the
   * expanded but empty `<sheetProtection></sheetProtection>`, along with their
   * `<x:sheetProtection/>` prefixed forms. Only whitespace may sit between the
   * open and close tags, so this never removes an element that carries content.
   * This is the L0 seam an operation reaches for instead of editing XML itself:
   * the whole workbook/worksheet protection strip becomes one named call.
   */
  removeEmptyElements(partPath: string, localName: string): number {
    const xml = this.readText(partPath);
    if (xml === undefined) {
      return 0;
    }
    const namePrefix = "(?:[A-Za-z_][\\w.-]*:)?";
    const pattern = new RegExp(
      `<${namePrefix}${localName}\\b[^>]*?(?:/\\s*>|>\\s*</${namePrefix}${localName}\\s*>)`,
      "gu",
    );
    let removed = 0;
    const kept = xml.replace(pattern, () => {
      removed += 1;
      return "";
    });
    if (removed > 0) {
      this.setPartText(partPath, kept);
    }
    return removed;
  }

  /**
   * Relationships are cached under the part that declares them, not under the
   * `.rels` part that stores them, so writing a `.rels` part has to drop the
   * cache rather than one entry keyed by its own path.
   */
  #invalidateRelationships(partPath: string): void {
    if (partPath.endsWith(".rels")) {
      this.#relationships.clear();
      return;
    }
    this.#relationships.delete(partPath);
  }

  relationships(
    forPartPath: string = PACKAGE_ROOT,
  ): Promise<readonly RelationshipEntry[]> {
    return Promise.resolve(this.relationshipsOf(forPartPath));
  }

  /** Drop one relationship by id from the part that declares it. */
  removeRelationship(forPartPath: string | undefined, id: string): void {
    const sourcePart = forPartPath ?? PACKAGE_ROOT;
    const partPath = relationshipsPartPath(sourcePart);
    const xml = this.readText(partPath);
    if (xml === undefined) {
      return;
    }
    const kept = removeElementsWhere(
      xml,
      "Relationship",
      (tag) => rawAttribute(tag, "Id") === id,
    );
    if (kept !== xml) {
      this.setPartText(partPath, kept);
    }
  }

  // -- Synchronous conveniences -------------------------------------------
  //
  // Part contents are decoded once at load, so the reads below need no
  // promise. The seam's async accessors are kept for callers that prefer it.

  /** Every part path, in source order. */
  partNames(): string[] {
    return [...this.#parts.keys()];
  }

  has(partPath: string): boolean {
    return this.#parts.has(partPath);
  }

  readBytes(partPath: string): Uint8Array | undefined {
    return this.#parts.get(partPath);
  }

  readText(partPath: string): string | undefined {
    const bytes = this.#parts.get(partPath);
    return bytes === undefined ? undefined : decoder.decode(bytes);
  }

  requireText(partPath: string): string {
    const text = this.readText(partPath);
    if (text === undefined) {
      throw new Error(
        `Workbook package part is missing: ${partPath} (${this.#sourceLabel})`,
      );
    }
    return text;
  }

  writeText(partPath: string, content: string): void {
    this.setPartText(partPath, content);
  }

  writeBytes(partPath: string, content: Uint8Array): void {
    this.setPartBytes(partPath, content);
  }

  remove(partPath: string): void {
    this.removePart(partPath);
  }

  /** Part paths matching a predicate, in source order. */
  partsMatching(pattern: RegExp): string[] {
    return this.partNames().filter((partPath) => pattern.test(partPath));
  }

  /** The relationships `sourcePart` declares, or none when it declares no part. */
  relationshipsOf(sourcePart: string): PackageRelationship[] {
    const cached = this.#relationships.get(sourcePart);
    if (cached) {
      return cached;
    }

    const partPath = relationshipsPartPath(sourcePart);
    const xml = this.readText(partPath);
    const relationships: PackageRelationship[] = [];

    if (xml !== undefined) {
      forEachOpenTag(xml, partPath, (tag) => {
        if (tag.local !== "Relationship") {
          return;
        }
        const id = attributeValue(tag, "Id");
        const target = attributeValue(tag, "Target");
        const type = attributeValue(tag, "Type");
        if (id && target && type) {
          relationships.push({
            id,
            target,
            type,
            targetMode:
              attributeValue(tag, "TargetMode") === "External"
                ? "External"
                : undefined,
          });
        }
      });
    }

    this.#relationships.set(sourcePart, relationships);
    return relationships;
  }

  /** Resolve a relationship target declared by `sourcePart` to a part path. */
  resolvePart(sourcePart: string, target: string): string {
    return resolveRelationshipTarget(sourcePart, target);
  }

  /**
   * Remove a part together with the relationship entries and content-type
   * override that point at it, so the package stays internally consistent.
   */
  removePartAndReferences(partPath: string, sourcePart: string): void {
    this.remove(partPath);
    this.removeRelationshipsTo(sourcePart, partPath);
    this.removeContentTypeOverride(partPath);
  }

  /** Drop every relationship `sourcePart` declares that resolves to `target`. */
  removeRelationshipsTo(sourcePart: string, target: string): void {
    const partPath = relationshipsPartPath(sourcePart);
    const xml = this.readText(partPath);
    if (xml === undefined) {
      return;
    }

    const kept = removeElementsWhere(xml, "Relationship", (tag) => {
      const relationshipTarget = rawAttribute(tag, "Target");
      if (relationshipTarget === undefined) {
        return false;
      }
      try {
        return this.resolvePart(sourcePart, relationshipTarget) === target;
      } catch {
        return false;
      }
    });
    if (kept !== xml) {
      this.setPartText(partPath, kept);
    }
  }

  /**
   * The content type `[Content_Types].xml` declares for `partPath`.
   *
   * Read with the real parser rather than by matching text, because the answer
   * decides how a package is labelled and a near-miss would be silent.
   */
  contentTypeOverride(partPath: string): string | undefined {
    const xml = this.readText(CONTENT_TYPES_PART);
    if (xml === undefined) {
      return undefined;
    }
    const target = `/${partPath}`.toLocaleLowerCase();
    let contentType: string | undefined;
    forEachOpenTag(xml, CONTENT_TYPES_PART, (tag) => {
      if (
        contentType !== undefined ||
        tag.local.toLocaleLowerCase() !== "override" ||
        attributeValue(tag, "PartName")?.toLocaleLowerCase() !== target
      ) {
        return;
      }
      contentType = attributeValue(tag, "ContentType");
    });
    return contentType;
  }

  /** Drop the `[Content_Types].xml` override that declares `partPath`. */
  removeContentTypeOverride(partPath: string): void {
    const xml = this.readText(CONTENT_TYPES_PART);
    if (xml === undefined) {
      return;
    }
    const target = `/${partPath}`;
    const kept = removeElementsWhere(
      xml,
      "Override",
      (tag) => rawAttribute(tag, "PartName") === target,
    );
    if (kept !== xml) {
      this.writeText(CONTENT_TYPES_PART, kept);
    }
  }

  /**
   * Serialize the package. Identical inputs and identical edits produce
   * identical bytes: fixed timestamps on written parts, no folder entries, and
   * source part ordering.
   */
  async save(): Promise<Uint8Array> {
    const archive = new JSZip();
    for (const [partPath, bytes] of this.#parts) {
      archive.file(partPath, bytes, {
        createFolders: false,
        date: this.#dates.get(partPath) ?? FIXED_PACKAGE_DATE,
      });
    }
    return archive.generateAsync({
      compression: "DEFLATE",
      type: "uint8array",
    });
  }
}

/**
 * Remove every self-closing element with the given local name whose opening
 * tag satisfies `predicate`. Operates on parsed tag spans rather than on a
 * pattern applied to the whole document.
 */
function removeElementsWhere(
  xml: string,
  localName: string,
  predicate: (tag: string) => boolean,
): string {
  const pattern = new RegExp(
    `<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*?/\\s*>`,
    "gu",
  );
  return xml.replace(pattern, (tag) => (predicate(tag) ? "" : tag));
}

/** Read one attribute out of a single opening tag's text. */
function rawAttribute(tag: string, name: string): string | undefined {
  const match = new RegExp(`\\b${name}=(?:"([^"]*)"|'([^']*)')`, "u").exec(tag);
  return match?.[1] ?? match?.[2];
}
