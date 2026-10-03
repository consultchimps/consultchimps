import { relationshipsPartPath, resolveRelationshipTarget } from "./paths.js";
import type { RelationshipEntry, WorkbookPartReader } from "./types.js";
import {
  parseRelationships,
  type PackageRelationship,
} from "./workbook-package.js";
import type { ZipReader } from "./zip-reader.js";

/**
 * The small structural parts of a zip-read workbook, loaded on request and then
 * answered synchronously, so the readers written against the editable package
 * read a streamed workbook the same way without holding its worksheets.
 */
export class PreloadedParts implements WorkbookPartReader {
  readonly #zip: ZipReader;
  readonly #label: string;
  readonly #texts = new Map<string, string | undefined>();
  readonly #relationships = new Map<string, PackageRelationship[]>();

  constructor(zip: ZipReader, label: string) {
    this.#zip = zip;
    this.#label = label;
  }

  /** Load a part, once; undefined when the package does not hold it. */
  async load(partPath: string): Promise<string | undefined> {
    if (!this.#texts.has(partPath)) {
      this.#texts.set(partPath, await this.#zip.readText(partPath));
    }
    return this.#texts.get(partPath);
  }

  /** Load the relationships part that belongs to `sourcePart`. */
  async loadRelationships(sourcePart: string): Promise<void> {
    await this.load(relationshipsPartPath(sourcePart));
  }

  readText(partPath: string): string | undefined {
    if (!this.#texts.has(partPath)) {
      throw new Error(`The part ${partPath} was read before it was loaded.`);
    }
    return this.#texts.get(partPath);
  }

  requireText(partPath: string): string {
    const text = this.readText(partPath);
    if (text === undefined) {
      throw new Error(
        `Workbook package part is missing: ${partPath} (${this.#label})`,
      );
    }
    return text;
  }

  relationshipsOf(sourcePart: string): readonly RelationshipEntry[] {
    let relationships = this.#relationships.get(sourcePart);
    if (relationships === undefined) {
      const partPath = relationshipsPartPath(sourcePart);
      const xml = this.readText(partPath);
      relationships =
        xml === undefined ? [] : parseRelationships(xml, partPath);
      this.#relationships.set(sourcePart, relationships);
    }
    return relationships;
  }

  resolvePart(sourcePart: string, target: string): string {
    return resolveRelationshipTarget(sourcePart, target);
  }
}
