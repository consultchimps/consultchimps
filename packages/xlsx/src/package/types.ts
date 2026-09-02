/**
 * L0 seam: the package model. See packages/xlsx/ARCHITECTURE.md.
 *
 * This file freezes the interface between parallel implementation streams.
 * Deviations require justification at integration time; do not edit casually.
 */

/** A single part (file) inside the OOXML package, path relative to the root. */
export interface PackagePart {
  readonly path: string;
  text(): Promise<string>;
  bytes(): Promise<Uint8Array>;
}

export interface RelationshipEntry {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode?: "External" | undefined;
}

/**
 * Deterministic, platform-pure OOXML package editor. The single owner of ZIP
 * concerns: fixed DOS dates, no synthesized folder entries, stable part
 * ordering on save. Untouched parts round-trip byte-identically.
 */
export interface WorkbookPackage {
  /** Paths of all parts, in stable package order. */
  partPaths(): readonly string[];
  part(path: string): PackagePart | undefined;
  setPartText(path: string, xml: string): void;
  setPartBytes(path: string, bytes: Uint8Array): void;
  removePart(path: string): void;
  /**
   * Remove every empty element with the given local name from a part, and report
   * how many were removed. Both spellings of an empty element are matched,
   * whatever namespace prefix it carries: the self-closing `<sheetProtection/>`
   * and the expanded but empty `<sheetProtection></sheetProtection>`, including
   * their `<x:sheetProtection/>` prefixed forms. Only whitespace may sit between
   * the open and close tags, so an element that carries content is never
   * removed. Editing the markup is L0 work: an operation asks for the change by
   * name and never touches the XML itself. Returns 0 when the part does not exist
   * or holds no matching element, and rewrites the part only when at least one
   * was removed.
   */
  removeEmptyElements(partPath: string, localName: string): number;
  /** Read the relationships part for a given part path (or the package root). */
  relationships(forPartPath?: string): Promise<readonly RelationshipEntry[]>;
  removeRelationship(forPartPath: string | undefined, id: string): void;
  /**
   * The content type `[Content_Types].xml` declares for a part path, if it
   * declares one. The main workbook part's override is what says whether a
   * package is macro-enabled, whatever the file happens to be named.
   */
  contentTypeOverride(partPath: string): string | undefined;
  /** Remove a content-type override for a part path, if present. */
  removeContentTypeOverride(partPath: string): void;
  /** Deterministic serialization of the current state. */
  save(): Promise<Uint8Array>;
}

export interface LoadWorkbookPackageOptions {
  /** Label used in error messages (a filename or "memory"). */
  readonly sourceLabel?: string | undefined;
}
