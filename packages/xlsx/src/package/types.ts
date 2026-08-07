/**
 * L0 seam — the package model. See packages/xlsx/ARCHITECTURE.md.
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
  /** Read the relationships part for a given part path (or the package root). */
  relationships(forPartPath?: string): Promise<readonly RelationshipEntry[]>;
  removeRelationship(forPartPath: string | undefined, id: string): void;
  /** Remove a content-type override for a part path, if present. */
  removeContentTypeOverride(partPath: string): void;
  /** Deterministic serialization of the current state. */
  save(): Promise<Uint8Array>;
}

export interface LoadWorkbookPackageOptions {
  /** Label used in error messages (a filename or "memory"). */
  readonly sourceLabel?: string | undefined;
}
