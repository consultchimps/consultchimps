/**
 * L0 - the package model. Deterministic OOXML load/save over parts,
 * relationships and content types, with no worksheet semantics and no
 * filesystem or platform dependency.
 *
 * The layer's seam lives in `./types.js`; `WorkbookPackage` here is the class
 * that implements it.
 */
export {
  joinPackagePath,
  normalizePackagePath,
  packagePartDirectory,
  packagePartName,
  relationshipsPartPath,
  resolveRelationshipTarget,
} from "./paths.js";
export type {
  LoadWorkbookPackageOptions,
  PackagePart,
  RelationshipEntry,
} from "./types.js";
export {
  forEachOpenTag,
  tagAttribute,
  WorkbookPackage,
  type PackageRelationship,
} from "./workbook-package.js";
