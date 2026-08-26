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
  MACRO_WORKBOOK_MAIN_CONTENT_TYPE,
  tagAttribute,
  VBA_PROJECT_PART,
  WORKBOOK_MAIN_CONTENT_TYPE,
  WORKBOOK_MAIN_PART,
  WorkbookPackage,
  type PackageRelationship,
} from "./workbook-package.js";
