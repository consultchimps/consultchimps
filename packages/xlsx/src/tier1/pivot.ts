/**
 * Tier-1 confidentiality utility: remove pivot tables and their caches.
 *
 * A pivot cache is a private copy of the source rows that travels inside the
 * package. Splitting a workbook by column filters the worksheets but leaves the
 * cache intact, so every recipient can read every other group's rows out of
 * `xl/pivotCache/pivotCacheRecords*.xml`. Until the pivot cache can be filtered
 * alongside the rows it caches, removing the pivot parts is the only correct
 * answer -- the pivot table is recoverable, the leaked rows are not.
 *
 * The removal follows the technique `values-only.ts` uses for the calculation
 * chain: drop the parts, drop the relationship entries that point at them, drop
 * their content-type overrides, and drop the one workbook-level element that
 * would otherwise dangle.
 */
import JSZip from "jszip";

import { generatePackageBytes, replacePackagePart } from "../package-zip.js";

const PIVOT_PART_PREFIXES = ["xl/pivottables/", "xl/pivotcache/"] as const;
const PIVOT_TABLE_DEFINITION_PATTERN = /^xl\/pivotTables\/[^/]+\.xml$/iu;
const PIVOT_CACHE_DEFINITION_PATTERN =
  /^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/iu;
const PIVOT_RELATIONSHIP_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?Relationship\b(?=[^>]*\bType=(?:"[^"]*\/pivot(?:Table|CacheDefinition|CacheRecords)"|'[^']*\/pivot(?:Table|CacheDefinition|CacheRecords)'))[^>]*\/\s*>/gu;
const PIVOT_CONTENT_TYPE_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?Override\b(?=[^>]*\bPartName=(?:"\/xl\/pivot(?:Tables|Cache)\/[^"]*"|'\/xl\/pivot(?:Tables|Cache)\/[^']*'))[^>]*\/\s*>/gu;
const PIVOT_CACHES_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?pivotCaches\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?pivotCaches\s*>)/gu;
const CONTENT_TYPES_PART = "[Content_Types].xml";
const WORKBOOK_PART = "xl/workbook.xml";

export interface PivotStripResult {
  bytes: Buffer;
  /** Pivot cache definitions removed, one per cache. */
  removedCaches: number;
  /** Pivot table definitions removed, one per pivot table. */
  removedPivotTables: number;
}

function isPivotPart(partPath: string): boolean {
  const normalized = partPath.toLowerCase();
  return PIVOT_PART_PREFIXES.some((prefix) => normalized.startsWith(prefix));
}

async function replaceIfChanged(
  archive: JSZip,
  partPath: string,
  transform: (xml: string) => string,
): Promise<void> {
  const entry = archive.file(partPath);
  if (!entry) {
    return;
  }
  const xml = await entry.async("text");
  const rewritten = transform(xml);
  if (rewritten !== xml) {
    replacePackagePart(archive, partPath, rewritten);
  }
}

/**
 * Remove every pivot table and pivot cache part from a workbook package, along
 * with the relationships, content-type overrides and `pivotCaches` registry
 * that referenced them. Packages without pivot parts are returned byte-for-byte
 * unchanged, so wiring this in unconditionally costs nothing.
 */
export async function stripPivotParts(
  workbookBytes: Uint8Array,
): Promise<PivotStripResult> {
  const archive = await JSZip.loadAsync(workbookBytes);
  const pivotParts = Object.keys(archive.files).filter(isPivotPart);
  if (pivotParts.length === 0) {
    return {
      bytes: Buffer.from(workbookBytes),
      removedCaches: 0,
      removedPivotTables: 0,
    };
  }

  const removedPivotTables = pivotParts.filter((partPath) =>
    PIVOT_TABLE_DEFINITION_PATTERN.test(partPath),
  ).length;
  const removedCaches = pivotParts.filter((partPath) =>
    PIVOT_CACHE_DEFINITION_PATTERN.test(partPath),
  ).length;
  for (const partPath of pivotParts) {
    archive.remove(partPath);
  }

  // Relationship entries live in the workbook rels and in the rels of whichever
  // worksheet hosted the pivot table, so every remaining rels part is scrubbed.
  const relationshipParts = Object.keys(archive.files).filter((partPath) =>
    partPath.toLowerCase().endsWith(".rels"),
  );
  for (const partPath of relationshipParts) {
    await replaceIfChanged(archive, partPath, (xml) =>
      xml.replace(PIVOT_RELATIONSHIP_PATTERN, ""),
    );
  }
  await replaceIfChanged(archive, CONTENT_TYPES_PART, (xml) =>
    xml.replace(PIVOT_CONTENT_TYPE_PATTERN, ""),
  );
  // The workbook part registers each cache by relationship id; leaving the
  // registry behind would point Excel at a relationship that no longer exists.
  await replaceIfChanged(archive, WORKBOOK_PART, (xml) =>
    xml.replace(PIVOT_CACHES_PATTERN, ""),
  );

  return {
    bytes: await generatePackageBytes(archive),
    removedCaches,
    removedPivotTables,
  };
}
