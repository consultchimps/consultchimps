/**
 * Tier-1 correctness utility: keep `xl/calcChain.xml` consistent with the rows
 * a split removed.
 *
 * The calculation chain is a cached evaluation order, not data, but Excel
 * refuses to open a workbook whose chain names cells that are gone or that now
 * hold something else. Deleting rows therefore has to prune the chain entries
 * that pointed into them and renumber the entries whose cells moved up.
 *
 * The chain omits the sheet index on any entry that shares the previous entry's
 * sheet, so pruning has to track the carried-over index and re-state it on a
 * kept entry whose predecessor was dropped.
 */
import JSZip from "jszip";

import { generatePackageBytes, replacePackagePart } from "./legacy-io.js";
import { readWorkbookSheetIdentities, tagAttribute } from "./sheets.js";

const CALC_CHAIN_PART = "xl/calcChain.xml";
const CONTENT_TYPES_PART = "[Content_Types].xml";
const WORKBOOK_RELATIONSHIPS_PART = "xl/_rels/workbook.xml.rels";
const CALC_CHAIN_ENTRY_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?c\b[^>]*?(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?c\s*>)/gu;
const CALC_CHAIN_RELATIONSHIP_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?Relationship\b(?=[^>]*\bType=(?:"[^"]*\/calcChain"|'[^']*\/calcChain'))[^>]*\/\s*>/gu;
const CALC_CHAIN_CONTENT_TYPE_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?Override\b(?=[^>]*\bPartName=(?:"\/xl\/calcChain\.xml"|'\/xl\/calcChain\.xml'))[^>]*\/\s*>/gu;
const REFERENCE_ATTRIBUTE_PATTERN =
  /((?:^|\s)(?:[A-Za-z_][\w.-]*:)?r\s*=\s*)(["'])[^"']*\2/u;
const CELL_REFERENCE_PATTERN = /^(\$?[A-Za-z]{1,3})\$?(\d+)$/u;
const ENTRY_NAME_PATTERN = /^<((?:[A-Za-z_][\w.-]*:)?c)\b/u;

/** Deleted cells, and the rows survivors moved to, keyed by worksheet part. */
export type DeletedCellsByPart = ReadonlyMap<string, ReadonlySet<string>>;
export type RenumberedRowsByPart = ReadonlyMap<
  string,
  ReadonlyMap<number, number>
>;

export interface CalcChainPruneResult {
  bytes: Buffer;
  /** Entries dropped because the cell they named was deleted. */
  removedEntries: number;
  /** Entries whose reference was rewritten because the row moved. */
  renumberedEntries: number;
}

function parseCellReference(
  reference: string,
): { column: string; row: number } | undefined {
  const match = CELL_REFERENCE_PATTERN.exec(reference);
  const column = match?.[1];
  const row = Number(match?.[2]);
  return column && Number.isInteger(row) ? { column, row } : undefined;
}

function withReference(entryXml: string, reference: string): string {
  return entryXml.replace(
    REFERENCE_ATTRIBUTE_PATTERN,
    (_match, prefix: string, quote: string) =>
      `${prefix}${quote}${reference}${quote}`,
  );
}

function withSheetIndex(entryXml: string, sheetIndex: number): string {
  const name = ENTRY_NAME_PATTERN.exec(entryXml)?.[1];
  return name === undefined
    ? entryXml
    : entryXml.replace(`<${name}`, `<${name} i="${String(sheetIndex)}"`);
}

/**
 * Drop the calculation-chain entries that name deleted cells and rewrite the
 * ones whose row moved. When nothing is left the part, its relationship and its
 * content-type override are removed exactly as a values-only conversion removes
 * them. Packages with no chain, or with nothing to change, come back unchanged.
 */
export async function pruneCalcChain(
  workbookBytes: Uint8Array,
  deletedCells: DeletedCellsByPart,
  renumberedRows?: RenumberedRowsByPart,
): Promise<CalcChainPruneResult> {
  const unchanged = {
    bytes: Buffer.from(workbookBytes),
    removedEntries: 0,
    renumberedEntries: 0,
  };
  const archive = await JSZip.loadAsync(workbookBytes);
  const calcChainEntry = archive.file(CALC_CHAIN_PART);
  if (!calcChainEntry) {
    return unchanged;
  }

  const identities = await readWorkbookSheetIdentities(archive);
  // calcChain's `i` is the sheet id; workbooks that omit sheet ids fall back to
  // the sheet's position, which is what Excel writes in practice anyway.
  const partBySheetId = new Map<number, string>();
  const partByPosition = new Map<number, string>();
  for (const identity of identities) {
    if (identity.sheetId !== undefined) {
      partBySheetId.set(identity.sheetId, identity.worksheetPart);
    }
    partByPosition.set(identity.position, identity.worksheetPart);
  }

  const calcChainXml = await calcChainEntry.async("text");
  let removedEntries = 0;
  let renumberedEntries = 0;
  let keptEntries = 0;
  let sourceSheetIndex: number | undefined;
  let emittedSheetIndex: number | undefined;

  const rewritten = calcChainXml.replace(CALC_CHAIN_ENTRY_PATTERN, (entry) => {
    const declared = Number(tagAttribute(entry, "i"));
    if (Number.isInteger(declared)) {
      sourceSheetIndex = declared;
    }
    const sheetIndex = sourceSheetIndex;
    const worksheetPart =
      sheetIndex === undefined
        ? undefined
        : (partBySheetId.get(sheetIndex) ?? partByPosition.get(sheetIndex));
    const reference = tagAttribute(entry, "r")?.toUpperCase();

    if (
      worksheetPart !== undefined &&
      reference !== undefined &&
      deletedCells.get(worksheetPart)?.has(reference) === true
    ) {
      removedEntries += 1;
      return "";
    }

    let output = entry;
    const parsed =
      worksheetPart !== undefined && reference !== undefined
        ? parseCellReference(reference)
        : undefined;
    if (parsed && worksheetPart !== undefined) {
      const destination = renumberedRows?.get(worksheetPart)?.get(parsed.row);
      if (destination !== undefined && destination !== parsed.row) {
        output = withReference(
          output,
          `${parsed.column}${String(destination)}`,
        );
        renumberedEntries += 1;
      }
    }
    // An entry that inherited its sheet from a dropped predecessor has to say
    // which sheet it belongs to itself.
    if (
      !Number.isInteger(declared) &&
      sheetIndex !== undefined &&
      sheetIndex !== emittedSheetIndex
    ) {
      output = withSheetIndex(output, sheetIndex);
    }
    emittedSheetIndex = sheetIndex;
    keptEntries += 1;
    return output;
  });

  if (removedEntries === 0 && renumberedEntries === 0) {
    return unchanged;
  }

  if (keptEntries === 0) {
    archive.remove(CALC_CHAIN_PART);
    for (const partPath of [WORKBOOK_RELATIONSHIPS_PART, CONTENT_TYPES_PART]) {
      const entry = archive.file(partPath);
      if (entry) {
        replacePackagePart(
          archive,
          partPath,
          (await entry.async("text"))
            .replace(CALC_CHAIN_RELATIONSHIP_PATTERN, "")
            .replace(CALC_CHAIN_CONTENT_TYPE_PATTERN, ""),
        );
      }
    }
  } else {
    replacePackagePart(archive, CALC_CHAIN_PART, rewritten);
  }

  return {
    bytes: await generatePackageBytes(archive),
    removedEntries,
    renumberedEntries,
  };
}
