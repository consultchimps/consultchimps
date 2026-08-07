/**
 * Sheet identity shared by the Tier-1 utilities.
 *
 * Every utility in this directory has to answer the same two questions -- which
 * package part holds a sheet, and which name or calcChain index points at it --
 * so they answer them the same way here rather than three slightly different
 * ways. Parsing stays regex-based and namespace-tolerant, matching the surgical
 * package editing style of `values-only.ts`.
 */
import type JSZip from "jszip";

import { joinPackagePath, packagePartDirectory } from "../package/index.js";

const WORKBOOK_PART = "xl/workbook.xml";
const WORKBOOK_RELATIONSHIPS_PART = "xl/_rels/workbook.xml.rels";
const SHEET_ELEMENT_PATTERN = /<(?:[A-Za-z_][\w.-]*:)?sheet\b[^>]*>/gu;
const RELATIONSHIP_ELEMENT_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?Relationship\b[^>]*>/gu;
const WORKSHEET_RELATIONSHIP_SUFFIX = "/worksheet";

export interface WorkbookSheetIdentity {
  /** The sheet name formulas reference the sheet by. */
  name: string;
  /** 1-based position in the workbook's sheet order. */
  position: number;
  /** The declared `sheetId`, which calcChain entries index by. */
  sheetId: number | undefined;
  /** The worksheet package part path, e.g. `xl/worksheets/sheet1.xml`. */
  worksheetPart: string;
}

/** Resolve the five predefined XML entities in element text or an attribute. */
export function unescapeXml(value: string): string {
  return value
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

/**
 * One attribute of a start tag, matched by local name so a namespace prefix
 * such as `r:id` resolves the same as an unprefixed `id`.
 */
export function tagAttribute(
  elementXml: string,
  localName: string,
): string | undefined {
  const pattern = new RegExp(
    `(?:^|\\s)(?:[A-Za-z_][\\w.-]*:)?${localName}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`,
    "u",
  );
  const match = pattern.exec(elementXml);
  const raw = match?.[1] ?? match?.[2];
  return raw === undefined ? undefined : unescapeXml(raw);
}

/** Every worksheet the workbook part declares, in workbook order. */
export async function readWorkbookSheetIdentities(
  archive: JSZip,
): Promise<WorkbookSheetIdentity[]> {
  const workbookEntry = archive.file(WORKBOOK_PART);
  if (!workbookEntry) {
    return [];
  }
  const workbookXml = await workbookEntry.async("text");
  const relationshipsEntry = archive.file(WORKBOOK_RELATIONSHIPS_PART);
  const relationshipsXml = relationshipsEntry
    ? await relationshipsEntry.async("text")
    : "";

  const worksheetPartById = new Map<string, string>();
  for (const match of relationshipsXml.matchAll(RELATIONSHIP_ELEMENT_PATTERN)) {
    const element = match[0];
    const id = tagAttribute(element, "Id");
    const target = tagAttribute(element, "Target");
    const type = tagAttribute(element, "Type");
    if (!id || !target || !type?.endsWith(WORKSHEET_RELATIONSHIP_SUFFIX)) {
      continue;
    }
    worksheetPartById.set(
      id,
      target.startsWith("/")
        ? target.slice(1)
        : joinPackagePath(packagePartDirectory(WORKBOOK_PART), target),
    );
  }

  const identities: WorkbookSheetIdentity[] = [];
  let position = 0;
  for (const match of workbookXml.matchAll(SHEET_ELEMENT_PATTERN)) {
    const element = match[0];
    const name = tagAttribute(element, "name");
    const relationshipId = tagAttribute(element, "id");
    if (!name || !relationshipId) {
      continue;
    }
    // Chart and dialog sheets keep their place in the sheet order even though
    // they have no worksheet part, so the position counter advances first.
    position += 1;
    const worksheetPart = worksheetPartById.get(relationshipId);
    if (!worksheetPart) {
      continue;
    }
    const sheetId = Number(tagAttribute(element, "sheetId"));
    identities.push({
      name,
      position,
      sheetId: Number.isInteger(sheetId) ? sheetId : undefined,
      worksheetPart,
    });
  }
  return identities;
}
