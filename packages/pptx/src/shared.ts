/**
 * Platform-neutral internals shared by the path-based and byte-based
 * PowerPoint operations. This module must stay free of node:fs and node:path
 * imports so the byte entry point can run in browsers.
 */
import {
  ConsultChimpsError,
  throwIfAborted,
  type AbortOutputContext,
  type OperationControlOptions,
  type OperationResult,
} from "@consultchimps/core";
import JSZip from "jszip";

export const PRESENTATION_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";
export const PRESENTATION_EXTENSION = ".pptx";
export const POPULATE_OPERATION = "pptx.populate";
export const INSPECT_OPERATION = "pptx.inspect-template";
export const DEFAULT_TEMPLATE_SLIDE = 1;

const OFFICE_DOCUMENT_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const SLIDE_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
// Identical inputs must produce byte-identical presentations, so rewritten
// package parts carry a fixed timestamp instead of the current time.
const FIXED_ZIP_DATE = new Date("1980-01-01T00:00:00.000Z");
const TEXT_RUN_PATTERN =
  /<a:t(?<attributes>\s[^>]*)?>(?<text>[\s\S]*?)<\/a:t>/gu;
const SHAPE_PATTERN = /<p:sp(?:\s[^>]*)?>[\s\S]*?<\/p:sp>/gu;
const PLACEHOLDER_PATTERN = /\{\{(?<field>[^{}]*)\}\}/gu;

/**
 * Stable, published error codes thrown by @consultchimps/pptx. Values are
 * part of the versioned public API; never change an existing value.
 * XLSX_WORKBOOK_NOT_FOUND is intentionally XLSX-prefixed: it reports the
 * missing Excel data source of a PowerPoint operation.
 */
export const PPTX_ERRORS = {
  PPTX_GENERATION_FAILED: "PPTX_GENERATION_FAILED",
  PPTX_INVALID_DATA_SOURCE: "PPTX_INVALID_DATA_SOURCE",
  PPTX_INVALID_TEMPLATE: "PPTX_INVALID_TEMPLATE",
  PPTX_INVALID_TEMPLATE_SLIDE: "PPTX_INVALID_TEMPLATE_SLIDE",
  PPTX_MALFORMED_PLACEHOLDER: "PPTX_MALFORMED_PLACEHOLDER",
  PPTX_MISSING_EXCEL_COLUMN: "PPTX_MISSING_EXCEL_COLUMN",
  PPTX_NO_DATA_ROWS: "PPTX_NO_DATA_ROWS",
  PPTX_NO_PLACEHOLDERS: "PPTX_NO_PLACEHOLDERS",
  PPTX_OUTPUT_NOT_FILE: "PPTX_OUTPUT_NOT_FILE",
  PPTX_OUTPUT_ROLLBACK_FAILED: "PPTX_OUTPUT_ROLLBACK_FAILED",
  PPTX_TEMPLATE_NOT_FOUND: "PPTX_TEMPLATE_NOT_FOUND",
  PPTX_TEMPLATE_SLIDE_NOT_FOUND: "PPTX_TEMPLATE_SLIDE_NOT_FOUND",
  PPTX_UNSUPPORTED_OUTPUT_TYPE: "PPTX_UNSUPPORTED_OUTPUT_TYPE",
  PPTX_UNSUPPORTED_PLACEHOLDER_PLACEMENT:
    "PPTX_UNSUPPORTED_PLACEHOLDER_PLACEMENT",
  PPTX_UNSUPPORTED_SPLIT_RUN_PLACEHOLDER:
    "PPTX_UNSUPPORTED_SPLIT_RUN_PLACEHOLDER",
  PPTX_UNSUPPORTED_TEMPLATE_TYPE: "PPTX_UNSUPPORTED_TEMPLATE_TYPE",
  XLSX_WORKBOOK_NOT_FOUND: "XLSX_WORKBOOK_NOT_FOUND",
} as const;

export type PptxErrorCode = (typeof PPTX_ERRORS)[keyof typeof PPTX_ERRORS];

export type PopulatePowerPointTemplateMetric =
  | "generatedSlides"
  | "inputRows"
  | "outputFiles"
  | "placeholderFields"
  | "placeholderOccurrences"
  | "replacements"
  | "skippedRows"
  | "warnings";
export type PopulatePowerPointTemplatePlanMetric = Exclude<
  PopulatePowerPointTemplateMetric,
  "replacements" | "warnings"
>;

/**
 * The counts an inspection reports. An inspection reads one slide and writes
 * nothing, so it has no output metric; every name here is a count of what the
 * slide contains.
 */
export type InspectPowerPointTemplateMetric =
  | "malformedPlaceholderLocations"
  | "placeholderFields"
  | "placeholderOccurrences"
  | "unsupportedPlacementPlaceholders"
  | "unsupportedSplitRunPlaceholders";

export interface PowerPointPlaceholder {
  name: string;
  occurrences: number;
}

export interface PowerPointTemplateInspection {
  malformedPlaceholderCount: number;
  placeholderOccurrences: number;
  placeholders: PowerPointPlaceholder[];
  slideNumber: number;
  unsupportedPlacementPlaceholders: string[];
  unsupportedSplitRunPlaceholders: string[];
}

export interface PresentationPackage {
  contentTypesPath: string;
  presentationPath: string;
  presentationRelationshipsPath: string;
  selectedSlidePath: string;
  zip: JSZip;
}

/** Records a population reads, whatever supplied them. */
export interface PopulationRecords {
  columns: string[];
  /** The message reported when the source carries no usable rows. */
  noDataMessage: string;
  rows: Array<Record<string, string>>;
  skippedEmptyRows: number;
}

interface ShapeInspection {
  malformed: boolean;
  occurrences: string[];
  splitRunPlaceholders: string[];
}

interface SlideInspection {
  malformedPlaceholderCount: number;
  occurrences: string[];
  unsupportedPlacementPlaceholders: string[];
  unsupportedSplitRunPlaceholders: string[];
}

interface XmlRelationship {
  id: string;
  target: string;
  type: string;
}

// OOXML package part paths are POSIX-style on every platform, so these
// helpers replace node:path.posix.
function partDirectory(partPath: string): string {
  const separator = partPath.lastIndexOf("/");
  return separator < 0 ? "" : partPath.slice(0, separator);
}

function partName(partPath: string): string {
  const separator = partPath.lastIndexOf("/");
  return separator < 0 ? partPath : partPath.slice(separator + 1);
}

function normalizePartPath(partPath: string): string {
  const segments: string[] = [];
  for (const segment of partPath.split("/")) {
    if (segment === "" || segment === ".") {
      continue;
    }
    if (segment === "..") {
      const previous = segments[segments.length - 1];
      if (previous !== undefined && previous !== "..") {
        segments.pop();
      } else {
        segments.push("..");
      }
      continue;
    }
    segments.push(segment);
  }
  return segments.join("/");
}

function relativePartPath(fromDirectory: string, partPath: string): string {
  const fromSegments = fromDirectory.split("/").filter((s) => s !== "");
  const toSegments = partPath.split("/").filter((s) => s !== "");
  let common = 0;
  while (
    common < fromSegments.length &&
    common < toSegments.length &&
    fromSegments[common] === toSegments[common]
  ) {
    common += 1;
  }
  return [
    ...fromSegments.slice(common).map(() => ".."),
    ...toSegments.slice(common),
  ].join("/");
}

function xmlDecode(value: string): string {
  return value
    .replace(/&#x(?<hex>[0-9a-f]+);/giu, (_, hex: string) =>
      String.fromCodePoint(Number.parseInt(hex, 16)),
    )
    .replace(/&#(?<decimal>[0-9]+);/gu, (_, decimal: string) =>
      String.fromCodePoint(Number.parseInt(decimal, 10)),
    )
    .replace(/&lt;/gu, "<")
    .replace(/&gt;/gu, ">")
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&amp;/gu, "&");
}

function xmlEncode(value: string): string {
  return value
    .replace(/&/gu, "&amp;")
    .replace(/</gu, "&lt;")
    .replace(/>/gu, "&gt;");
}

function attributeValue(xml: string, name: string): string | undefined {
  const match = new RegExp(`\\b${name}="([^"]*)"`, "u").exec(xml);
  return match?.[1] ? xmlDecode(match[1]) : undefined;
}

function relationshipsFromXml(xml: string): XmlRelationship[] {
  return [...xml.matchAll(/<Relationship\b[^>]*\/>/gu)]
    .map((match) => {
      const relationshipXml = match[0];
      const id = attributeValue(relationshipXml, "Id");
      const target = attributeValue(relationshipXml, "Target");
      const type = attributeValue(relationshipXml, "Type");
      return id && target && type ? { id, target, type } : undefined;
    })
    .filter(
      (relationship): relationship is XmlRelationship =>
        relationship !== undefined,
    );
}

function relationshipPartPath(partPath: string): string {
  const directory = partDirectory(partPath);
  return `${directory ? `${directory}/` : ""}_rels/${partName(partPath)}.rels`;
}

function resolveRelationshipTarget(
  sourcePartPath: string,
  target: string,
): string {
  if (target.startsWith("/")) {
    return target.slice(1);
  }
  const directory = partDirectory(sourcePartPath);
  return normalizePartPath(directory ? `${directory}/${target}` : target);
}

function placeholderMatches(text: string): Array<{
  end: number;
  name: string;
  start: number;
}> {
  return [...text.matchAll(PLACEHOLDER_PATTERN)]
    .map((match) => ({
      end: (match.index ?? 0) + match[0].length,
      name: match.groups?.field?.trim() ?? "",
      start: match.index ?? 0,
    }))
    .filter((match) => match.name.length > 0);
}

function hasMalformedPlaceholder(text: string): boolean {
  const withoutValidPlaceholders = text.replace(
    PLACEHOLDER_PATTERN,
    (_, field: string) => (field.trim() ? "" : "{{}}"),
  );
  return (
    withoutValidPlaceholders.includes("{{") ||
    withoutValidPlaceholders.includes("}}")
  );
}

function inspectShape(shapeXml: string): ShapeInspection {
  const runs = [...shapeXml.matchAll(TEXT_RUN_PATTERN)].map((match) =>
    xmlDecode(match.groups?.text ?? ""),
  );
  const combinedText = runs.join("");

  return {
    malformed: hasMalformedPlaceholder(combinedText),
    occurrences: placeholderMatches(combinedText).map((match) => match.name),
    splitRunPlaceholders: [],
  };
}

function inspectSlideXml(slideXml: string): SlideInspection {
  const shapeRanges: Array<{ end: number; start: number }> = [];
  const occurrences: string[] = [];
  const unsupportedSplitRunPlaceholders: string[] = [];
  let malformedPlaceholderCount = 0;

  for (const match of slideXml.matchAll(SHAPE_PATTERN)) {
    const start = match.index ?? 0;
    shapeRanges.push({ end: start + match[0].length, start });
    const inspection = inspectShape(match[0]);
    occurrences.push(...inspection.occurrences);
    unsupportedSplitRunPlaceholders.push(...inspection.splitRunPlaceholders);
    if (inspection.malformed) {
      malformedPlaceholderCount += 1;
    }
  }

  const unsupportedPlacementPlaceholders: string[] = [];
  for (const match of slideXml.matchAll(TEXT_RUN_PATTERN)) {
    const runIndex = match.index ?? 0;
    if (
      shapeRanges.some(
        (shape) => runIndex >= shape.start && runIndex < shape.end,
      )
    ) {
      continue;
    }
    const text = xmlDecode(match.groups?.text ?? "");
    unsupportedPlacementPlaceholders.push(
      ...placeholderMatches(text).map((placeholder) => placeholder.name),
    );
    if (hasMalformedPlaceholder(text)) {
      malformedPlaceholderCount += 1;
    }
  }

  return {
    malformedPlaceholderCount,
    occurrences,
    unsupportedPlacementPlaceholders: [
      ...new Set(unsupportedPlacementPlaceholders),
    ],
    unsupportedSplitRunPlaceholders: [
      ...new Set(unsupportedSplitRunPlaceholders),
    ],
  };
}

function validatePositiveSlideNumber(templateSlide: number): void {
  if (!Number.isInteger(templateSlide) || templateSlide < 1) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_INVALID_TEMPLATE_SLIDE,
      "The template slide number must be a positive integer counted from 1.",
      { details: { templateSlide } },
    );
  }
}

export async function readRequiredZipText(
  zip: JSZip,
  entryPath: string,
): Promise<string> {
  const entry = zip.file(entryPath);
  if (!entry) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_INVALID_TEMPLATE,
      "The PowerPoint template is missing a required presentation part.",
      { details: { part: entryPath } },
    );
  }
  return entry.async("string");
}

export async function loadPresentationPackage(
  templateBytes: Uint8Array,
  templateSlide: number,
): Promise<PresentationPackage> {
  validatePositiveSlideNumber(templateSlide);

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(templateBytes);
  } catch (error) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_INVALID_TEMPLATE,
      "The template is not a readable PowerPoint .pptx presentation.",
      { cause: error },
    );
  }

  const rootRelationships = relationshipsFromXml(
    await readRequiredZipText(zip, "_rels/.rels"),
  );
  const officeDocument = rootRelationships.find(
    (relationship) => relationship.type === OFFICE_DOCUMENT_RELATIONSHIP,
  );
  if (!officeDocument) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_INVALID_TEMPLATE,
      "The template does not contain a PowerPoint presentation document.",
    );
  }

  const presentationPath = normalizePartPath(
    officeDocument.target.replace(/^\/+/u, ""),
  );
  const presentationRelationshipsPath = relationshipPartPath(presentationPath);
  const presentationXml = await readRequiredZipText(zip, presentationPath);
  const presentationRelationshipsXml = await readRequiredZipText(
    zip,
    presentationRelationshipsPath,
  );
  const slideRelationshipIds = [
    ...presentationXml.matchAll(
      /<p:sldId\b[^>]*\br:id="(?<relationshipId>[^"]+)"[^>]*\/?>/gu,
    ),
  ].map((match) => match.groups?.relationshipId ?? "");

  if (templateSlide > slideRelationshipIds.length) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_TEMPLATE_SLIDE_NOT_FOUND,
      `Template slide ${templateSlide} does not exist.`,
      {
        details: {
          availableSlides: slideRelationshipIds.length,
          templateSlide,
        },
      },
    );
  }

  const selectedRelationshipId = slideRelationshipIds[templateSlide - 1];
  const selectedRelationship = relationshipsFromXml(
    presentationRelationshipsXml,
  ).find(
    (relationship) =>
      relationship.id === selectedRelationshipId &&
      relationship.type === SLIDE_RELATIONSHIP,
  );
  if (!selectedRelationship) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_INVALID_TEMPLATE,
      "The selected template slide has an invalid package relationship.",
      { details: { templateSlide } },
    );
  }

  const selectedSlidePath = resolveRelationshipTarget(
    presentationPath,
    selectedRelationship.target,
  );
  await readRequiredZipText(zip, selectedSlidePath);
  await readRequiredZipText(zip, "[Content_Types].xml");

  return {
    contentTypesPath: "[Content_Types].xml",
    presentationPath,
    presentationRelationshipsPath,
    selectedSlidePath,
    zip,
  };
}

function publicInspection(
  inspection: SlideInspection,
  slideNumber: number,
): PowerPointTemplateInspection {
  const counts = new Map<string, number>();
  for (const name of inspection.occurrences) {
    counts.set(name, (counts.get(name) ?? 0) + 1);
  }

  return {
    malformedPlaceholderCount: inspection.malformedPlaceholderCount,
    placeholderOccurrences: inspection.occurrences.length,
    placeholders: [...counts.entries()].map(([name, occurrences]) => ({
      name,
      occurrences,
    })),
    slideNumber,
    unsupportedPlacementPlaceholders:
      inspection.unsupportedPlacementPlaceholders,
    unsupportedSplitRunPlaceholders: inspection.unsupportedSplitRunPlaceholders,
  };
}

/** Inspect the placeholders of one slide inside an already-loaded package. */
export async function inspectPresentationSlide(
  presentation: PresentationPackage,
  slideNumber: number,
): Promise<PowerPointTemplateInspection> {
  const slideXml = await readRequiredZipText(
    presentation.zip,
    presentation.selectedSlidePath,
  );
  return publicInspection(inspectSlideXml(slideXml), slideNumber);
}

function quotedList(names: readonly string[]): string {
  return names.map((name) => `"${name}"`).join(", ");
}

/**
 * The conditions an inspection reports as warnings: every one of them is a
 * reason `validateTemplateInspection` would refuse the same slide, so the
 * warnings tell a reader what a populate will do before they attempt one.
 * The inspection itself succeeded, which is why these are warnings rather
 * than errors.
 *
 * Order is fixed rather than derived from the slide, so identical inputs
 * always produce an identical result.
 */
function templateInspectionWarnings(
  inspection: PowerPointTemplateInspection,
): string[] {
  const warnings: string[] = [];
  const { slideNumber } = inspection;

  if (inspection.malformedPlaceholderCount > 0) {
    const count = inspection.malformedPlaceholderCount;
    warnings.push(
      `Slide ${slideNumber} has ${count} location${
        count === 1 ? "" : "s"
      } with malformed placeholder braces. A populate would refuse this template; use the exact {{field_name}} syntax.`,
    );
  }
  if (inspection.unsupportedSplitRunPlaceholders.length > 0) {
    warnings.push(
      `Placeholders split across multiple PowerPoint text runs are not supported: ${quotedList(
        inspection.unsupportedSplitRunPlaceholders,
      )}. A populate would refuse this template.`,
    );
  }
  if (inspection.unsupportedPlacementPlaceholders.length > 0) {
    warnings.push(
      `Placeholders outside a supported text shape are not populated: ${quotedList(
        inspection.unsupportedPlacementPlaceholders,
      )}. A populate would refuse this template.`,
    );
  }
  if (inspection.placeholderOccurrences === 0) {
    warnings.push(
      `Slide ${slideNumber} does not contain any valid {{field_name}} placeholders. A populate would refuse this template.`,
    );
  }

  return warnings;
}

/**
 * Present an inspection as the structured operation result every completed
 * ConsultChimps operation reports. An inspection creates nothing, so it has
 * no artifacts; the placeholder detail travels beside this result rather than
 * inside it, because metrics are counts and the names are not.
 */
export function templateInspectionResult(
  inspection: PowerPointTemplateInspection,
): OperationResult<InspectPowerPointTemplateMetric> {
  return {
    operation: INSPECT_OPERATION,
    artifacts: [],
    warnings: templateInspectionWarnings(inspection),
    metrics: {
      malformedPlaceholderLocations: inspection.malformedPlaceholderCount,
      placeholderFields: inspection.placeholders.length,
      placeholderOccurrences: inspection.placeholderOccurrences,
      unsupportedPlacementPlaceholders:
        inspection.unsupportedPlacementPlaceholders.length,
      unsupportedSplitRunPlaceholders:
        inspection.unsupportedSplitRunPlaceholders.length,
    },
  };
}

export function validateTemplateInspection(
  inspection: PowerPointTemplateInspection,
): void {
  if (inspection.malformedPlaceholderCount > 0) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_MALFORMED_PLACEHOLDER,
      "The template slide contains malformed placeholder braces. Use the exact {{field_name}} syntax.",
      {
        details: {
          malformedPlaceholderCount: inspection.malformedPlaceholderCount,
          templateSlide: inspection.slideNumber,
        },
      },
    );
  }
  if (inspection.unsupportedSplitRunPlaceholders.length > 0) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_UNSUPPORTED_SPLIT_RUN_PLACEHOLDER,
      `Template placeholder "${inspection.unsupportedSplitRunPlaceholders[0]}" is split across multiple PowerPoint text runs.`,
      {
        details: {
          placeholders: inspection.unsupportedSplitRunPlaceholders,
          templateSlide: inspection.slideNumber,
        },
      },
    );
  }
  if (inspection.unsupportedPlacementPlaceholders.length > 0) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_UNSUPPORTED_PLACEHOLDER_PLACEMENT,
      `Template placeholder "${inspection.unsupportedPlacementPlaceholders[0]}" is not in a supported text shape.`,
      {
        details: {
          placeholders: inspection.unsupportedPlacementPlaceholders,
          templateSlide: inspection.slideNumber,
        },
      },
    );
  }
  if (inspection.placeholderOccurrences === 0) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_NO_PLACEHOLDERS,
      "The selected template slide does not contain any valid {{field_name}} placeholders.",
      { details: { templateSlide: inspection.slideNumber } },
    );
  }
}

/**
 * Check that the records can fill the template: they must have data rows and
 * a column for every placeholder the slide uses.
 */
export function validateRecordsForTemplate(
  records: PopulationRecords,
  inspection: PowerPointTemplateInspection,
  details: Record<string, unknown>,
): void {
  if (records.rows.length === 0) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_NO_DATA_ROWS,
      records.noDataMessage,
      { details },
    );
  }

  const missingColumns = inspection.placeholders
    .map((placeholder) => placeholder.name)
    .filter((placeholder) => !records.columns.includes(placeholder));
  if (missingColumns.length > 0) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_MISSING_EXCEL_COLUMN,
      `Template placeholder "${missingColumns[0]}" does not match any Excel column.`,
      {
        details: {
          availableColumns: records.columns,
          missingColumns,
          ...details,
        },
      },
    );
  }
}

export function skippedRowsWarnings(records: PopulationRecords): string[] {
  return records.skippedEmptyRows > 0
    ? [
        `Skipped ${records.skippedEmptyRows} empty worksheet row${
          records.skippedEmptyRows === 1 ? "" : "s"
        }.`,
      ]
    : [];
}

function replaceSlideText(
  slideXml: string,
  values: Readonly<Record<string, string>>,
): { replacements: number; xml: string } {
  let replacements = 0;
  const xml = slideXml.replace(SHAPE_PATTERN, (shapeXml) => {
    const runs = [...shapeXml.matchAll(TEXT_RUN_PATTERN)].map((match) => ({
      attributes: match.groups?.attributes ?? "",
      entire: match[0],
      text: xmlDecode(match.groups?.text ?? ""),
    }));
    if (runs.length === 0) {
      return shapeXml;
    }

    const combinedText = runs.map((run) => run.text).join("");
    let textOffset = 0;
    const runBoundaries = runs.map((run) => {
      const start = textOffset;
      textOffset += run.text.length;
      return { end: textOffset, start };
    });
    const replacementsForText = placeholderMatches(combinedText).filter(
      (match) => match.name in values,
    );
    if (replacementsForText.length === 0) {
      return shapeXml;
    }

    const replacementsByRun = runs.map((run, index) => {
      const boundary = runBoundaries[index];
      if (!boundary) {
        return run.entire;
      }
      let cursor = boundary.start;
      let text = "";
      for (const match of replacementsForText) {
        if (match.end <= boundary.start || match.start >= boundary.end) {
          continue;
        }
        if (cursor < Math.min(match.start, boundary.end)) {
          text += combinedText.slice(
            cursor,
            Math.min(match.start, boundary.end),
          );
        }
        if (match.start >= boundary.start && match.start < boundary.end) {
          text += values[match.name] ?? "";
          replacements += 1;
        }
        cursor = Math.max(cursor, match.end);
        if (cursor >= boundary.end) {
          break;
        }
      }
      if (cursor < boundary.end) {
        text += combinedText.slice(cursor, boundary.end);
      }

      let attributes = run.attributes;
      if (/^\s|\s$/u.test(text) && !/\bxml:space=/u.test(attributes)) {
        attributes += ' xml:space="preserve"';
      }
      return `<a:t${attributes}>${xmlEncode(text)}</a:t>`;
    });

    let runIndex = 0;
    return shapeXml.replace(TEXT_RUN_PATTERN, () => {
      const replacement = replacementsByRun[runIndex];
      runIndex += 1;
      return replacement ?? "";
    });
  });
  return { replacements, xml };
}

function nextSlidePartPaths(zip: JSZip, count: number): string[] {
  const slideNumbers = Object.keys(zip.files)
    .map((entryPath) =>
      /^ppt\/slides\/slide(?<number>\d+)\.xml$/u.exec(entryPath),
    )
    .filter((match): match is RegExpExecArray => match !== null)
    .map((match) => Number.parseInt(match.groups?.number ?? "0", 10));
  let nextNumber = Math.max(0, ...slideNumbers) + 1;
  return Array.from({ length: count }, () => {
    const slidePath = `ppt/slides/slide${nextNumber}.xml`;
    nextNumber += 1;
    return slidePath;
  });
}

function appendBeforeClosingTag(
  xml: string,
  closingTag: string,
  addition: string,
): string {
  const index = xml.lastIndexOf(closingTag);
  if (index < 0) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_INVALID_TEMPLATE,
      "The PowerPoint template contains invalid package XML.",
      { details: { closingTag } },
    );
  }
  return `${xml.slice(0, index)}${addition}${xml.slice(index)}`;
}

function removeRelationshipsByType(xml: string, type: string): string {
  return xml.replace(/<Relationship\b[^>]*\/>/gu, (relationshipXml) =>
    attributeValue(relationshipXml, "Type") === type ? "" : relationshipXml,
  );
}

function removeContentTypeOverrides(
  xml: string,
  partPaths: ReadonlySet<string>,
): string {
  return xml.replace(/<Override\b[^>]*\/>/gu, (overrideXml) => {
    const partName = attributeValue(overrideXml, "PartName")?.replace(
      /^\/+/u,
      "",
    );
    return partName && partPaths.has(partName) ? "" : overrideXml;
  });
}

/**
 * Write a package part with a fixed timestamp and without adding folder
 * entries, so identical inputs produce byte-identical presentations and the
 * output package keeps exactly the entries the template had.
 */
function writePackagePart(zip: JSZip, partPath: string, content: string): void {
  zip.file(partPath, content, {
    createFolders: false,
    date: FIXED_ZIP_DATE,
  });
}

export interface GeneratedPresentation {
  bytes: Uint8Array;
  replacements: number;
}

/**
 * Replace the template's slides with one populated slide per record. The
 * abort context tells a cancelled caller whether completed output files may
 * remain on disk.
 */
export async function createOutputPresentation(
  presentation: PresentationPackage,
  rows: Array<Record<string, string>>,
  control: OperationControlOptions = {},
  abortContext: AbortOutputContext = "files",
): Promise<GeneratedPresentation> {
  const {
    contentTypesPath,
    presentationPath,
    presentationRelationshipsPath,
    selectedSlidePath,
    zip,
  } = presentation;
  const selectedSlideXml = await readRequiredZipText(zip, selectedSlidePath);
  const selectedSlideRelationshipsPath =
    relationshipPartPath(selectedSlidePath);
  const selectedSlideRelationships = zip.file(selectedSlideRelationshipsPath)
    ? await readRequiredZipText(zip, selectedSlideRelationshipsPath)
    : undefined;
  let presentationXml = await readRequiredZipText(zip, presentationPath);
  let presentationRelationshipsXml = await readRequiredZipText(
    zip,
    presentationRelationshipsPath,
  );
  let contentTypesXml = await readRequiredZipText(zip, contentTypesPath);

  const existingRelationships = relationshipsFromXml(
    presentationRelationshipsXml,
  );
  const originalSlidePaths = new Set(
    existingRelationships
      .filter((relationship) => relationship.type === SLIDE_RELATIONSHIP)
      .map((relationship) =>
        resolveRelationshipTarget(presentationPath, relationship.target),
      ),
  );
  const usedRelationshipIds = new Set(
    existingRelationships.map((relationship) => relationship.id),
  );
  const existingSlideIds = [
    ...presentationXml.matchAll(
      /<p:sldId\b[^>]*\bid="(?<slideId>\d+)"[^>]*\/?>/gu,
    ),
  ].map((match) => Number.parseInt(match.groups?.slideId ?? "0", 10));
  let nextSlideId = Math.max(255, ...existingSlideIds) + 1;
  let relationshipSequence = 1;
  const slideParts = nextSlidePartPaths(zip, rows.length);
  const generatedSlideReferences: string[] = [];
  const newRelationships: string[] = [];
  const newContentTypes: string[] = [];
  let replacements = 0;

  for (const [index, row] of rows.entries()) {
    throwIfAborted(control.signal, POPULATE_OPERATION, abortContext);
    const slidePath = slideParts[index];
    if (!slidePath) {
      continue;
    }
    while (usedRelationshipIds.has(`rIdConsultChimps${relationshipSequence}`)) {
      relationshipSequence += 1;
    }
    const relationshipId = `rIdConsultChimps${relationshipSequence}`;
    relationshipSequence += 1;
    usedRelationshipIds.add(relationshipId);

    const populated = replaceSlideText(selectedSlideXml, row);
    replacements += populated.replacements;
    writePackagePart(zip, slidePath, populated.xml);
    if (selectedSlideRelationships) {
      writePackagePart(
        zip,
        relationshipPartPath(slidePath),
        selectedSlideRelationships,
      );
    }

    newRelationships.push(
      `<Relationship Id="${relationshipId}" Type="${SLIDE_RELATIONSHIP}" Target="${relativePartPath(
        partDirectory(presentationPath),
        slidePath,
      )}"/>`,
    );
    generatedSlideReferences.push(
      `<p:sldId id="${nextSlideId}" r:id="${relationshipId}"/>`,
    );
    newContentTypes.push(
      `<Override PartName="/${slidePath}" ContentType="${SLIDE_CONTENT_TYPE}"/>`,
    );
    nextSlideId += 1;
    control.onProgress?.({
      operation: POPULATE_OPERATION,
      stage: "generating-slides",
      completed: index + 1,
      total: rows.length,
      detail: partName(slidePath),
    });
  }

  for (const originalSlidePath of originalSlidePaths) {
    zip.remove(originalSlidePath);
    zip.remove(relationshipPartPath(originalSlidePath));
  }
  presentationRelationshipsXml = removeRelationshipsByType(
    presentationRelationshipsXml,
    SLIDE_RELATIONSHIP,
  );
  contentTypesXml = removeContentTypeOverrides(
    contentTypesXml,
    originalSlidePaths,
  );

  const slideListPattern =
    /<p:sldIdLst(?<attributes>\s[^>]*)?>[\s\S]*?<\/p:sldIdLst>/u;
  if (!slideListPattern.test(presentationXml)) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_INVALID_TEMPLATE,
      "The PowerPoint template does not contain a valid slide list.",
    );
  }
  presentationXml = presentationXml.replace(
    slideListPattern,
    (_, attributes: string | undefined) =>
      `<p:sldIdLst${attributes ?? ""}>${generatedSlideReferences.join("")}</p:sldIdLst>`,
  );
  presentationRelationshipsXml = appendBeforeClosingTag(
    presentationRelationshipsXml,
    "</Relationships>",
    newRelationships.join(""),
  );
  contentTypesXml = appendBeforeClosingTag(
    contentTypesXml,
    "</Types>",
    newContentTypes.join(""),
  );

  writePackagePart(zip, presentationPath, presentationXml);
  writePackagePart(
    zip,
    presentationRelationshipsPath,
    presentationRelationshipsXml,
  );
  writePackagePart(zip, contentTypesPath, contentTypesXml);

  const bytes = await zip.generateAsync({
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
    platform: "DOS",
    type: "uint8array",
  });
  return { bytes, replacements };
}

export function withoutPresentationExtension(name: string): string {
  return name.replace(/\.pptx$/iu, "");
}

export { safeNameFragment } from "@consultchimps/core";
