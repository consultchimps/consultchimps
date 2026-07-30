import {
  mkdtemp,
  readFile,
  rename,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import path from "node:path";

import {
  ConsultChimpsError,
  isConsultChimpsError,
  throwIfAborted,
  type OperationControlOptions,
  type OperationPlan,
  type OperationResult,
} from "@consultchimps/core";
import {
  ensureOutputAvailable,
  ensureParentDirectory,
  refuseInputOverwrite,
} from "@consultchimps/files";
import { readWorksheetRecords } from "@consultchimps/xlsx";
import JSZip from "jszip";

export interface InspectPowerPointTemplateOptions {
  templateSlide?: number | undefined;
}

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

export interface PopulatePowerPointTemplateOptions extends OperationControlOptions {
  headerRow?: number | undefined;
  outputPath: string;
  overwrite?: boolean | undefined;
  templatePath: string;
  templateSlide?: number | undefined;
  workbookPath: string;
  worksheet?: string | undefined;
}

interface PresentationPackage {
  contentTypesPath: string;
  presentationPath: string;
  presentationRelationshipsPath: string;
  selectedSlidePath: string;
  zip: JSZip;
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

const OFFICE_DOCUMENT_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const SLIDE_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const PRESENTATION_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const DEFAULT_TEMPLATE_SLIDE = 1;
const FIXED_ZIP_DATE = new Date("1980-01-01T00:00:00.000Z");
const TEXT_RUN_PATTERN =
  /<a:t(?<attributes>\s[^>]*)?>(?<text>[\s\S]*?)<\/a:t>/gu;
const SHAPE_PATTERN = /<p:sp(?:\s[^>]*)?>[\s\S]*?<\/p:sp>/gu;
const PLACEHOLDER_PATTERN = /\{\{(?<field>[^{}]*)\}\}/gu;

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
  return path.posix.join(
    path.posix.dirname(partPath),
    "_rels",
    `${path.posix.basename(partPath)}.rels`,
  );
}

function resolveRelationshipTarget(
  sourcePartPath: string,
  target: string,
): string {
  if (target.startsWith("/")) {
    return target.slice(1);
  }
  return path.posix.normalize(
    path.posix.join(path.posix.dirname(sourcePartPath), target),
  );
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
      "PPTX_INVALID_TEMPLATE_SLIDE",
      "The template slide number must be a positive integer counted from 1.",
      { details: { templateSlide } },
    );
  }
}

async function readRequiredZipText(
  zip: JSZip,
  entryPath: string,
): Promise<string> {
  const entry = zip.file(entryPath);
  if (!entry) {
    throw new ConsultChimpsError(
      "PPTX_INVALID_TEMPLATE",
      "The PowerPoint template is missing a required presentation part.",
      { details: { part: entryPath } },
    );
  }
  return entry.async("string");
}

async function loadPresentationPackage(
  templateBytes: Buffer,
  templateSlide: number,
): Promise<PresentationPackage> {
  validatePositiveSlideNumber(templateSlide);

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(templateBytes);
  } catch (error) {
    throw new ConsultChimpsError(
      "PPTX_INVALID_TEMPLATE",
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
      "PPTX_INVALID_TEMPLATE",
      "The template does not contain a PowerPoint presentation document.",
    );
  }

  const presentationPath = path.posix
    .normalize(officeDocument.target)
    .replace(/^\/+/u, "");
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
      "PPTX_TEMPLATE_SLIDE_NOT_FOUND",
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
      "PPTX_INVALID_TEMPLATE",
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

export async function inspectPowerPointTemplate(
  templatePath: string,
  options: InspectPowerPointTemplateOptions,
): Promise<PowerPointTemplateInspection> {
  const templateSlide = options.templateSlide ?? DEFAULT_TEMPLATE_SLIDE;
  const absoluteTemplate = path.resolve(templatePath);
  if (path.extname(absoluteTemplate).toLocaleLowerCase() !== ".pptx") {
    throw new ConsultChimpsError(
      "PPTX_UNSUPPORTED_TEMPLATE_TYPE",
      "The PowerPoint template must be a .pptx file.",
      { details: { extension: path.extname(absoluteTemplate) } },
    );
  }

  let templateBytes: Buffer;
  try {
    templateBytes = await readFile(absoluteTemplate);
  } catch (error) {
    throw new ConsultChimpsError(
      "PPTX_TEMPLATE_NOT_FOUND",
      "The PowerPoint template could not be read. Check that the file exists and is accessible.",
      {
        cause: error,
        details: { templatePath: absoluteTemplate },
      },
    );
  }

  const presentation = await loadPresentationPackage(
    templateBytes,
    templateSlide,
  );
  const slideXml = await readRequiredZipText(
    presentation.zip,
    presentation.selectedSlidePath,
  );
  return publicInspection(inspectSlideXml(slideXml), templateSlide);
}

function validateTemplateInspection(
  inspection: PowerPointTemplateInspection,
): void {
  if (inspection.malformedPlaceholderCount > 0) {
    throw new ConsultChimpsError(
      "PPTX_MALFORMED_PLACEHOLDER",
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
      "PPTX_UNSUPPORTED_SPLIT_RUN_PLACEHOLDER",
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
      "PPTX_UNSUPPORTED_PLACEHOLDER_PLACEMENT",
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
      "PPTX_NO_PLACEHOLDERS",
      "The selected template slide does not contain any valid {{field_name}} placeholders.",
      { details: { templateSlide: inspection.slideNumber } },
    );
  }
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
      "PPTX_INVALID_TEMPLATE",
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

async function createOutputPresentation(
  presentation: PresentationPackage,
  rows: Array<Record<string, string>>,
  control: OperationControlOptions = {},
): Promise<{ bytes: Buffer; replacements: number }> {
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
    throwIfAborted(control.signal, POPULATE_OPERATION);
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
    zip.file(slidePath, populated.xml, { date: FIXED_ZIP_DATE });
    if (selectedSlideRelationships) {
      zip.file(relationshipPartPath(slidePath), selectedSlideRelationships, {
        date: FIXED_ZIP_DATE,
      });
    }

    const relativeTarget = path.posix.relative(
      path.posix.dirname(presentationPath),
      slidePath,
    );
    newRelationships.push(
      `<Relationship Id="${relationshipId}" Type="${SLIDE_RELATIONSHIP}" Target="${relativeTarget}"/>`,
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
      detail: path.posix.basename(slidePath),
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
      "PPTX_INVALID_TEMPLATE",
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

  zip.file(presentationPath, presentationXml, { date: FIXED_ZIP_DATE });
  zip.file(presentationRelationshipsPath, presentationRelationshipsXml, {
    date: FIXED_ZIP_DATE,
  });
  zip.file(contentTypesPath, contentTypesXml, { date: FIXED_ZIP_DATE });

  const bytes = await zip.generateAsync({
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
    platform: "DOS",
    type: "nodebuffer",
  });
  return { bytes, replacements };
}

function isMissingPathError(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    error.code === "ENOENT"
  );
}

async function validateInputFile(
  filePath: string,
  extension: ".pptx" | ".xlsx",
  errorCode: string,
  label: string,
): Promise<string> {
  const absolutePath = path.resolve(filePath);
  if (path.extname(absolutePath).toLocaleLowerCase() !== extension) {
    throw new ConsultChimpsError(
      errorCode,
      `The ${label} must be a ${extension} file.`,
      { details: { extension: path.extname(absolutePath) } },
    );
  }

  try {
    const inputStat = await stat(absolutePath);
    if (!inputStat.isFile()) {
      throw new ConsultChimpsError(
        errorCode,
        `The ${label} path is not a file.`,
      );
    }
  } catch (error) {
    if (isConsultChimpsError(error)) {
      throw error;
    }
    throw new ConsultChimpsError(
      errorCode,
      `The ${label} could not be read. Check that the file exists and is accessible.`,
      { cause: error },
    );
  }
  return absolutePath;
}

async function validateOutputPath(
  outputPath: string,
  inputPaths: string[],
  overwrite: boolean,
): Promise<{ absoluteOutput: string; outputExists: boolean }> {
  const absoluteOutput = path.resolve(outputPath);
  if (path.extname(absoluteOutput).toLocaleLowerCase() !== ".pptx") {
    throw new ConsultChimpsError(
      "PPTX_UNSUPPORTED_OUTPUT_TYPE",
      "The output presentation must use the .pptx file extension.",
      { details: { extension: path.extname(absoluteOutput) } },
    );
  }
  refuseInputOverwrite(absoluteOutput, inputPaths);

  let outputExists = false;
  try {
    const outputStat = await stat(absoluteOutput);
    if (!outputStat.isFile()) {
      throw new ConsultChimpsError(
        "PPTX_OUTPUT_NOT_FILE",
        "The output path exists but is not a file.",
        { details: { outputPath: absoluteOutput } },
      );
    }
    outputExists = true;
  } catch (error) {
    if (!isMissingPathError(error)) {
      throw error;
    }
  }
  await ensureOutputAvailable(absoluteOutput, { overwrite });
  return { absoluteOutput, outputExists };
}

async function commitOutput(
  outputPath: string,
  bytes: Buffer,
  outputExists: boolean,
): Promise<void> {
  await ensureParentDirectory(outputPath);
  const transactionDirectory = await mkdtemp(
    path.join(path.dirname(outputPath), ".consultchimps-pptx-"),
  );
  const stagedPath = path.join(transactionDirectory, "output.pptx");
  const backupPath = path.join(transactionDirectory, "previous-output.pptx");
  let backupCreated = false;

  try {
    await writeFile(stagedPath, bytes);
    if (outputExists) {
      await rename(outputPath, backupPath);
      backupCreated = true;
    }
    try {
      await rename(stagedPath, outputPath);
    } catch (error) {
      if (backupCreated) {
        try {
          await rename(backupPath, outputPath);
          backupCreated = false;
        } catch (rollbackError) {
          throw new ConsultChimpsError(
            "PPTX_OUTPUT_ROLLBACK_FAILED",
            "PowerPoint generation failed and the previous output could not be restored automatically.",
            {
              cause: error,
              details: {
                outputPath,
                rollbackError:
                  rollbackError instanceof Error
                    ? rollbackError.message
                    : String(rollbackError),
              },
            },
          );
        }
      }
      throw error;
    }
  } finally {
    await rm(transactionDirectory, { force: true, recursive: true }).catch(
      () => undefined,
    );
  }
}

const POPULATE_OPERATION = "pptx.populate";

interface ResolvedPopulate {
  absoluteOutput: string;
  absoluteTemplate: string;
  absoluteWorkbook: string;
  inspection: PowerPointTemplateInspection;
  outputExists: boolean;
  presentation: PresentationPackage;
  records: Awaited<ReturnType<typeof readWorksheetRecords>>;
}

async function resolvePopulatePowerPointTemplate(
  options: PopulatePowerPointTemplateOptions,
  enforceOverwrite: boolean,
): Promise<ResolvedPopulate> {
  const absoluteTemplate = await validateInputFile(
    options.templatePath,
    ".pptx",
    "PPTX_TEMPLATE_NOT_FOUND",
    "PowerPoint template",
  );
  const absoluteWorkbook = await validateInputFile(
    options.workbookPath,
    ".xlsx",
    "XLSX_WORKBOOK_NOT_FOUND",
    "Excel workbook",
  );
  const { absoluteOutput, outputExists } = await validateOutputPath(
    options.outputPath,
    [absoluteTemplate, absoluteWorkbook],
    enforceOverwrite ? options.overwrite === true : true,
  );

  const templateBytes = await readFile(absoluteTemplate);
  const templateSlide = options.templateSlide ?? DEFAULT_TEMPLATE_SLIDE;
  const presentation = await loadPresentationPackage(
    templateBytes,
    templateSlide,
  );
  const selectedSlideXml = await readRequiredZipText(
    presentation.zip,
    presentation.selectedSlidePath,
  );
  const inspection = publicInspection(
    inspectSlideXml(selectedSlideXml),
    templateSlide,
  );
  validateTemplateInspection(inspection);

  const records = await readWorksheetRecords(absoluteWorkbook, {
    headerRow: options.headerRow,
    worksheet: options.worksheet,
  });
  if (records.rows.length === 0) {
    throw new ConsultChimpsError(
      "PPTX_NO_DATA_ROWS",
      `Worksheet "${records.worksheet}" does not contain any nonempty data rows below the header.`,
      {
        details: {
          headerRow: options.headerRow,
          worksheet: records.worksheet,
        },
      },
    );
  }

  const missingColumns = inspection.placeholders
    .map((placeholder) => placeholder.name)
    .filter((placeholder) => !records.columns.includes(placeholder));
  if (missingColumns.length > 0) {
    throw new ConsultChimpsError(
      "PPTX_MISSING_EXCEL_COLUMN",
      `Template placeholder "${missingColumns[0]}" does not match any Excel column.`,
      {
        details: {
          availableColumns: records.columns,
          missingColumns,
          worksheet: records.worksheet,
        },
      },
    );
  }

  return {
    absoluteOutput,
    absoluteTemplate,
    absoluteWorkbook,
    inspection,
    outputExists,
    presentation,
    records,
  };
}

export async function planPopulatePowerPointTemplate(
  options: PopulatePowerPointTemplateOptions,
): Promise<OperationPlan> {
  const resolved = await resolvePopulatePowerPointTemplate(options, false);
  const warnings: string[] = [];
  if (resolved.records.skippedEmptyRows > 0) {
    warnings.push(
      `Skipped ${resolved.records.skippedEmptyRows} empty worksheet row${
        resolved.records.skippedEmptyRows === 1 ? "" : "s"
      }.`,
    );
  }
  if (resolved.outputExists && options.overwrite !== true) {
    warnings.push(
      "The planned output presentation already exists; executing without overwrite will fail.",
    );
  }

  return {
    operation: POPULATE_OPERATION,
    inputs: [resolved.absoluteTemplate, resolved.absoluteWorkbook],
    outputs: [
      {
        kind: "file",
        mediaType: PRESENTATION_MEDIA_TYPE,
        path: resolved.absoluteOutput,
        exists: resolved.outputExists,
      },
    ],
    warnings,
    metrics: {
      generatedSlides: resolved.records.rows.length,
      inputRows: resolved.records.rows.length,
      outputFiles: 1,
      placeholderFields: resolved.inspection.placeholders.length,
      placeholderOccurrences: resolved.inspection.placeholderOccurrences,
      skippedRows: resolved.records.skippedEmptyRows,
    },
  };
}

export async function populatePowerPointTemplate(
  options: PopulatePowerPointTemplateOptions,
): Promise<OperationResult> {
  try {
    throwIfAborted(options.signal, POPULATE_OPERATION);
    const { absoluteOutput, inspection, outputExists, presentation, records } =
      await resolvePopulatePowerPointTemplate(options, true);

    throwIfAborted(options.signal, POPULATE_OPERATION);
    const generated = await createOutputPresentation(
      presentation,
      records.rows,
      options,
    );
    throwIfAborted(options.signal, POPULATE_OPERATION);
    await commitOutput(absoluteOutput, generated.bytes, outputExists);
    options.onProgress?.({
      operation: POPULATE_OPERATION,
      stage: "writing-output",
      completed: 1,
      total: 1,
      detail: path.basename(absoluteOutput),
    });

    const warnings =
      records.skippedEmptyRows > 0
        ? [
            `Skipped ${records.skippedEmptyRows} empty worksheet row${
              records.skippedEmptyRows === 1 ? "" : "s"
            }.`,
          ]
        : [];
    return {
      operation: POPULATE_OPERATION,
      artifacts: [
        {
          kind: "file",
          mediaType: PRESENTATION_MEDIA_TYPE,
          path: absoluteOutput,
        },
      ],
      warnings,
      metrics: {
        generatedSlides: records.rows.length,
        inputRows: records.rows.length,
        outputFiles: 1,
        placeholderFields: inspection.placeholders.length,
        placeholderOccurrences: inspection.placeholderOccurrences,
        replacements: generated.replacements,
        skippedRows: records.skippedEmptyRows,
        warnings: warnings.length,
      },
    };
  } catch (error) {
    if (isConsultChimpsError(error)) {
      throw error;
    }
    throw new ConsultChimpsError(
      "PPTX_GENERATION_FAILED",
      "The PowerPoint presentation could not be generated. The source files and existing output were left unchanged.",
      { cause: error },
    );
  }
}
