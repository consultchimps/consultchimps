/**
 * Byte-level PowerPoint operations for environments without a filesystem,
 * such as browsers. Inputs and outputs are in-memory bytes; artifact paths in
 * the structured results carry portable output names. This module must stay
 * free of node:fs and node:path imports.
 */
import {
  ConsultChimpsError,
  throwIfAborted,
  type ByteArtifact,
  type ByteOperationOutcome,
  type OperationControlOptions,
  type OperationPlan,
  type OperationResult,
} from "@consultchimps/core";
import {
  readWorksheetRecordsBytes,
  type WorkbookInputBytes,
} from "@consultchimps/xlsx/bytes";

import {
  createOutputPresentation,
  DEFAULT_TEMPLATE_SLIDE,
  INSPECT_OPERATION,
  inspectPresentationSlide,
  loadPresentationPackage,
  POPULATE_OPERATION,
  PPTX_ERRORS,
  PRESENTATION_EXTENSION,
  PRESENTATION_MEDIA_TYPE,
  safeNameFragment,
  skippedRowsWarnings,
  templateInspectionResult,
  validateRecordsForTemplate,
  validateTemplateInspection,
  withoutPresentationExtension,
  type InspectPowerPointTemplateMetric,
  type PopulatePowerPointTemplateMetric,
  type PopulatePowerPointTemplatePlanMetric,
  type PopulationRecords,
  type PowerPointTemplateInspection,
  type PresentationPackage,
} from "./shared.js";

export interface PresentationInputBytes {
  name: string;
  bytes: Uint8Array;
}

/** One record per generated slide, keyed by placeholder name. */
export type PresentationRecord = Readonly<Record<string, string>>;

export interface InspectPresentationBytesOptions extends OperationControlOptions {
  templateSlide?: number | undefined;
}

export interface PopulatePresentationBytesOptions extends OperationControlOptions {
  template: PresentationInputBytes;
  /**
   * The records to populate from, supplied directly. Provide either records
   * or a workbook, never both.
   */
  records?: readonly PresentationRecord[] | undefined;
  /** The workbook whose worksheet rows become the records. */
  workbook?: WorkbookInputBytes | undefined;
  headerRow?: number | undefined;
  outputName?: string | undefined;
  templateSlide?: number | undefined;
  worksheet?: string | undefined;
}

interface ResolvedPopulateBytes {
  inspection: PowerPointTemplateInspection;
  outputName: string;
  presentation: PresentationPackage;
  records: PopulationRecords;
}

/** Column order is the order the fields first appear across the records. */
function recordColumns(records: readonly PresentationRecord[]): string[] {
  const columns: string[] = [];
  const seen = new Set<string>();

  for (const record of records) {
    for (const column of Object.keys(record)) {
      if (!seen.has(column)) {
        seen.add(column);
        columns.push(column);
      }
    }
  }

  return columns;
}

async function resolveRecords(
  options: PopulatePresentationBytesOptions,
): Promise<PopulationRecords> {
  const suppliedRecords = options.records !== undefined;
  const suppliedWorkbook = options.workbook !== undefined;
  if (suppliedRecords === suppliedWorkbook) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_INVALID_DATA_SOURCE,
      "Provide exactly one data source: either the records to populate from or the workbook bytes to read them from.",
      {
        details: {
          records: suppliedRecords,
          workbook: suppliedWorkbook,
        },
      },
    );
  }

  if (options.records) {
    return {
      columns: recordColumns(options.records),
      noDataMessage:
        "The supplied records contain no rows, so there is nothing to populate.",
      rows: options.records.map((record) => ({ ...record })),
      skippedEmptyRows: 0,
    };
  }

  const workbook = options.workbook!;
  // Last abort boundary before the workbook parse, which is synchronous and
  // cannot be interrupted once entered — a cancel that arrived while the
  // slide XML was being awaited must stop the task here.
  throwIfAborted(options.signal, POPULATE_OPERATION, "memory");
  const worksheetRecords = await readWorksheetRecordsBytes(workbook, {
    headerRow: options.headerRow,
    worksheet: options.worksheet,
  });
  return {
    columns: worksheetRecords.columns,
    noDataMessage: `Worksheet "${worksheetRecords.worksheet}" does not contain any nonempty data rows below the header.`,
    rows: worksheetRecords.rows,
    skippedEmptyRows: worksheetRecords.skippedEmptyRows,
  };
}

/**
 * Read the placeholders a template slide uses, without populating it. This is
 * the low-level report reader; `inspectPresentationOutcomeBytes` wraps the same
 * report in the structured operation result callers report to users.
 */
export async function inspectPresentationBytes(
  template: PresentationInputBytes,
  options: InspectPresentationBytesOptions = {},
): Promise<PowerPointTemplateInspection> {
  const templateSlide = options.templateSlide ?? DEFAULT_TEMPLATE_SLIDE;
  // Two package reads make up the whole cost: opening the archive, then
  // decompressing the selected slide. Both boundaries carry an abort check —
  // a page inspecting a different slide has no use for this answer and should
  // not queue behind it — and a progress event, so a caller can show which of
  // the two a large deck is currently in. The stages and their counts depend
  // only on the operation, never on the template, so they are identical for
  // identical inputs.
  throwIfAborted(options.signal, INSPECT_OPERATION, "memory");
  const presentation = await loadPresentationPackage(
    template.bytes,
    templateSlide,
  );

  throwIfAborted(options.signal, INSPECT_OPERATION, "memory");
  options.onProgress?.({
    operation: INSPECT_OPERATION,
    stage: "reading-slide",
    completed: 1,
    total: 2,
  });
  const inspection = await inspectPresentationSlide(
    presentation,
    templateSlide,
  );

  throwIfAborted(options.signal, INSPECT_OPERATION, "memory");
  options.onProgress?.({
    operation: INSPECT_OPERATION,
    stage: "inspecting-placeholders",
    completed: 2,
    total: 2,
  });
  return inspection;
}

/**
 * The outcome of a template inspection: the structured operation result every
 * completed operation reports, plus the placeholder report it describes. The
 * two travel side by side for the same reason `ByteOperationOutcome` keeps
 * `outputs` beside `result` — metrics are counts, and the placeholder names
 * are not counts.
 */
export interface PresentationInspectionOutcome {
  inspection: PowerPointTemplateInspection;
  result: OperationResult<InspectPowerPointTemplateMetric>;
}

/**
 * Inspect a template slide and report the outcome as a structured
 * `OperationResult`: counts as metrics, and one warning for every condition
 * that would make a populate refuse this template. Nothing is written, so the
 * result carries no artifacts.
 */
export async function inspectPresentationOutcomeBytes(
  template: PresentationInputBytes,
  options: InspectPresentationBytesOptions = {},
): Promise<PresentationInspectionOutcome> {
  const inspection = await inspectPresentationBytes(template, options);
  return { inspection, result: templateInspectionResult(inspection) };
}

/**
 * Resolving reads two whole packages, which is the slow half of both planning
 * and populating. The abort checks sit between those reads so a caller that
 * has moved on — a page replanning after a keystroke, say — stops paying for
 * work whose answer it will discard, instead of only ignoring it at the end.
 */
async function resolvePopulatePresentationBytes(
  options: PopulatePresentationBytesOptions,
): Promise<ResolvedPopulateBytes> {
  const templateSlide = options.templateSlide ?? DEFAULT_TEMPLATE_SLIDE;
  throwIfAborted(options.signal, POPULATE_OPERATION, "memory");
  const presentation = await loadPresentationPackage(
    options.template.bytes,
    templateSlide,
  );
  throwIfAborted(options.signal, POPULATE_OPERATION, "memory");
  const inspection = await inspectPresentationSlide(
    presentation,
    templateSlide,
  );
  validateTemplateInspection(inspection);

  const records = await resolveRecords(options);
  throwIfAborted(options.signal, POPULATE_OPERATION, "memory");
  validateRecordsForTemplate(records, inspection, {
    template: options.template.name,
    templateSlide,
  });

  const defaultName = `${withoutPresentationExtension(
    options.template.name,
  )}-populated`;
  const outputName = `${safeNameFragment(
    withoutPresentationExtension(options.outputName ?? defaultName),
    "presentation",
  )}${PRESENTATION_EXTENSION}`;

  return { inspection, outputName, presentation, records };
}

/**
 * Report the presentation a population would produce, and the rows it would
 * skip, without building any bytes.
 */
export async function planPopulatePresentationBytes(
  options: PopulatePresentationBytesOptions,
): Promise<OperationPlan<PopulatePowerPointTemplatePlanMetric>> {
  const resolved = await resolvePopulatePresentationBytes(options);

  return {
    operation: POPULATE_OPERATION,
    inputs: [
      options.template.name,
      ...(options.workbook ? [options.workbook.name] : []),
    ],
    outputs: [
      {
        kind: "file",
        mediaType: PRESENTATION_MEDIA_TYPE,
        path: resolved.outputName,
        exists: false,
      },
    ],
    warnings: skippedRowsWarnings(resolved.records),
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

/** Populate a template presentation with one slide per record, in memory. */
export async function populatePresentationBytes(
  options: PopulatePresentationBytesOptions,
): Promise<ByteOperationOutcome<PopulatePowerPointTemplateMetric>> {
  throwIfAborted(options.signal, POPULATE_OPERATION, "memory");
  const { inspection, outputName, presentation, records } =
    await resolvePopulatePresentationBytes(options);

  throwIfAborted(options.signal, POPULATE_OPERATION, "memory");
  const generated = await createOutputPresentation(
    presentation,
    records.rows,
    options,
    "memory",
  );
  // The presentation package was serialized asynchronously; honour a
  // cancellation that arrived while it was being written.
  throwIfAborted(options.signal, POPULATE_OPERATION, "memory");

  const output: ByteArtifact = {
    name: outputName,
    bytes: generated.bytes,
    mediaType: PRESENTATION_MEDIA_TYPE,
  };
  const warnings = skippedRowsWarnings(records);

  return {
    result: {
      operation: POPULATE_OPERATION,
      artifacts: [
        {
          kind: "file",
          mediaType: PRESENTATION_MEDIA_TYPE,
          path: output.name,
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
    },
    outputs: [output],
  };
}

export { PPTX_ERRORS } from "./shared.js";
export type {
  InspectPowerPointTemplateMetric,
  PopulatePowerPointTemplateMetric,
  PopulatePowerPointTemplatePlanMetric,
  PowerPointPlaceholder,
  PowerPointTemplateInspection,
  PptxErrorCode,
} from "./shared.js";
export type { WorkbookInputBytes } from "@consultchimps/xlsx/bytes";
