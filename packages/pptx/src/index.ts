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
import { readWorksheetRecords, type CsvReadOptions } from "@consultchimps/xlsx";

import {
  createOutputPresentation,
  DEFAULT_TEMPLATE_SLIDE,
  inspectPresentationSlide,
  loadPresentationPackage,
  POPULATE_OPERATION,
  PPTX_ERRORS,
  PRESENTATION_MEDIA_TYPE,
  recordsWarnings,
  validateRecordsForTemplate,
  validateTemplateInspection,
  type PopulatePowerPointTemplateMetric,
  type PopulatePowerPointTemplatePlanMetric,
  type PopulationRecords,
  type PowerPointTemplateInspection,
  type PptxErrorCode,
  type PresentationPackage,
} from "./shared.js";

export { PPTX_ERRORS } from "./shared.js";
export type {
  PopulatePowerPointTemplateMetric,
  PopulatePowerPointTemplatePlanMetric,
  PowerPointPlaceholder,
  PowerPointTemplateInspection,
  PptxErrorCode,
} from "./shared.js";

export interface InspectPowerPointTemplateOptions {
  templateSlide?: number | undefined;
}

export interface PopulatePowerPointTemplateOptions extends OperationControlOptions {
  /**
   * How to read `workbookPath` when it is a `.csv` file: its one worksheet,
   * every field text unless numbers or dates are asked for (ADR 0007).
   */
  csv?: CsvReadOptions | undefined;
  headerRow?: number | undefined;
  outputPath: string;
  overwrite?: boolean | undefined;
  templatePath: string;
  templateSlide?: number | undefined;
  workbookPath: string;
  worksheet?: string | undefined;
}

export async function inspectPowerPointTemplate(
  templatePath: string,
  options: InspectPowerPointTemplateOptions,
): Promise<PowerPointTemplateInspection> {
  const templateSlide = options.templateSlide ?? DEFAULT_TEMPLATE_SLIDE;
  const absoluteTemplate = path.resolve(templatePath);
  if (path.extname(absoluteTemplate).toLocaleLowerCase() !== ".pptx") {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_UNSUPPORTED_TEMPLATE_TYPE,
      "The PowerPoint template must be a .pptx file.",
      { details: { extension: path.extname(absoluteTemplate) } },
    );
  }

  let templateBytes: Buffer;
  try {
    templateBytes = await readFile(absoluteTemplate);
  } catch (error) {
    throw new ConsultChimpsError(
      PPTX_ERRORS.PPTX_TEMPLATE_NOT_FOUND,
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
  return inspectPresentationSlide(presentation, templateSlide);
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
  extension: ".pptx" | ".xlsx" | ".xlsx or .csv",
  errorCode: PptxErrorCode,
  label: string,
): Promise<string> {
  const absolutePath = path.resolve(filePath);
  const allowed =
    extension === ".xlsx or .csv" ? [".xlsx", ".csv"] : [extension];
  if (!allowed.includes(path.extname(absolutePath).toLocaleLowerCase())) {
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
      PPTX_ERRORS.PPTX_UNSUPPORTED_OUTPUT_TYPE,
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
        PPTX_ERRORS.PPTX_OUTPUT_NOT_FILE,
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
  bytes: Uint8Array,
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
            PPTX_ERRORS.PPTX_OUTPUT_ROLLBACK_FAILED,
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

interface ResolvedPopulate {
  absoluteOutput: string;
  absoluteTemplate: string;
  absoluteWorkbook: string;
  inspection: PowerPointTemplateInspection;
  outputExists: boolean;
  presentation: PresentationPackage;
  records: PopulationRecords;
}

async function resolvePopulatePowerPointTemplate(
  options: PopulatePowerPointTemplateOptions,
  enforceOverwrite: boolean,
): Promise<ResolvedPopulate> {
  const absoluteTemplate = await validateInputFile(
    options.templatePath,
    ".pptx",
    PPTX_ERRORS.PPTX_TEMPLATE_NOT_FOUND,
    "PowerPoint template",
  );
  const absoluteWorkbook = await validateInputFile(
    options.workbookPath,
    ".xlsx or .csv",
    PPTX_ERRORS.XLSX_WORKBOOK_NOT_FOUND,
    "data file",
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
  const inspection = await inspectPresentationSlide(
    presentation,
    templateSlide,
  );
  validateTemplateInspection(inspection);

  const worksheetRecords = await readWorksheetRecords(absoluteWorkbook, {
    csv: options.csv,
    headerRow: options.headerRow,
    worksheet: options.worksheet,
  });
  const records: PopulationRecords = {
    columns: worksheetRecords.columns,
    noDataMessage: absoluteWorkbook.toLowerCase().endsWith(".csv")
      ? `CSV file "${worksheetRecords.worksheet}" does not contain any nonempty data rows below the header.`
      : `Worksheet "${worksheetRecords.worksheet}" does not contain any nonempty data rows below the header.`,
    rows: worksheetRecords.rows,
    skippedEmptyRows: worksheetRecords.skippedEmptyRows,
    uncachedFormulas: worksheetRecords.uncachedFormulas,
    ...(worksheetRecords.warnings === undefined
      ? {}
      : { sourceWarnings: worksheetRecords.warnings }),
    csvSource: absoluteWorkbook.toLowerCase().endsWith(".csv"),
  };
  validateRecordsForTemplate(records, inspection, {
    headerRow: options.headerRow,
    worksheet: worksheetRecords.worksheet,
  });

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
): Promise<OperationPlan<PopulatePowerPointTemplatePlanMetric>> {
  const resolved = await resolvePopulatePowerPointTemplate(options, false);
  const warnings = recordsWarnings(resolved.records);
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
      csvInputFiles: resolved.records.csvSource === true ? 1 : 0,
      formulaCellsWithoutCachedValues: resolved.records.uncachedFormulas.length,
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
): Promise<OperationResult<PopulatePowerPointTemplateMetric>> {
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

    const warnings = recordsWarnings(records);
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
        csvInputFiles: records.csvSource === true ? 1 : 0,
        formulaCellsWithoutCachedValues: records.uncachedFormulas.length,
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
      PPTX_ERRORS.PPTX_GENERATION_FAILED,
      "The PowerPoint presentation could not be generated. The source files and existing output were left unchanged.",
      { cause: error },
    );
  }
}
