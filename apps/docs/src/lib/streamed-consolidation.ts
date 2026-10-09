/**
 * Consolidation as the operation worker runs it: each input read through
 * `Blob.slice`, never whole, and the workbook written to an output target as it
 * is produced (see `output-storage.ts`). The bytes are the command line's.
 */
import type { OperationControlOptions } from "@consultchimps/core";
import {
  consolidateWorkbookSources,
  suggestColumnMappingSources,
  type ConsolidateWorkbooksBytesOptions,
  type CsvReadOptions,
  type ConsolidateWorkbooksBytesResult,
} from "@consultchimps/xlsx/bytes";
import type { ColumnMappingSuggestion } from "@consultchimps/tabular";

import type { NamedFile, OutputFile } from "./operation-tasks";
import { openOutputTarget, type OutputPlace } from "./output-storage";
import { PieceReads } from "./piece-source";

/** The operation's name, as the library reports a cancellation. */
const CONSOLIDATE_OPERATION = "sheets.consolidate";

const WORKBOOK_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

export type StreamedConsolidationOptions = Omit<
  ConsolidateWorkbooksBytesOptions,
  "inputs" | "onProgress" | "signal" | "suggestMapping"
>;

export interface StreamedConsolidation {
  readonly result: ConsolidateWorkbooksBytesResult;
  readonly outputs: readonly OutputFile[];
}

/** Consolidate into OPFS, or into memory where OPFS is unavailable. */
export async function consolidateFiles(
  inputs: readonly NamedFile[],
  options: StreamedConsolidationOptions,
  controls: Required<OperationControlOptions>,
  place: OutputPlace,
  created: Set<string>,
): Promise<StreamedConsolidation> {
  const csvOutput =
    options.outputFormat === "csv" ||
    (options.outputFormat === undefined &&
      options.outputName?.toLowerCase().endsWith(".csv") === true);
  const target = await openOutputTarget(
    place,
    csvOutput ? "text/csv" : WORKBOOK_MEDIA_TYPE,
    created,
  );
  const reads = new PieceReads(CONSOLIDATE_OPERATION, controls.signal);
  let consolidated: Awaited<ReturnType<typeof consolidateWorkbookSources>>;
  try {
    consolidated = await reads.run(() =>
      consolidateWorkbookSources({
        ...options,
        ...controls,
        inputs: inputs.map((input) => reads.source(input)),
        output: target.sink,
      }),
    );
  } catch (error) {
    // A run the reads fail after it finished has written an output already.
    await target.sink.abort().catch(() => undefined);
    throw error;
  }
  const { result, outputName, mappingDraft } = consolidated;
  let finished: Awaited<ReturnType<typeof target.finish>>;
  try {
    finished = await target.finish();
  } catch (error) {
    await target.sink.abort().catch(() => undefined);
    throw error;
  }
  const outputs: OutputFile[] = [
    {
      name: outputName,
      blob: finished.blob,
      mediaType: result.artifacts[0]?.mediaType ?? WORKBOOK_MEDIA_TYPE,
      ...(finished.inMemory ? { inMemory: true } : {}),
    },
  ];
  if (mappingDraft) {
    outputs.push({
      name: mappingDraft.name,
      blob: new Blob([mappingDraft.bytes as Uint8Array<ArrayBuffer>], {
        type: mappingDraft.mediaType ?? "",
      }),
      mediaType: mappingDraft.mediaType,
    });
  }
  return { result, outputs };
}

/**
 * Draft a mapping from the tables a consolidation would read: the headers are
 * read and nothing is written.
 */
export async function suggestMappingFromFiles(
  inputs: readonly NamedFile[],
  includeHiddenSheets: boolean | undefined,
  controls: Required<OperationControlOptions>,
  csv?: CsvReadOptions,
): Promise<ColumnMappingSuggestion | undefined> {
  const reads = new PieceReads(CONSOLIDATE_OPERATION, controls.signal);
  return reads.run(() =>
    suggestColumnMappingSources({
      ...controls,
      inputs: inputs.map((input) => reads.source(input)),
      includeHiddenSheets,
      csv,
    }),
  );
}
