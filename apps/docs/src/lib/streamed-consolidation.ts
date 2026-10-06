/**
 * Consolidation as the operation worker runs it: each input read through
 * `Blob.slice`, never whole, and the workbook written to an output target as it
 * is produced (see `output-storage.ts`). The bytes are the command line's.
 */
import type {
  OperationControlOptions,
  RandomAccessSource,
} from "@consultchimps/core";
import {
  blobSource,
  consolidateWorkbookSources,
  type ConsolidateWorkbooksBytesOptions,
  type ConsolidateWorkbooksBytesResult,
} from "@consultchimps/xlsx/bytes";
import type { ColumnMappingSuggestion } from "@consultchimps/tabular";

import type { NamedFile, OutputFile } from "./operation-tasks";
import { openOutputTarget, type OutputPlace } from "./output-storage";
import { FILE_UNREADABLE, unreadableFile } from "./unreadable-file";

const WORKBOOK_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/**
 * A chosen file read in pieces. A read the browser refuses fails as the file
 * being unreadable, as a whole-file read does, not as a damaged workbook.
 */
function pieceSource(input: NamedFile): RandomAccessSource {
  const source = blobSource(input.name, input.file);
  return {
    name: source.name,
    size: source.size,
    readAt: (offset, length, signal) =>
      source.readAt(offset, length, signal).catch((error: unknown) => {
        if (error instanceof RangeError) throw error;
        throw unreadableFile(input.name, error);
      }),
  };
}

/**
 * The workbook reader reports any failed read as a damaged workbook. When the
 * cause was the file itself going unreadable, that is the answer to give.
 */
async function readingFiles<T>(work: () => Promise<T>): Promise<T> {
  try {
    return await work();
  } catch (error) {
    let cause: unknown = error;
    for (let depth = 0; cause instanceof Error && depth < 8; depth += 1) {
      if ((cause as { code?: unknown }).code === FILE_UNREADABLE) throw cause;
      cause = cause.cause;
    }
    throw error;
  }
}

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
  const target = await openOutputTarget(place, WORKBOOK_MEDIA_TYPE, created);
  const { result, outputName, mappingDraft } = await readingFiles(() =>
    consolidateWorkbookSources({
      ...options,
      ...controls,
      inputs: inputs.map(pieceSource),
      output: target.sink,
    }),
  );
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
      mediaType: WORKBOOK_MEDIA_TYPE,
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
 * Draft a mapping from the tables a consolidation would read. The workbook the
 * run writes is not wanted, so its bytes are dropped as they are produced.
 */
export async function suggestMappingFromFiles(
  inputs: readonly NamedFile[],
  includeHiddenSheets: boolean | undefined,
  controls: Required<OperationControlOptions>,
): Promise<ColumnMappingSuggestion | undefined> {
  const { result } = await readingFiles(() =>
    consolidateWorkbookSources({
      ...controls,
      inputs: inputs.map(pieceSource),
      includeHiddenSheets,
      suggestMapping: true,
      output: {
        write: () => undefined,
        flush: () => Promise.resolve(),
        abort: () => Promise.resolve(),
      },
    }),
  );
  return result.suggestion;
}
