/**
 * Merging as the operation worker runs it: each input read through
 * `Blob.slice`, never whole, and the workbook written to an output target as it
 * is produced (see `output-storage.ts`). The bytes are the command line's.
 */
import type { OperationControlOptions } from "@consultchimps/core";
import {
  blobSource,
  mergeWorkbookSources,
  type MergeWorkbooksBytesOptions,
} from "@consultchimps/xlsx/bytes";

import type { NamedFile, OutputFile } from "./operation-tasks";
import { openOutputTarget, type OutputPlace } from "./output-storage";
import { PieceReads } from "./piece-source";

/** The operation's name, as the library reports a cancellation. */
const MERGE_OPERATION = "sheets.merge";

export type StreamedMergeOptions = Omit<
  MergeWorkbooksBytesOptions,
  "inputs" | "onProgress" | "signal"
>;

/** Merge into OPFS, or into memory where OPFS is unavailable. */
export async function mergeFiles(
  inputs: readonly NamedFile[],
  options: StreamedMergeOptions,
  controls: Required<OperationControlOptions>,
  place: OutputPlace,
  created: Set<string>,
): Promise<{
  readonly result: Awaited<ReturnType<typeof mergeWorkbookSources>>["result"];
  readonly outputs: readonly OutputFile[];
}> {
  // The media type is settled only once the merge knows whether a macro
  // project travels; the page relabels a download by its name, as for every
  // OPFS output.
  const target = await openOutputTarget(
    place,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    created,
  );
  const reads = new PieceReads(MERGE_OPERATION, controls.signal);
  let merged: Awaited<ReturnType<typeof mergeWorkbookSources>>;
  try {
    merged = await reads.run(() =>
      mergeWorkbookSources({
        ...options,
        ...controls,
        inputs: inputs.map((input) => reads.source(input)),
        output: target.sink,
        // Each CSV input's worksheet is written to a file of its own that the
        // merge reads as it writes, kept with this run's outputs and removed
        // with them.
        scratch: async (name) => {
          const file = await openOutputTarget(
            place,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            created,
          );
          return {
            sink: file.sink,
            finish: async () => blobSource(name, (await file.finish()).blob),
          };
        },
      }),
    );
  } catch (error) {
    await target.sink.abort().catch(() => undefined);
    throw error;
  }
  let finished: Awaited<ReturnType<typeof target.finish>>;
  try {
    finished = await target.finish();
  } catch (error) {
    await target.sink.abort().catch(() => undefined);
    throw error;
  }
  const mediaType = merged.result.artifacts[0]?.mediaType;
  return {
    result: merged.result,
    outputs: [
      {
        name: merged.outputName,
        blob: finished.blob,
        ...(mediaType === undefined ? {} : { mediaType }),
        ...(finished.inMemory ? { inMemory: true } : {}),
      },
    ],
  };
}
