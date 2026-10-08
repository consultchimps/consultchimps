/**
 * Splitting as the operation worker runs it: the chosen file read through
 * `Blob.slice`, never whole, and each output written to an output target as it
 * is produced (see `output-storage.ts`). The bytes are the command line's.
 */
import type { OperationControlOptions } from "@consultchimps/core";
import {
  planSplitWorkbookSource,
  splitWorkbookSource,
  type SplitWorkbookBytesOptions,
  type SplitWorkbookBytesResult,
} from "@consultchimps/xlsx/bytes";

import type { NamedFile, OutputFile } from "./operation-tasks";
import {
  openOutputTarget,
  type OutputPlace,
  type OutputTarget,
} from "./output-storage";
import { PieceReads } from "./piece-source";

/** The operation's name, as the library reports a cancellation. */
const SPLIT_OPERATION = "sheets.split-by-column";

export type StreamedSplitOptions = Omit<
  SplitWorkbookBytesOptions,
  "input" | "onProgress" | "signal"
>;

/** Plan a split from the file read in pieces. */
export function planSplitFile(
  input: NamedFile,
  options: StreamedSplitOptions,
  signal?: AbortSignal,
): ReturnType<typeof planSplitWorkbookSource> {
  const reads = new PieceReads(SPLIT_OPERATION, signal);
  return reads.run(() =>
    planSplitWorkbookSource({ ...options, input: reads.source(input) }),
  );
}

/** Split into OPFS, or into memory where OPFS is unavailable. */
export async function splitFile(
  input: NamedFile,
  options: StreamedSplitOptions,
  controls: Required<OperationControlOptions>,
  place: OutputPlace,
  created: Set<string>,
): Promise<{
  readonly result: SplitWorkbookBytesResult;
  readonly outputs: readonly OutputFile[];
}> {
  const reads = new PieceReads(SPLIT_OPERATION, controls.signal);
  const opened: Array<{
    name: string;
    mediaType: string;
    target: OutputTarget;
    finished?: { readonly blob: Blob; readonly inMemory: boolean };
  }> = [];
  try {
    const { result } = await reads.run(() =>
      splitWorkbookSource({
        ...options,
        ...controls,
        input: reads.source(input),
        output: async (name, mediaType) => {
          const target = await openOutputTarget(place, mediaType, created);
          const entry: (typeof opened)[number] = { name, mediaType, target };
          opened.push(entry);
          return {
            write: (chunk) => target.sink.write(chunk),
            flush: async () => {
              entry.finished = await target.finish();
            },
            abort: () => target.sink.abort(),
          };
        },
      }),
    );
    return {
      result,
      outputs: opened.map((entry) => ({
        name: entry.name,
        blob: entry.finished!.blob,
        mediaType: entry.mediaType,
        ...(entry.finished!.inMemory ? { inMemory: true } : {}),
      })),
    };
  } catch (error) {
    // A run the reads fail after it finished has written its outputs already.
    for (const entry of opened) {
      await entry.target.sink.abort().catch(() => undefined);
    }
    throw error;
  }
}
