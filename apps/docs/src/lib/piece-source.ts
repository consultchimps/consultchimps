/**
 * Chosen files read in pieces through `Blob.slice`, never whole, for the
 * operations that read workbooks as streams (ADR 0006).
 */
import {
  OPERATION_ABORTED,
  throwIfAborted,
  type RandomAccessSource,
} from "@consultchimps/core";
import { blobSource } from "@consultchimps/xlsx/bytes";

import type { NamedFile } from "./operation-tasks";
import { FILE_UNREADABLE, unreadableFile } from "./unreadable-file";

/**
 * The reads of one operation over chosen files. A read the browser refuses
 * fails as the file being unreadable, as a whole-file read does, not as a
 * damaged workbook; and a cancel answers at once, even while a read of a
 * large or cloud-backed file is still waiting.
 */
export class PieceReads {
  readonly #operation: string;
  readonly #signal: AbortSignal | undefined;
  #failure: Error | undefined;

  constructor(operation: string, signal: AbortSignal | undefined) {
    this.#operation = operation;
    this.#signal = signal;
  }

  /** One chosen file, read in pieces. */
  source(input: NamedFile): RandomAccessSource {
    const source = blobSource(input.name, input.file);
    return {
      name: source.name,
      size: source.size,
      readAt: (offset, length, signal) =>
        this.#abortable(
          source.readAt(offset, length, signal).catch((error: unknown) => {
            if (error instanceof RangeError) throw error;
            this.#failure ??= unreadableFile(input.name, error);
            throw this.#failure;
          }),
        ),
    };
  }

  /**
   * Run the operation. The workbook reader reports a failed read as a damaged
   * workbook, or may not report it at all; when the file itself went
   * unreadable, or the run was cancelled, that is the answer to give.
   */
  async run<T>(work: () => Promise<T>): Promise<T> {
    try {
      return await work();
    } catch (error) {
      let cause: unknown = error;
      for (let depth = 0; cause instanceof Error && depth < 8; depth += 1) {
        const code = (cause as { code?: unknown }).code;
        if (code === OPERATION_ABORTED || code === FILE_UNREADABLE) throw cause;
        cause = cause.cause;
      }
      throw this.#failure ?? error;
    }
  }

  #abortable<T>(reading: Promise<T>): Promise<T> {
    const signal = this.#signal;
    if (signal === undefined) return reading;
    return new Promise<T>((resolve, reject) => {
      const abort = (): void => {
        try {
          throwIfAborted(signal, this.#operation, "memory");
        } catch (error) {
          reject(error);
        }
      };
      if (signal.aborted) {
        abort();
        return;
      }
      signal.addEventListener("abort", abort, { once: true });
      reading.then(
        (value) => {
          signal.removeEventListener("abort", abort);
          resolve(value);
        },
        (error: unknown) => {
          signal.removeEventListener("abort", abort);
          reject(error);
        },
      );
    });
  }
}
