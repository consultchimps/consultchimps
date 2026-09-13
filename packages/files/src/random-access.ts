import { mkdtemp, open, rm, stat } from "node:fs/promises";
import type { FileHandle } from "node:fs/promises";
import path from "node:path";

import {
  ConsultChimpsError,
  throwIfAborted,
  type RandomAccessFile,
  type RandomAccessSource,
} from "@consultchimps/core";

export interface FileSource extends RandomAccessSource {
  verifyUnchanged(): Promise<void>;
  close(): Promise<void>;
}

export interface ScratchDirectory {
  create(): Promise<RandomAccessFile>;
  close(): Promise<void>;
}

function assertRange(offset: number, length: number): void {
  if (
    !Number.isSafeInteger(offset) ||
    !Number.isSafeInteger(length) ||
    offset < 0 ||
    length < 0 ||
    !Number.isSafeInteger(offset + length)
  ) {
    throw new ConsultChimpsError(
      "FILES_INVALID_RANGE",
      "The requested file range is invalid. Use nonnegative whole-byte offsets and lengths.",
    );
  }
}

async function readRange(
  handle: FileHandle,
  size: number,
  offset: number,
  length: number,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  assertRange(offset, length);
  throwIfAborted(signal, "files.read");
  const bytes = new Uint8Array(Math.min(length, Math.max(0, size - offset)));
  let completed = 0;
  while (completed < bytes.length) {
    throwIfAborted(signal, "files.read");
    const result = await handle.read(
      bytes,
      completed,
      bytes.length - completed,
      offset + completed,
    );
    if (result.bytesRead === 0) break;
    completed += result.bytesRead;
  }
  return bytes.subarray(0, completed);
}

export async function openRandomAccessSource(
  filePath: string,
): Promise<FileSource> {
  let handle: FileHandle;
  try {
    handle = await open(filePath, "r");
  } catch (cause) {
    throw new ConsultChimpsError(
      "FILES_INPUT_UNREADABLE",
      "The input file could not be opened. Check that it exists and that you can read it.",
      { cause },
    );
  }
  try {
    const baseline = await handle.stat({ bigint: true });
    if (!baseline.isFile() || baseline.size > BigInt(Number.MAX_SAFE_INTEGER)) {
      throw new ConsultChimpsError(
        "FILES_INVALID_SOURCE",
        "Choose a regular file whose size can be addressed safely.",
      );
    }
    return {
      name: path.basename(filePath),
      size: Number(baseline.size),
      readAt: (offset, length, signal) =>
        readRange(handle, Number(baseline.size), offset, length, signal),
      async verifyUnchanged() {
        const current = await handle.stat({ bigint: true });
        let named;
        try {
          named = await stat(filePath, { bigint: true });
        } catch (cause) {
          throw new ConsultChimpsError(
            "FILES_SOURCE_CHANGED",
            "The input file moved or changed while it was being read. Retry with a stable copy.",
            { cause },
          );
        }
        if (
          [current, named].some(
            (value) =>
              value.dev !== baseline.dev ||
              value.ino !== baseline.ino ||
              value.size !== baseline.size ||
              value.mtimeNs !== baseline.mtimeNs ||
              value.ctimeNs !== baseline.ctimeNs,
          )
        ) {
          throw new ConsultChimpsError(
            "FILES_SOURCE_CHANGED",
            "The input file changed while it was being read. Retry with a stable copy.",
          );
        }
      },
      close: () => handle.close(),
    };
  } catch (error) {
    await handle.close();
    throw error;
  }
}

export async function createScratchDirectory(
  parentPath: string,
): Promise<ScratchDirectory> {
  const directory = await mkdtemp(
    path.join(path.resolve(parentPath), "cc-scratch-"),
  );
  interface ScratchHandle {
    readonly handle: FileHandle;
    closeAttempt?: Promise<void> | undefined;
  }
  const handles = new Set<ScratchHandle>();
  const pendingCreates = new Set<Promise<void>>();
  let sequence = 0;
  let closeAttempt: Promise<void> | undefined;
  let closeRequested = false;
  let removed = false;

  const cleanupFailure = (causes: readonly unknown[]) =>
    new ConsultChimpsError(
      "FILES_SCRATCH_CLEANUP_FAILED",
      "Temporary storage could not be fully removed. Resolve the file access problem, then call close again.",
      {
        cause:
          causes.length === 1
            ? causes[0]
            : new AggregateError(
                causes,
                "More than one temporary file could not be closed.",
              ),
        details: { directory },
      },
    );

  const closeHandle = (entry: ScratchHandle): Promise<void> => {
    if (entry.closeAttempt) return entry.closeAttempt;
    if (!handles.has(entry)) return Promise.resolve();
    const attempt = (async () => {
      await entry.handle.close();
      handles.delete(entry);
    })();
    entry.closeAttempt = attempt;
    void attempt.then(
      () => {
        if (entry.closeAttempt === attempt) entry.closeAttempt = undefined;
      },
      () => {
        if (entry.closeAttempt === attempt) entry.closeAttempt = undefined;
      },
    );
    return attempt;
  };

  const closeDirectory = (): Promise<void> => {
    closeRequested = true;
    if (closeAttempt) return closeAttempt;
    if (removed) return Promise.resolve();
    const attempt = (async () => {
      await Promise.all([...pendingCreates]);
      const results = await Promise.allSettled([...handles].map(closeHandle));
      const failures = results.flatMap((result) =>
        result.status === "rejected" ? [result.reason] : [],
      );
      if (failures.length > 0) throw cleanupFailure(failures);
      try {
        await rm(directory, { recursive: true, force: true });
        removed = true;
      } catch (cause) {
        throw cleanupFailure([cause]);
      }
    })();
    closeAttempt = attempt;
    void attempt.then(
      () => {
        if (closeAttempt === attempt) closeAttempt = undefined;
      },
      () => {
        if (closeAttempt === attempt) closeAttempt = undefined;
      },
    );
    return attempt;
  };

  return {
    async create() {
      if (closeRequested) {
        throw new ConsultChimpsError(
          "FILES_SCRATCH_CLOSED",
          "Temporary storage is closed. Start a new operation before writing more data.",
        );
      }
      const name = `part-${sequence++}`;
      let finishCreate!: () => void;
      const pendingCreate = new Promise<void>((resolve) => {
        finishCreate = resolve;
      });
      pendingCreates.add(pendingCreate);
      try {
        const handle = await open(path.join(directory, name), "wx+", 0o600);
        const entry: ScratchHandle = { handle };
        handles.add(entry);
        if (closeRequested) {
          throw new ConsultChimpsError(
            "FILES_SCRATCH_CLOSED",
            "Temporary storage began closing before the file was ready. Start a new operation.",
          );
        }
        let size = 0;
        return {
          name,
          get size() {
            return size;
          },
          readAt: (offset, length, signal) =>
            readRange(handle, size, offset, length, signal),
          async writeAt(offset, bytes) {
            assertRange(offset, bytes.length);
            let completed = 0;
            while (completed < bytes.length) {
              const result = await handle.write(
                bytes,
                completed,
                bytes.length - completed,
                offset + completed,
              );
              if (result.bytesWritten === 0) {
                throw new ConsultChimpsError(
                  "FILES_WRITE_FAILED",
                  "Temporary storage could not accept more data. Check available disk space.",
                );
              }
              completed += result.bytesWritten;
            }
            size = Math.max(size, offset + bytes.length);
          },
          async truncate(length) {
            assertRange(0, length);
            await handle.truncate(length);
            size = length;
          },
          close: () => closeHandle(entry),
        };
      } finally {
        pendingCreates.delete(pendingCreate);
        finishCreate();
      }
    },
    close: closeDirectory,
  };
}
