/**
 * Where the operation worker writes a streamed output (ADR 0006, browser
 * streaming): a file in the Origin Private File System, written through a sync
 * access handle, so the output sits on disk rather than in the tab. The page
 * downloads the OPFS `File` it ends as. Where OPFS is missing or refuses, the
 * output is collected as `Blob` parts instead and marked as held in memory.
 *
 * Runtime-neutral apart from the storage it is handed, so it is tested with a
 * stand-in for OPFS.
 */
import type { ByteSink } from "@consultchimps/xlsx/bytes";

export interface OutputSyncAccessHandle {
  write(buffer: Uint8Array, options: { readonly at: number }): number;
  truncate(size: number): void;
  flush(): void;
  close(): void;
}

export interface OutputFileHandle {
  readonly kind: "file";
  readonly name: string;
  getFile(): Promise<File>;
  createSyncAccessHandle(): Promise<OutputSyncAccessHandle>;
}

export interface OutputDirectoryHandle {
  readonly kind: "directory";
  readonly name: string;
  values(): AsyncIterable<OutputDirectoryHandle | OutputFileHandle>;
  getDirectoryHandle(
    name: string,
    options?: { readonly create?: boolean },
  ): Promise<OutputDirectoryHandle>;
  getFileHandle(
    name: string,
    options?: { readonly create?: boolean },
  ): Promise<OutputFileHandle>;
  removeEntry(name: string): Promise<void>;
}

export interface OutputStorage {
  getDirectory(): Promise<OutputDirectoryHandle>;
}

/** The folder in the origin's private file system that holds outputs. */
export const OUTPUT_DIRECTORY = "consultchimps-outputs";

/** How long an output left by a closed tab is kept before a sweep removes it. */
export const OUTPUT_RETENTION_MS = 24 * 60 * 60 * 1000;

/** A streamed output being written, and what it becomes once finished. */
export interface OutputTarget {
  readonly sink: ByteSink;
  /** Close the file and hand back what was written. */
  finish(): Promise<{ readonly blob: Blob; readonly inMemory: boolean }>;
}

/** This origin's private file system, or undefined where there is none. */
export function browserOutputStorage(): OutputStorage | undefined {
  const storage: unknown =
    typeof navigator === "undefined" ? undefined : navigator.storage;
  if (
    typeof storage !== "object" ||
    storage === null ||
    !("getDirectory" in storage) ||
    typeof storage.getDirectory !== "function"
  ) {
    return undefined;
  }
  return storage as OutputStorage;
}

function memoryTarget(mediaType: string): OutputTarget {
  let parts: Uint8Array<ArrayBuffer>[] = [];
  return {
    sink: {
      write: (chunk) => {
        parts.push(chunk as Uint8Array<ArrayBuffer>);
      },
      flush: () => Promise.resolve(),
      abort: () => {
        parts = [];
        return Promise.resolve();
      },
    },
    finish: () => {
      const blob = new Blob(parts, { type: mediaType });
      parts = [];
      return Promise.resolve({ blob, inMemory: true });
    },
  };
}

/**
 * Names carry their creation time, so a sweep can tell an old output from one
 * another tab is still offering, without opening either.
 */
function outputFileName(now: number): string {
  return `${now}-${crypto.randomUUID()}.part`;
}

function createdAt(name: string): number | undefined {
  const match = /^(\d+)-/u.exec(name);
  return match ? Number(match[1]) : undefined;
}

/**
 * Open a new output. `created` collects the names written, so the worker can
 * delete them once the page no longer offers them.
 */
export async function openOutputTarget(
  storage: OutputStorage | undefined,
  mediaType: string,
  created: Set<string>,
  now: number = Date.now(),
): Promise<OutputTarget> {
  if (storage === undefined) return memoryTarget(mediaType);
  let directory: OutputDirectoryHandle;
  let handle: OutputFileHandle;
  let access: OutputSyncAccessHandle;
  const name = outputFileName(now);
  try {
    directory = await (
      await storage.getDirectory()
    ).getDirectoryHandle(OUTPUT_DIRECTORY, { create: true });
    handle = await directory.getFileHandle(name, { create: true });
  } catch {
    // A private window, or a browser without OPFS in workers.
    return memoryTarget(mediaType);
  }
  created.add(name);
  try {
    access = await handle.createSyncAccessHandle();
  } catch {
    await removeOutput(directory, name, created);
    return memoryTarget(mediaType);
  }

  let at = 0;
  let open = true;
  const close = (): void => {
    if (open) {
      open = false;
      access.close();
    }
  };
  return {
    sink: {
      write: (chunk) => {
        let rest = chunk;
        while (rest.length > 0) {
          const written = access.write(rest, { at });
          // No progress means the disk is full or the handle is gone; going
          // round again would never end.
          if (written <= 0) {
            throw new Error(
              "The browser stopped accepting the output, which usually means its storage is full",
            );
          }
          at += written;
          rest = rest.subarray(written);
        }
      },
      flush: () => Promise.resolve(),
      abort: async () => {
        try {
          close();
        } finally {
          await removeOutput(directory, name, created);
        }
      },
    },
    finish: async () => {
      try {
        access.flush();
      } finally {
        close();
      }
      // The OPFS file has no type; the page relabels it when it downloads.
      return { blob: await handle.getFile(), inMemory: false };
    },
  };
}

async function removeOutput(
  directory: OutputDirectoryHandle,
  name: string,
  created: Set<string>,
): Promise<void> {
  created.delete(name);
  try {
    await directory.removeEntry(name);
  } catch {
    // Already gone, or still open somewhere; a later sweep retries.
  }
}

/**
 * Delete outputs: the ones named in `created`, which this worker wrote and the
 * page no longer offers, and any older than the retention period, which a
 * closed tab left behind. Failures are left for the next sweep.
 */
export async function removeOutputs(
  storage: OutputStorage | undefined,
  created: Set<string>,
  now: number = Date.now(),
): Promise<void> {
  if (storage === undefined) return;
  let directory: OutputDirectoryHandle;
  try {
    directory = await (
      await storage.getDirectory()
    ).getDirectoryHandle(OUTPUT_DIRECTORY, { create: true });
  } catch {
    return;
  }
  const names = new Set(created);
  try {
    for await (const entry of directory.values()) {
      const time = createdAt(entry.name);
      if (time !== undefined && now - time > OUTPUT_RETENTION_MS) {
        names.add(entry.name);
      }
    }
  } catch {
    // Listing failed; the outputs this worker knows of can still go.
  }
  for (const name of names) {
    await removeOutput(directory, name, created);
  }
}
