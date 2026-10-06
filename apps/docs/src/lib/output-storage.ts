/**
 * Where the operation worker writes a streamed output (ADR 0006, browser
 * streaming): a file in the Origin Private File System, written through a sync
 * access handle, so the output sits on disk rather than in the tab. The page
 * downloads the OPFS `File` it ends as. Where OPFS is missing or refuses, the
 * output is collected as `Blob` parts instead and marked as held in memory.
 *
 * While a worker offers an output it holds a Web Lock named for it, which the
 * browser releases when the tab closes. A sweep deletes every output whose lock
 * is free, so a closed tab's output goes the next time any tool page starts its
 * worker, and an output another open tab still offers is left alone.
 *
 * Runtime-neutral apart from the storage and locks it is handed, so it is
 * tested with stand-ins for both.
 */
import { ConsultChimpsError } from "@consultchimps/core";
import type { ByteSink } from "@consultchimps/xlsx/bytes";

/** The code a run fails with when the browser stops accepting its output. */
export const OUTPUT_STORAGE_FULL = "OUTPUT_STORAGE_FULL";

export interface OutputSyncAccessHandle {
  write(buffer: Uint8Array, options: { readonly at: number }): number;
  truncate(size: number): void;
  // Safari 15.2 to 16 return promises from these two.
  flush(): void | Promise<void>;
  close(): void | Promise<void>;
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

/** The part of the Web Locks API used to mark an output as still offered. */
export interface OutputLocks {
  request(
    name: string,
    options: { readonly mode: "exclusive"; readonly ifAvailable?: boolean },
    callback: (lock: unknown) => Promise<void>,
  ): Promise<unknown>;
}

/** Where outputs go, and the locks that mark the ones still offered. */
export interface OutputPlace {
  readonly storage: OutputStorage | undefined;
  readonly locks: OutputLocks | undefined;
}

/** The folder in the origin's private file system that holds outputs. */
export const OUTPUT_DIRECTORY = "consultchimps-outputs";

/**
 * In a browser without Web Locks, ownership cannot be seen, so an output is
 * swept only once it is this old.
 */
export const OUTPUT_RETENTION_MS = 24 * 60 * 60 * 1000;

/** A streamed output being written, and what it becomes once finished. */
export interface OutputTarget {
  readonly sink: ByteSink;
  /** Close the file and hand back what was written. */
  finish(): Promise<{ readonly blob: Blob; readonly inMemory: boolean }>;
}

const lockName = (name: string): string => `consultchimps-output:${name}`;

/** Lock releases for the outputs this worker holds, by file name. */
const held = new Map<string, () => void>();

/** Take an output's lock and keep it until `release` is called. */
function hold(locks: OutputLocks | undefined, name: string): Promise<void> {
  if (locks === undefined) return Promise.resolve();
  return new Promise((granted, refused) => {
    locks
      .request(
        lockName(name),
        { mode: "exclusive" },
        () =>
          new Promise<void>((release) => {
            held.set(name, release);
            granted();
          }),
      )
      .catch(refused);
  });
}

function release(name: string): void {
  held.get(name)?.();
  held.delete(name);
}

/** This origin's private file system and locks, where the browser has them. */
export function browserOutputPlace(): OutputPlace {
  const storage: unknown =
    typeof navigator === "undefined" ? undefined : navigator.storage;
  const locks: unknown =
    typeof navigator === "undefined" ? undefined : navigator.locks;
  return {
    storage:
      typeof storage === "object" &&
      storage !== null &&
      "getDirectory" in storage &&
      typeof storage.getDirectory === "function"
        ? (storage as OutputStorage)
        : undefined,
    locks:
      typeof locks === "object" &&
      locks !== null &&
      "request" in locks &&
      typeof locks.request === "function"
        ? (locks as OutputLocks)
        : undefined,
  };
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

/** Names carry their creation time, for a sweep in a browser without locks. */
function outputFileName(now: number): string {
  // getRandomValues rather than randomUUID, which Safari lacks before 15.4.
  const random = crypto.getRandomValues(new Uint8Array(16));
  const id = Array.from(random, (byte) =>
    byte.toString(16).padStart(2, "0"),
  ).join("");
  return `${now}-${id}.part`;
}

function createdAt(name: string): number | undefined {
  const match = /^(\d+)-/u.exec(name);
  return match ? Number(match[1]) : undefined;
}

async function outputDirectory(
  storage: OutputStorage,
): Promise<OutputDirectoryHandle> {
  return (await storage.getDirectory()).getDirectoryHandle(OUTPUT_DIRECTORY, {
    create: true,
  });
}

/**
 * Open a new output. `created` collects the names written, so the worker can
 * delete them once the page no longer offers them.
 */
export async function openOutputTarget(
  place: OutputPlace,
  mediaType: string,
  created: Set<string>,
  now: number = Date.now(),
): Promise<OutputTarget> {
  const { storage, locks } = place;
  if (storage === undefined) return memoryTarget(mediaType);
  let directory: OutputDirectoryHandle;
  let handle: OutputFileHandle;
  let access: OutputSyncAccessHandle;
  const name = outputFileName(now);
  try {
    directory = await outputDirectory(storage);
    // Held before the file exists, so no sweep can take it while it is new.
    await hold(locks, name);
    handle = await directory.getFileHandle(name, { create: true });
  } catch {
    release(name);
    // A private window, or a browser without OPFS in workers.
    return memoryTarget(mediaType);
  }
  created.add(name);
  try {
    // Safari 15.2 and 15.3 have sync access handles but no getFile, which
    // the download needs, so they use memory from the start.
    if (typeof handle.getFile !== "function") {
      throw new TypeError("This browser cannot read an OPFS file back");
    }
    access = await handle.createSyncAccessHandle();
  } catch {
    await removeOutput(directory, name, created);
    return memoryTarget(mediaType);
  }

  let at = 0;
  let open = true;
  const close = async (): Promise<void> => {
    if (open) {
      open = false;
      await access.close();
    }
  };
  return {
    sink: {
      write: (chunk) => {
        let rest = chunk;
        while (rest.length > 0) {
          let written: number;
          try {
            written = access.write(rest, { at });
          } catch (error) {
            if (
              error instanceof DOMException &&
              error.name === "QuotaExceededError"
            ) {
              throw storageFull(at, error);
            }
            throw error;
          }
          // No progress means the disk is full or the handle is gone; going
          // round again would never end.
          if (written <= 0) throw storageFull(at, undefined);
          at += written;
          rest = rest.subarray(written);
        }
      },
      flush: () => Promise.resolve(),
      abort: async () => {
        try {
          await close();
        } finally {
          await removeOutput(directory, name, created);
        }
      },
    },
    finish: async () => {
      try {
        await access.flush();
      } finally {
        await close();
      }
      // The OPFS file has no type; the page relabels it when it downloads.
      return { blob: await handle.getFile(), inMemory: false };
    },
  };
}

function storageFull(bytesWritten: number, cause: unknown): ConsultChimpsError {
  return new ConsultChimpsError(
    OUTPUT_STORAGE_FULL,
    "The browser stopped accepting the output, which usually means the storage it gives this site is full. Free some disk space or clear this site's data in the browser settings, then run the task again",
    { cause, details: { bytesWritten } },
  );
}

async function removeOutput(
  directory: OutputDirectoryHandle,
  name: string,
  created: Set<string>,
): Promise<void> {
  try {
    await directory.removeEntry(name);
  } catch (error) {
    // Still open somewhere: keep the name so the next removal retries.
    if (!(error instanceof DOMException && error.name === "NotFoundError")) {
      return;
    }
  }
  created.delete(name);
  release(name);
}

/**
 * Delete outputs: the ones named in `created`, which this worker wrote and the
 * page no longer offers, and any whose lock no tab holds, which a closed tab
 * left behind. Without locks, only outputs older than the retention period
 * count as left behind. Failures are left for the next sweep.
 */
export async function removeOutputs(
  place: OutputPlace,
  created: Set<string>,
  now: number = Date.now(),
): Promise<void> {
  const { storage, locks } = place;
  if (storage === undefined) return;
  let directory: OutputDirectoryHandle;
  try {
    directory = await outputDirectory(storage);
  } catch {
    return;
  }
  for (const name of [...created]) {
    await removeOutput(directory, name, created);
  }
  const others: string[] = [];
  try {
    for await (const entry of directory.values()) {
      if (!held.has(entry.name)) others.push(entry.name);
    }
  } catch {
    return;
  }
  for (const name of others) {
    if (locks === undefined) {
      const time = createdAt(name);
      if (time !== undefined && now - time > OUTPUT_RETENTION_MS) {
        await removeOutput(directory, name, new Set());
      }
      continue;
    }
    await locks.request(
      lockName(name),
      { mode: "exclusive", ifAvailable: true },
      async (lock) => {
        // Another open tab still offers it.
        if (lock === null) return;
        await removeOutput(directory, name, new Set());
      },
    );
  }
}
