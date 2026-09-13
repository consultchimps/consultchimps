import {
  throwIfAborted,
  type ProgressReporter,
  type RandomAccessFile,
  type RandomAccessSource,
} from "@consultchimps/core";

import { databaseError } from "./errors.js";

const COPY_CHUNK_BYTES = 1024 * 1024;

export interface BrowserWritable {
  write(data: {
    readonly type: "write";
    readonly position: number;
    readonly data: Uint8Array;
  }): Promise<void>;
  truncate(size: number): Promise<void>;
  close(): Promise<void>;
  abort?(reason?: unknown): Promise<void>;
}

export interface BrowserFileHandle {
  readonly kind: "file";
  readonly name: string;
  getFile(): Promise<Blob>;
  createWritable(): Promise<BrowserWritable>;
}

export interface BrowserReadableFileHandle {
  getFile(): Promise<{
    readonly size: number;
    slice(start?: number, end?: number): Blob;
  }>;
}

type BrowserCopyOperation = "db.browser.export" | "db.browser.import";

async function readExact(
  source: RandomAccessSource,
  offset: number,
  length: number,
  signal: AbortSignal | undefined,
): Promise<Uint8Array> {
  const bytes = await source.readAt(offset, length, signal);
  if (bytes.length !== length) {
    throw databaseError(
      "DB_SOURCE_SHORT_READ",
      "The database source returned a different number of bytes than requested.",
      { offset, expected: length, actual: bytes.length },
    );
  }
  return bytes;
}

export async function copySourceToBrowserFile(options: {
  readonly source: RandomAccessSource;
  readonly destination: BrowserFileHandle;
  readonly operation: BrowserCopyOperation;
  readonly signal?: AbortSignal | undefined;
  readonly onProgress?: ProgressReporter | undefined;
}): Promise<void> {
  const writable = await options.destination.createWritable();
  let closed = false;
  try {
    for (
      let offset = 0;
      offset < options.source.size;
      offset += COPY_CHUNK_BYTES
    ) {
      throwIfAborted(options.signal, options.operation);
      const length = Math.min(COPY_CHUNK_BYTES, options.source.size - offset);
      const bytes = await readExact(
        options.source,
        offset,
        length,
        options.signal,
      );
      await writable.write({ type: "write", position: offset, data: bytes });
      options.onProgress?.({
        operation: options.operation,
        stage: "copying",
        completed: offset + bytes.length,
        total: options.source.size,
        detail: options.source.name,
      });
    }
    throwIfAborted(options.signal, options.operation);
    await writable.truncate(options.source.size);
    await writable.close();
    closed = true;
  } catch (error) {
    if (!closed && writable.abort !== undefined) {
      await writable.abort(error).catch(() => undefined);
    }
    throw error;
  }
}

export async function restoreBrowserFile(options: {
  readonly source: Blob;
  readonly destination: RandomAccessFile;
  readonly expectedSize: number;
}): Promise<void> {
  if (options.source.size !== options.expectedSize) {
    throw databaseError(
      "DB_BROWSER_EXPORT_BACKUP_INVALID",
      "The browser export backup has an unexpected size.",
      { expected: options.expectedSize, actual: options.source.size },
    );
  }
  await options.destination.truncate(0);
  for (
    let offset = 0;
    offset < options.source.size;
    offset += COPY_CHUNK_BYTES
  ) {
    const length = Math.min(COPY_CHUNK_BYTES, options.source.size - offset);
    const bytes = new Uint8Array(
      await options.source.slice(offset, offset + length).arrayBuffer(),
    );
    if (bytes.length !== length) {
      throw databaseError(
        "DB_BROWSER_EXPORT_BACKUP_INVALID",
        "The browser export backup returned a different number of bytes than requested.",
        { offset, expected: length, actual: bytes.length },
      );
    }
    await options.destination.writeAt(offset, bytes);
  }
  await options.destination.truncate(options.expectedSize);
}

export async function copyBrowserFileToDestination(options: {
  readonly source: BrowserReadableFileHandle;
  readonly destination: RandomAccessFile;
  readonly signal?: AbortSignal | undefined;
}): Promise<number> {
  const source = await options.source.getFile();
  throwIfAborted(options.signal, "db.browser.export");
  await options.destination.truncate(0);
  for (let offset = 0; offset < source.size; offset += COPY_CHUNK_BYTES) {
    throwIfAborted(options.signal, "db.browser.export");
    const length = Math.min(COPY_CHUNK_BYTES, source.size - offset);
    const bytes = new Uint8Array(
      await source.slice(offset, offset + length).arrayBuffer(),
    );
    if (bytes.length !== length) {
      throw databaseError(
        "DB_SOURCE_SHORT_READ",
        "The browser database returned a different number of bytes than requested.",
        { offset, expected: length, actual: bytes.length },
      );
    }
    await options.destination.writeAt(offset, bytes);
  }
  throwIfAborted(options.signal, "db.browser.export");
  await options.destination.truncate(source.size);
  return source.size;
}
