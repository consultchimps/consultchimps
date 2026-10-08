/**
 * L3: a split output staged on disk as it is produced, written in pieces of up
 * to 1 MiB. A write may take less than it was given; the rest follows, and a
 * write that takes nothing fails the run rather than leaving a short file.
 */
import { open } from "node:fs/promises";

export interface StagedFile {
  write(chunk: Uint8Array): Promise<void>;
  /** Write what is pending and close. */
  close(): Promise<void>;
  /** Close without writing what is pending. */
  abort(): Promise<void>;
}

const PIECE = 1024 * 1024;

/** What `stagedFile` needs of a file handle. */
export interface StagedFileHandle {
  write(
    buffer: Uint8Array,
    offset: number,
    length: number,
  ): Promise<{ bytesWritten: number }>;
  close(): Promise<void>;
}

export async function stagedFile(
  filePath: string,
  openFile: (filePath: string) => Promise<StagedFileHandle> = (path) =>
    open(path, "wx"),
): Promise<StagedFile> {
  const handle = await openFile(filePath);
  let pending: Uint8Array[] = [];
  let pendingSize = 0;
  const flush = async (): Promise<void> => {
    if (pendingSize === 0) return;
    const joined = new Uint8Array(pendingSize);
    let offset = 0;
    for (const chunk of pending) {
      joined.set(chunk, offset);
      offset += chunk.length;
    }
    pending = [];
    pendingSize = 0;
    for (let written = 0; written < joined.length;) {
      const { bytesWritten } = await handle.write(
        joined,
        written,
        joined.length - written,
      );
      if (bytesWritten === 0) {
        throw new Error(
          `A split output could not be written to ${filePath}: the disk accepted no more data.`,
        );
      }
      written += bytesWritten;
    }
  };
  return {
    async write(chunk) {
      pending.push(chunk);
      pendingSize += chunk.length;
      if (pendingSize >= PIECE) await flush();
    },
    async close() {
      try {
        await flush();
      } finally {
        await handle.close();
      }
    },
    async abort() {
      await handle.close();
    },
  };
}
