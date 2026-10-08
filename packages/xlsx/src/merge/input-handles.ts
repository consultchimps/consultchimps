/**
 * The merge's inputs, read on the command line through a bounded number of
 * open file handles. Merging reads every input's rows again when the output is
 * written, so a handle per input would hold one descriptor per workbook. An
 * input whose handle was closed is opened again when next read, and refused
 * if it is no longer the file first read.
 */
import { stat } from "node:fs/promises";

import {
  ConsultChimpsError,
  type RandomAccessSource,
} from "@consultchimps/core";
import { openRandomAccessSource, type FileSource } from "@consultchimps/files";

/** How many inputs stay open at once. */
export const OPEN_INPUT_LIMIT = 16;

export interface MergeInput extends RandomAccessSource {
  /** Refuse the input if it changed since it was first opened. */
  verifyUnchanged(): Promise<void>;
}

type Identity = Awaited<ReturnType<typeof identityOf>>;

function identityOf(filePath: string) {
  return stat(filePath, { bigint: true });
}

function sameFile(first: Identity, now: Identity): boolean {
  return (
    first.dev === now.dev &&
    first.ino === now.ino &&
    first.size === now.size &&
    first.mtimeNs === now.mtimeNs &&
    first.ctimeNs === now.ctimeNs
  );
}

function changed(cause?: unknown): ConsultChimpsError {
  return new ConsultChimpsError(
    "FILES_SOURCE_CHANGED",
    "The input file changed while it was being read. Retry with a stable copy.",
    cause === undefined ? undefined : { cause },
  );
}

interface Entry {
  readonly filePath: string;
  readonly first: Identity;
  /** Reads under way, during which the handle is not closed. */
  reading: number;
  /** The handle being opened again, which every read waits for. */
  reopening: Promise<FileSource> | undefined;
}

export class MergeInputs {
  readonly #limit: number;
  /** Open handles, least recently read first. */
  readonly #open = new Map<Entry, FileSource>();

  constructor(limit: number = OPEN_INPUT_LIMIT) {
    this.#limit = limit;
  }

  /** Open an input; it stays open until others push it out. */
  async open(filePath: string): Promise<MergeInput> {
    const source = await openRandomAccessSource(filePath);
    let first: Identity;
    try {
      first = await identityOf(filePath);
      // The path still names the file the handle opened.
      await source.verifyUnchanged();
    } catch (error) {
      await source.close().catch(() => undefined);
      throw error instanceof ConsultChimpsError ? error : changed(error);
    }
    const entry: Entry = {
      filePath,
      first,
      reading: 0,
      reopening: undefined,
    };
    await this.#admit(entry, source);
    return {
      name: source.name,
      size: source.size,
      readAt: async (offset, length, signal) => {
        entry.reading += 1;
        try {
          const handle = await this.#handle(entry);
          return await handle.readAt(offset, length, signal);
        } finally {
          entry.reading -= 1;
        }
      },
      verifyUnchanged: async () => {
        await this.#open.get(entry)?.verifyUnchanged();
        let now: Identity;
        try {
          now = await identityOf(entry.filePath);
        } catch (error) {
          throw changed(error);
        }
        if (!sameFile(entry.first, now)) throw changed();
      },
    };
  }

  async close(): Promise<void> {
    const sources = [...this.#open.values()];
    this.#open.clear();
    for (const source of sources) await source.close().catch(() => undefined);
  }

  async #handle(entry: Entry): Promise<FileSource> {
    const held = this.#open.get(entry);
    if (held) {
      this.#open.delete(entry);
      this.#open.set(entry, held);
      return held;
    }
    entry.reopening ??= this.#reopen(entry).finally(() => {
      entry.reopening = undefined;
    });
    return entry.reopening;
  }

  async #reopen(entry: Entry): Promise<FileSource> {
    let source: FileSource;
    try {
      source = await openRandomAccessSource(entry.filePath);
    } catch (error) {
      throw changed(error);
    }
    try {
      const now = await identityOf(entry.filePath);
      if (!sameFile(entry.first, now)) throw changed();
    } catch (error) {
      await source.close().catch(() => undefined);
      throw error instanceof ConsultChimpsError ? error : changed(error);
    }
    await this.#admit(entry, source);
    return source;
  }

  async #admit(entry: Entry, source: FileSource): Promise<void> {
    this.#open.set(entry, source);
    for (const [other, open] of this.#open) {
      if (this.#open.size <= this.#limit) break;
      if (other === entry || other.reading > 0) continue;
      this.#open.delete(other);
      await open.close().catch(() => undefined);
    }
  }
}
