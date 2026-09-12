import type { RandomAccessFile, RandomAccessSource } from "@consultchimps/core";

interface OpfsFileHandle {
  readonly name: string;
  getFile(): Promise<File>;
  createSyncAccessHandle(): Promise<OpfsSyncAccessHandle>;
}

interface OpfsSyncAccessHandle {
  read(buffer: Uint8Array, options: { readonly at: number }): number;
  write(buffer: Uint8Array, options: { readonly at: number }): number;
  truncate(size: number): void;
  flush(): void;
  close(): void;
}

interface OpfsDirectoryHandle {
  getFileHandle(
    name: string,
    options?: { readonly create?: boolean },
  ): Promise<OpfsFileHandle>;
  removeEntry(name: string): Promise<void>;
}

interface OpfsStorageManager {
  getDirectory(): Promise<OpfsDirectoryHandle>;
}

function storageManager(): OpfsStorageManager {
  const storage: unknown = navigator.storage;
  if (
    typeof storage !== "object" ||
    storage === null ||
    !("getDirectory" in storage) ||
    typeof storage.getDirectory !== "function"
  ) {
    throw new Error(
      "This browser does not provide origin-private file storage. Use a current Chromium browser and try again.",
    );
  }
  return storage as OpfsStorageManager;
}

export class BrowserBlobSource implements RandomAccessSource {
  readonly name: string;
  readonly size: number;
  readonly #blob: Blob;

  constructor(name: string, blob: Blob) {
    this.name = name;
    this.size = blob.size;
    this.#blob = blob;
  }

  async readAt(
    offset: number,
    length: number,
    signal?: AbortSignal,
  ): Promise<Uint8Array> {
    signal?.throwIfAborted();
    if (
      !Number.isSafeInteger(offset) ||
      !Number.isSafeInteger(length) ||
      offset < 0 ||
      length < 0 ||
      offset + length > this.size
    ) {
      throw new RangeError("The requested file range is outside the source");
    }
    const bytes = new Uint8Array(
      await this.#blob.slice(offset, offset + length).arrayBuffer(),
    );
    signal?.throwIfAborted();
    return bytes;
  }
}

export class BrowserOpfsFile implements RandomAccessFile {
  readonly name: string;
  readonly #handle: OpfsFileHandle;
  readonly #access: OpfsSyncAccessHandle;
  readonly #removeOnClose: boolean;
  #size: number;
  #closed = false;

  private constructor(
    name: string,
    handle: OpfsFileHandle,
    access: OpfsSyncAccessHandle,
    size: number,
    removeOnClose: boolean,
  ) {
    this.name = name;
    this.#handle = handle;
    this.#access = access;
    this.#size = size;
    this.#removeOnClose = removeOnClose;
  }

  static async open(
    name: string,
    create = false,
    removeOnClose = false,
  ): Promise<BrowserOpfsFile> {
    const root = await storageManager().getDirectory();
    const handle = await root.getFileHandle(name, { create });
    const access = await handle.createSyncAccessHandle();
    return new BrowserOpfsFile(
      name,
      handle,
      access,
      (await handle.getFile()).size,
      removeOnClose,
    );
  }

  get size(): number {
    return this.#size;
  }

  async readAt(
    offset: number,
    length: number,
    signal?: AbortSignal,
  ): Promise<Uint8Array> {
    signal?.throwIfAborted();
    if (offset < 0 || length < 0 || offset + length > this.#size) {
      throw new RangeError(
        "The requested file range is outside the working file",
      );
    }
    const bytes = new Uint8Array(length);
    const read = this.#access.read(bytes, { at: offset });
    signal?.throwIfAborted();
    return read === length ? bytes : bytes.subarray(0, read);
  }

  async writeAt(offset: number, bytes: Uint8Array): Promise<void> {
    const written = this.#access.write(bytes, { at: offset });
    if (written !== bytes.byteLength) {
      throw new Error("The browser scratch file accepted only part of a write");
    }
    this.#size = Math.max(this.#size, offset + bytes.byteLength);
  }

  async truncate(size: number): Promise<void> {
    this.#access.truncate(size);
    this.#size = size;
  }

  async file(): Promise<File> {
    this.#access.flush();
    return this.#handle.getFile();
  }

  async close(): Promise<void> {
    if (this.#closed) return;
    this.#closed = true;
    this.#access.flush();
    this.#access.close();
    if (!this.#removeOnClose) return;
    try {
      await removeOpfsFile(this.name);
    } catch (error) {
      if (!(error instanceof DOMException && error.name === "NotFoundError")) {
        throw error;
      }
    }
  }
}

export async function removeOpfsFile(name: string): Promise<void> {
  const root = await storageManager().getDirectory();
  await root.removeEntry(name);
}

export const browserScratchFactory = {
  async create(): Promise<RandomAccessFile> {
    return BrowserOpfsFile.open(
      `.consultchimps-scratch-${globalThis.crypto.randomUUID()}`,
      true,
      true,
    );
  },
};
