import { Inflate } from "fflate";

/**
 * Bytes that can be read at any offset: a file handle on the command line, a
 * byte array in the browser. A workbook is read from its zip central directory,
 * which sits at the end of the file, so reading needs random access (ADR 0006).
 */
export interface RandomAccessSource {
  readonly size: number;
  read(offset: number, length: number): Promise<Uint8Array>;
}

/** A source over bytes already in memory. */
export function bytesSource(bytes: Uint8Array): RandomAccessSource {
  return {
    size: bytes.length,
    read: (offset, length) =>
      Promise.resolve(bytes.subarray(offset, offset + length)),
  };
}

interface ZipEntry {
  readonly method: number;
  readonly compressedSize: number;
  readonly size: number;
  readonly localHeader: number;
}

const READ_CHUNK = 1024 * 1024;
const END_OF_DIRECTORY = 0x06054b50;
const ZIP64_LOCATOR = 0x07064b50;
const ZIP64_END = 0x06064b50;
const DIRECTORY_ENTRY = 0x02014b50;
const LOCAL_HEADER = 0x04034b50;
const STORED = 0;
const DEFLATED = 8;

const u16 = (b: Uint8Array, o: number): number => b[o]! | (b[o + 1]! << 8);
const u32 = (b: Uint8Array, o: number): number =>
  (b[o]! | (b[o + 1]! << 8) | (b[o + 2]! << 16) | (b[o + 3]! << 24)) >>> 0;
const u64 = (b: Uint8Array, o: number): number =>
  u32(b, o) + u32(b, o + 4) * 2 ** 32;

/**
 * A zip archive read entry by entry from its central directory. Entries are
 * inflated as a stream, one read chunk at a time, so a worksheet is never held
 * whole. Encrypted entries and compression methods other than stored and
 * deflate are refused, as is any entry whose data reaches past the source or
 * whose inflated length differs from the length the directory declares.
 */
export class ZipReader {
  readonly #source: RandomAccessSource;
  readonly #entries: ReadonlyMap<string, ZipEntry>;

  private constructor(
    source: RandomAccessSource,
    entries: ReadonlyMap<string, ZipEntry>,
  ) {
    this.#source = source;
    this.#entries = entries;
  }

  static async open(source: RandomAccessSource): Promise<ZipReader> {
    const tailLength = Math.min(source.size, 65_557);
    const tailStart = source.size - tailLength;
    const tail = await source.read(tailStart, tailLength);
    let end = -1;
    for (let index = tail.length - 22; index >= 0; index -= 1) {
      if (u32(tail, index) === END_OF_DIRECTORY) {
        end = index;
        break;
      }
    }
    if (end < 0) {
      throw new Error("The file is not a zip archive.");
    }
    let count = u16(tail, end + 10);
    let directorySize = u32(tail, end + 12);
    let directoryOffset = u32(tail, end + 16);
    if (
      count === 0xffff ||
      directorySize === 0xffffffff ||
      directoryOffset === 0xffffffff
    ) {
      const locatorOffset = tailStart + end - 20;
      const locator =
        locatorOffset >= 0 ? await source.read(locatorOffset, 20) : undefined;
      if (locator === undefined || u32(locator, 0) !== ZIP64_LOCATOR) {
        throw new Error("The zip64 end of central directory is missing.");
      }
      const record = await source.read(u64(locator, 8), 56);
      if (record.length < 56 || u32(record, 0) !== ZIP64_END) {
        throw new Error("The zip64 end of central directory is unreadable.");
      }
      count = u64(record, 32);
      directorySize = u64(record, 40);
      directoryOffset = u64(record, 48);
    }
    if (directoryOffset + directorySize > source.size) {
      throw new Error("The zip central directory lies outside the file.");
    }

    const directory = await source.read(directoryOffset, directorySize);
    const decoder = new TextDecoder();
    const entries = new Map<string, ZipEntry>();
    let offset = 0;
    for (let index = 0; index < count; index += 1) {
      if (
        offset + 46 > directory.length ||
        u32(directory, offset) !== DIRECTORY_ENTRY
      ) {
        throw new Error("The zip central directory is damaged.");
      }
      const flags = u16(directory, offset + 8);
      const method = u16(directory, offset + 10);
      let compressedSize = u32(directory, offset + 20);
      let size = u32(directory, offset + 24);
      const nameLength = u16(directory, offset + 28);
      const extraLength = u16(directory, offset + 30);
      const commentLength = u16(directory, offset + 32);
      let localHeader = u32(directory, offset + 42);
      const nameStart = offset + 46;
      const name = decoder.decode(
        directory.subarray(nameStart, nameStart + nameLength),
      );
      const extraEnd = nameStart + nameLength + extraLength;
      for (let field = nameStart + nameLength; field + 4 <= extraEnd;) {
        const id = u16(directory, field);
        const length = u16(directory, field + 2);
        let cursor = field + 4;
        if (id === 1) {
          if (size === 0xffffffff) {
            size = u64(directory, cursor);
            cursor += 8;
          }
          if (compressedSize === 0xffffffff) {
            compressedSize = u64(directory, cursor);
            cursor += 8;
          }
          if (localHeader === 0xffffffff) {
            localHeader = u64(directory, cursor);
          }
        }
        field += 4 + length;
      }
      offset = extraEnd + commentLength;
      if ((flags & 1) !== 0) {
        throw new Error(`The zip entry ${name} is encrypted.`);
      }
      if (name.endsWith("/")) {
        continue;
      }
      entries.set(name, { method, compressedSize, size, localHeader });
    }
    return new ZipReader(source, entries);
  }

  has(name: string): boolean {
    return this.#entries.has(name);
  }

  /** The whole inflated entry, for the small parts a reader holds anyway. */
  async readBytes(name: string): Promise<Uint8Array | undefined> {
    const entry = this.#entries.get(name);
    if (entry === undefined) {
      return undefined;
    }
    const bytes = new Uint8Array(entry.size);
    let offset = 0;
    await this.stream(name, (chunk) => {
      bytes.set(chunk, offset);
      offset += chunk.length;
    });
    return bytes;
  }

  async readText(name: string): Promise<string | undefined> {
    const bytes = await this.readBytes(name);
    return bytes === undefined ? undefined : new TextDecoder().decode(bytes);
  }

  /**
   * Inflate an entry into `onData`, one source read at a time. `between` runs
   * after each read, which is where a caller flushes output or yields.
   */
  async stream(
    name: string,
    onData: (chunk: Uint8Array) => void,
    between?: () => Promise<void>,
  ): Promise<void> {
    const entry = this.#entries.get(name);
    if (entry === undefined) {
      throw new Error(`The zip entry ${name} is missing.`);
    }
    if (entry.method !== STORED && entry.method !== DEFLATED) {
      throw new Error(
        `The zip entry ${name} uses compression method ${String(entry.method)}.`,
      );
    }
    const header = await this.#source.read(entry.localHeader, 30);
    if (header.length < 30 || u32(header, 0) !== LOCAL_HEADER) {
      throw new Error(`The zip entry ${name} has no local header.`);
    }
    let position = entry.localHeader + 30 + u16(header, 26) + u16(header, 28);
    const end = position + entry.compressedSize;
    if (end > this.#source.size) {
      throw new Error(
        `The zip entry ${name} reaches past the end of the file.`,
      );
    }

    let produced = 0;
    const accept = (chunk: Uint8Array): void => {
      produced += chunk.length;
      if (produced > entry.size) {
        throw new Error(`The zip entry ${name} is longer than declared.`);
      }
      if (chunk.length > 0) onData(chunk);
    };
    const inflate =
      entry.method === DEFLATED
        ? new Inflate((chunk) => {
            accept(chunk);
          })
        : undefined;

    if (position === end) {
      inflate?.push(new Uint8Array(0), true);
    }
    while (position < end) {
      const length = Math.min(READ_CHUNK, end - position);
      const chunk = await this.#source.read(position, length);
      position += length;
      if (inflate) inflate.push(chunk, position >= end);
      else accept(chunk);
      if (between) await between();
    }
    if (produced !== entry.size) {
      throw new Error(`The zip entry ${name} is shorter than declared.`);
    }
  }
}
