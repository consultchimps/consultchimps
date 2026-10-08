/**
 * L0: a package written part by part with exactly the bytes JSZip 3.10 writes
 * for `WorkbookPackage.save` (ADR 0006).
 *
 * The operations that edit a workbook in place, such as the workbook-keeping
 * split, have always produced their outputs through JSZip. Writing the same
 * bytes from a stream keeps those outputs identical while no part is held
 * whole: each part is compressed as it is produced, with pako 1.0, which JSZip
 * compresses with, at its default level. Deflate's output does not depend on
 * how its input is divided, so pieces of any size give JSZip's bytes.
 *
 * JSZip writes each part's sizes into its local header, so a part's compressed
 * bytes are held until the part ends: the working set is one compressed part.
 */
import pako from "pako";

import { updateCrc } from "./zip-reader.js";

/** Receives the package's bytes in order. */
export type ZipSink = (chunk: Uint8Array) => void | Promise<void>;

const encoder = new TextEncoder();

/** JSZip's CRC of a binary string, for the Unicode path extra field. */
function crcOf(bytes: Uint8Array): number {
  return (updateCrc(-1, bytes) ^ -1) >>> 0;
}

/** JSZip's DOS time and date for a part's `Date`, read in UTC. */
function dosFields(date: Date): { time: number; date: number } {
  const time =
    (((date.getUTCHours() << 6) | date.getUTCMinutes()) << 5) |
    (date.getUTCSeconds() / 2);
  const day =
    ((((date.getUTCFullYear() - 1980) << 4) | (date.getUTCMonth() + 1)) << 5) |
    date.getUTCDate();
  return { time: time & 0xffff, date: day & 0xffff };
}

class ByteWriter {
  readonly bytes: Uint8Array;
  readonly #view: DataView;
  #offset = 0;

  constructor(length: number) {
    this.bytes = new Uint8Array(length);
    this.#view = new DataView(this.bytes.buffer);
  }

  u16(value: number): void {
    this.#view.setUint16(this.#offset, value, true);
    this.#offset += 2;
  }

  u32(value: number): void {
    this.#view.setUint32(this.#offset, value >>> 0, true);
    this.#offset += 4;
  }

  raw(bytes: Uint8Array): void {
    this.bytes.set(bytes, this.#offset);
    this.#offset += bytes.length;
  }
}

interface WrittenEntry {
  readonly name: Uint8Array;
  readonly extra: Uint8Array;
  readonly flags: number;
  readonly time: number;
  readonly date: number;
  readonly crc: number;
  readonly compressedSize: number;
  readonly size: number;
  readonly offset: number;
}

/** One part being compressed. */
export interface ZipPartWriter {
  /** Add the next piece of the part's uncompressed bytes. */
  push(bytes: Uint8Array): void;
  /** Add text, encoded as UTF-8. */
  text(text: string): void;
  /**
   * End the part and write it, with `date` in place of the time it was
   * started with when that is only settled once its content is.
   */
  close(date?: Date): Promise<void>;
}

/**
 * A package written to `sink` in JSZip's layout: every part deflated, DOS
 * platform, no data descriptors, a part name with characters outside ASCII
 * flagged UTF-8 and given a Unicode path field, as JSZip does.
 */
export class JsZipWriter {
  readonly #sink: ZipSink;
  readonly #entries: WrittenEntry[] = [];
  #offset = 0;
  #open = false;

  constructor(sink: ZipSink) {
    this.#sink = sink;
  }

  async #emit(bytes: Uint8Array): Promise<void> {
    this.#offset += bytes.length;
    await this.#sink(bytes);
  }

  /** Start a part; parts are written one at a time, in order. */
  part(name: string, date: Date): ZipPartWriter {
    if (this.#open) {
      throw new Error("The previous package part is still being written.");
    }
    this.#open = true;
    const deflate = new pako.Deflate({ raw: true, level: -1 });
    const compressed: Uint8Array[] = [];
    let compressedSize = 0;
    deflate.onData = (chunk) => {
      compressed.push(chunk);
      compressedSize += chunk.length;
    };
    let crc = -1;
    let size = 0;
    const push = (bytes: Uint8Array): void => {
      if (bytes.length === 0) return;
      crc = updateCrc(crc, bytes);
      size += bytes.length;
      deflate.push(bytes, false);
    };
    return {
      push,
      text: (text) => {
        push(encoder.encode(text));
      },
      close: async (settled?: Date) => {
        deflate.push([], true);
        if (deflate.err !== 0) {
          throw new Error(`Could not compress ${name}: ${deflate.msg}`);
        }
        await this.#writePart(
          name,
          settled ?? date,
          (crc ^ -1) >>> 0,
          size,
          compressedSize,
          compressed,
        );
        this.#open = false;
      },
    };
  }

  async #writePart(
    name: string,
    date: Date,
    crc: number,
    size: number,
    compressedSize: number,
    compressed: readonly Uint8Array[],
  ): Promise<void> {
    const encodedName = encoder.encode(name);
    const utf8 = encodedName.length !== name.length;
    let extra: Uint8Array = new Uint8Array(0);
    if (utf8) {
      // Info-ZIP Unicode Path: version 1, the CRC of the name as written,
      // and the name again.
      const field = new ByteWriter(4 + 5 + encodedName.length);
      field.u16(0x7075);
      field.u16(5 + encodedName.length);
      field.raw(new Uint8Array([1]));
      field.u32(crcOf(encodedName));
      field.raw(encodedName);
      extra = field.bytes;
    }
    const { time, date: day } = dosFields(date);
    const entry: WrittenEntry = {
      name: encodedName,
      extra,
      flags: utf8 ? 0x0800 : 0,
      time,
      date: day,
      crc,
      compressedSize,
      size,
      offset: this.#offset,
    };
    const header = new ByteWriter(30 + encodedName.length + extra.length);
    header.u32(0x04034b50);
    writeCommonHeader(header, entry);
    header.raw(encodedName);
    header.raw(extra);
    await this.#emit(header.bytes);
    for (const chunk of compressed) await this.#emit(chunk);
    this.#entries.push(entry);
  }

  /** Write the central directory and its end record. */
  async finish(): Promise<void> {
    if (this.#open) {
      throw new Error("A package part is still being written.");
    }
    const start = this.#offset;
    for (const entry of this.#entries) {
      const record = new ByteWriter(
        46 + entry.name.length + entry.extra.length,
      );
      record.u32(0x02014b50);
      record.u16(0x0014);
      writeCommonHeader(record, entry);
      record.u16(0); // comment length
      record.u16(0); // disk number start
      record.u16(0); // internal attributes
      record.u32(0); // external attributes
      record.u32(entry.offset);
      record.raw(entry.name);
      record.raw(entry.extra);
      await this.#emit(record.bytes);
    }
    const end = new ByteWriter(22);
    end.u32(0x06054b50);
    end.u16(0);
    end.u16(0);
    end.u16(this.#entries.length);
    end.u16(this.#entries.length);
    end.u32(this.#offset - start);
    end.u32(start);
    end.u16(0);
    await this.#emit(end.bytes);
  }
}

/** The fields the local header and the directory record share. */
function writeCommonHeader(writer: ByteWriter, entry: WrittenEntry): void {
  writer.u16(0x000a);
  writer.u16(entry.flags);
  writer.u16(8);
  writer.u16(entry.time);
  writer.u16(entry.date);
  writer.u32(entry.crc);
  writer.u32(entry.compressedSize);
  writer.u32(entry.size);
  writer.u16(entry.name.length);
  writer.u16(entry.extra.length);
}
