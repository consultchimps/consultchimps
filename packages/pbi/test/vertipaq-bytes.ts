/**
 * Crafted VertiPaq members for the hostile-input tests. Each builder writes the
 * layout the reader expects and takes the one field a test wants to lie about,
 * so a refusal can be attributed to that field and nothing else.
 */

export class Writer {
  #parts: Uint8Array[] = [];

  tag(text: string): this {
    // The six-byte markers the format writes between structures.
    const bytes = new Uint8Array(6);
    for (let index = 0; index < Math.min(text.length, 6); index++)
      bytes[index] = text.charCodeAt(index);
    this.#parts.push(bytes);
    return this;
  }

  u1(value: number): this {
    this.#parts.push(new Uint8Array([value & 0xff]));
    return this;
  }

  u4(value: number): this {
    const bytes = new Uint8Array(4);
    new DataView(bytes.buffer).setUint32(0, value >>> 0, true);
    this.#parts.push(bytes);
    return this;
  }

  u8(value: number | bigint): this {
    const bytes = new Uint8Array(8);
    new DataView(bytes.buffer).setBigUint64(0, BigInt(value), true);
    this.#parts.push(bytes);
    return this;
  }

  s8(value: number | bigint): this {
    const bytes = new Uint8Array(8);
    new DataView(bytes.buffer).setBigInt64(0, BigInt(value), true);
    this.#parts.push(bytes);
    return this;
  }

  f8(value: number): this {
    const bytes = new Uint8Array(8);
    new DataView(bytes.buffer).setFloat64(0, value, true);
    this.#parts.push(bytes);
    return this;
  }

  raw(bytes: Uint8Array): this {
    this.#parts.push(bytes);
    return this;
  }

  bytes(): Uint8Array {
    let total = 0;
    for (const part of this.#parts) total += part.length;
    const out = new Uint8Array(total);
    let offset = 0;
    for (const part of this.#parts) {
      out.set(part, offset);
      offset += part.length;
    }
    return out;
  }
}

export interface SegmentSpec {
  readonly records: number | bigint;
  readonly compressionClass?: number;
  readonly subCompressionClass?: number;
  readonly minDataId?: number;
  readonly hasNulls?: boolean;
  readonly subsegmentRecords?: number | bigint;
}

/** A `.idfmeta` member carrying one descriptor per spec. */
export function idfmeta(segments: readonly SegmentSpec[]): Uint8Array {
  const writer = new Writer().tag("<1:CP").u8(segments.length);
  for (const segment of segments) {
    const compressionClass = segment.compressionClass ?? 703066;
    writer
      .tag("<1:CS")
      .u8(segment.records)
      .u8(0) // base id
      .u4(compressionClass)
      .u4(segment.subCompressionClass ?? 703032);
    if (compressionClass === 703066) writer.u8(0).u8(0).u8(0).u1(0);
    writer
      .u4(0) // first run value
      .tag("<1:SS")
      .u8(0) // distinct states
      .u4(segment.minDataId ?? 3)
      .u4(0)
      .u4(0)
      .s8(0)
      .u8(segment.records)
      .u1(segment.hasNulls === true ? 1 : 0)
      .u8(0)
      .u8(0)
      .tag("SS:1>");
    if (segment.subsegmentRecords === undefined) writer.u1(0);
    else
      writer
        .u1(1)
        .tag("<1:CS")
        .u8(segment.subsegmentRecords)
        .u8(0)
        .u1(0)
        .tag("CS:1>");
    writer.tag("CS:1>");
  }
  return writer.bytes();
}

export interface IdfSpec {
  /** Run-length entries as data value and repeat value pairs. */
  readonly runs: readonly (readonly [number, number])[];
  readonly subWords?: number | bigint;
  readonly sub?: Uint8Array;
  /** Lie about the entry count without writing the entries. */
  readonly declaredEntries?: number | bigint;
}

export function idf(segments: readonly IdfSpec[]): Uint8Array {
  const writer = new Writer();
  for (const segment of segments) {
    writer.u8(segment.declaredEntries ?? segment.runs.length);
    for (const [dataValue, repeat] of segment.runs)
      writer.u4(dataValue).u4(repeat);
    const sub = segment.sub ?? new Uint8Array(0);
    writer.u8(segment.subWords ?? sub.length / 8).raw(sub);
  }
  return writer.bytes();
}

/** A numeric `.dictionary` member: type 0 is int64, type 1 is double. */
export function numericDictionary(
  type: 0 | 1,
  values: readonly (number | bigint)[],
  options: { declaredCount?: number | bigint; elementSize?: number } = {},
): Uint8Array {
  const writer = new Writer();
  const header = new Uint8Array(28);
  new DataView(header.buffer).setInt32(0, type, true);
  writer.raw(header);
  writer
    .u8(options.declaredCount ?? values.length)
    .u4(options.elementSize ?? 8);
  for (const value of values) {
    if (type === 0) writer.s8(value as bigint);
    else writer.f8(value as number);
  }
  return writer.bytes();
}

/** A `.dictionary` member whose type the reader does not support. */
export function unsupportedDictionary(type: number): Uint8Array {
  const header = new Uint8Array(28);
  new DataView(header.buffer).setInt32(0, type, true);
  return header;
}

export interface StringPageSpec {
  /** An uncompressed page: NUL-terminated UTF-16 runs. */
  readonly text?: readonly string[];
  /** A Huffman page: the encode array, the buffer, and its declared bit count. */
  readonly huffman?: {
    readonly encodeArray: Uint8Array;
    readonly buffer: Uint8Array;
    readonly totalBits: number;
    readonly charsetTypeId?: number;
    readonly charsetByte?: number;
  };
}

/** A string `.dictionary` member with the pages and record handles given. */
export function stringDictionary(
  pages: readonly StringPageSpec[],
  handles: readonly (readonly [offset: number, page: number])[],
): Uint8Array {
  const writer = new Writer();
  const header = new Uint8Array(28);
  new DataView(header.buffer).setInt32(0, 2, true);
  writer.raw(header);
  writer.s8(0).u1(0).s8(0).s8(pages.length);
  for (const page of pages) {
    writer.u8(0).u1(0).u8(0).u8(0);
    if (page.huffman === undefined) {
      const joined = `${(page.text ?? []).join("\0")}\0`;
      const chars = new Uint8Array(joined.length * 2);
      const view = new DataView(chars.buffer);
      for (let index = 0; index < joined.length; index++)
        view.setUint16(index * 2, joined.charCodeAt(index), true);
      writer
        .u1(0)
        .u4(0)
        .u8(0)
        .u8(joined.length)
        .u8(chars.length)
        .raw(chars)
        .u4(0);
    } else {
      const charsetTypeId = page.huffman.charsetTypeId ?? 0x000aba91;
      writer
        .u1(1)
        .u4(0)
        .u4(page.huffman.totalBits)
        .u4(charsetTypeId)
        .u8(page.huffman.buffer.length);
      if (charsetTypeId === 0x000aba91)
        writer.u1(page.huffman.charsetByte ?? 0);
      writer
        .u4(0)
        .raw(page.huffman.encodeArray)
        .u8(0)
        .raw(page.huffman.buffer)
        .u4(0);
    }
  }
  writer.u8(handles.length).u4(8);
  for (const [offset, page] of handles) writer.u4(offset).u4(page);
  return writer.bytes();
}
