/**
 * VertiPaq column storage: the `.idfmeta` segment descriptors, the `.idf`
 * run-length plus bit-packed hybrid vectors, and the `.dictionary` pages.
 * The algorithms follow pbixray (MIT) and MS-XLDM. Every read is bounds
 * checked, which the reference implementation does not do, so a truncated or
 * crafted member excludes one column instead of reading past its member.
 */
import {
  buildTable,
  decodePage,
  expandEncodeArray,
  HuffmanError,
  swapPairs,
} from "./huffman.js";

/** The reserved data id every null cell carries. */
export const XM_DATA_ID_NULL = 2;
/** The first data id a dictionary entry can hold. */
export const XM_FIRST_DATA_ID = 3;

const HUFFMAN_CHARSET_BASED = 0x000aba91;
const HUFFMAN_GENERAL = 0x000aba92;
/** Hybrid run-length encoding: the bit width lives in the sub class. */
const HYBRID_RLE = 703066;

const utf16 = new TextDecoder("utf-16le");
const latin1 = new TextDecoder("latin1");

/** A member that is structurally inconsistent. One column, not the model. */
export class VertipaqError extends Error {}

class Reader {
  readonly bytes: Uint8Array;
  readonly view: DataView;
  position = 0;

  constructor(bytes: Uint8Array) {
    this.bytes = bytes;
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }

  get atEnd(): boolean {
    return this.position >= this.bytes.length;
  }

  #need(count: number): number {
    const at = this.position;
    if (count < 0 || at > this.bytes.length - count)
      throw new VertipaqError("read past the end of the member");
    this.position = at + count;
    return at;
  }

  u1(): number {
    return this.view.getUint8(this.#need(1));
  }

  u4(): number {
    return this.view.getUint32(this.#need(4), true);
  }

  s4(): number {
    return this.view.getInt32(this.#need(4), true);
  }

  /** A 64-bit count. Anything past the safe-integer range is inconsistent. */
  u8count(): number {
    const value = this.view.getBigUint64(this.#need(8), true);
    if (value > BigInt(Number.MAX_SAFE_INTEGER))
      throw new VertipaqError("count is out of range");
    return Number(value);
  }

  s8count(): number {
    const value = this.view.getBigInt64(this.#need(8), true);
    if (
      value > BigInt(Number.MAX_SAFE_INTEGER) ||
      value < BigInt(Number.MIN_SAFE_INTEGER)
    )
      throw new VertipaqError("count is out of range");
    return Number(value);
  }

  s8(): bigint {
    return this.view.getBigInt64(this.#need(8), true);
  }

  f8(): number {
    return this.view.getFloat64(this.#need(8), true);
  }

  take(count: number): Uint8Array {
    const at = this.#need(count);
    return this.bytes.subarray(at, at + count);
  }

  get remaining(): number {
    return this.bytes.length - this.position;
  }

  /**
   * A declared element count the bytes still to be read could actually supply.
   * Every allocation sized from a file field goes through this first, so no
   * typed array is ever created from a number nothing has backed yet.
   */
  bounded(count: number, elementBytes: number, what: string): number {
    if (
      !Number.isSafeInteger(count) ||
      count < 0 ||
      elementBytes <= 0 ||
      count > Math.floor(this.remaining / elementBytes)
    )
      throw new VertipaqError(`${what} is inconsistent`);
    return count;
  }
}

/**
 * A decoded column that cannot sit on the catalog's declared row count. It is
 * the row-alignment exclusion, not a corrupt member.
 */
export class VertipaqAlignmentError extends VertipaqError {}

/** Bit width implied by a segment's compression class (the MS-XLDM NoSplit map). */
function bitWidthOf(compressionClass: number, subClass: number): number {
  const width = (value: number): number =>
    value >= 703031 && value <= 703040
      ? value - 703030
      : value === 703042
        ? 12
        : value === 703046
          ? 16
          : value === 703051
            ? 21
            : value === 703062
              ? 32
              : 0;
  return compressionClass === HYBRID_RLE
    ? width(subClass)
    : width(compressionClass);
}

export interface SegmentStatistics {
  readonly minDataId: number;
  readonly hasNulls: boolean;
  readonly recordCount: number;
}

export interface SegmentDescriptor {
  readonly records: number;
  readonly bitWidth: number;
  readonly countBitPacked: number;
  readonly minDataId: number;
  readonly hasNulls: boolean;
  readonly compressionClass: number;
  readonly subCompressionClass: number;
  /** The base data id a bit-packed subsegment is offset from. */
  readonly minDataIdBase: number;
}

interface RawSegment {
  records: number;
  hasSubsegment: number;
  subsegmentRecords: number;
  bitWidth: number;
  compressionClass: number;
  subCompressionClass: number;
  minDataId: number;
  hasNulls: boolean;
}

function readCompressionSegment(reader: Reader, isSub: boolean): RawSegment {
  reader.take(6); // "<1:CS\0"
  const records = reader.u8count();
  reader.u8count(); // base id, unused by the decoder
  let compressionClass = 0;
  let subCompressionClass = 0;
  let minDataId = 0;
  let hasNulls = false;
  if (!isSub) {
    compressionClass = reader.u4();
    subCompressionClass = reader.u4();
    if (compressionClass === HYBRID_RLE) {
      reader.u8count();
      reader.u8count();
      reader.u8count();
      reader.u1();
    }
    reader.u4(); // first run value
    reader.take(6); // "<1:SS\0"
    reader.u8count(); // distinct states
    minDataId = reader.u4();
    reader.u4(); // max data id
    reader.u4(); // original min segment data id
    reader.s8count(); // RLE sort order
    reader.u8count(); // row count
    hasNulls = reader.u1() !== 0;
    reader.u8count(); // RLE runs
    reader.u8count(); // others RLE runs
    reader.take(6); // "SS:1>\0"
  }
  const hasSubsegment = reader.u1();
  let subsegmentRecords = 0;
  if (hasSubsegment !== 0)
    subsegmentRecords = readCompressionSegment(reader, true).records;
  reader.take(6); // "CS:1>\0"
  return {
    records,
    hasSubsegment,
    subsegmentRecords,
    bitWidth: isSub ? 0 : bitWidthOf(compressionClass, subCompressionClass),
    compressionClass,
    subCompressionClass,
    minDataId,
    hasNulls,
  };
}

/**
 * The per-segment shape the decoder consumes, in descriptor order.
 *
 * `rowCount` is the catalog's declared row count for the column's table. A
 * descriptor claiming more records than the table has rows cannot be placed on
 * the table's rows at all, so it is the row-alignment exclusion rather than an
 * invitation to allocate whatever the header asks for.
 */
export function parseIdfmeta(
  member: Uint8Array,
  rowCount: number,
): SegmentDescriptor[] {
  const reader = new Reader(member);
  reader.take(6); // "<1:CP\0"
  // A descriptor is tens of bytes at minimum, so the member's own remaining
  // length is the ceiling, checked before anything is allocated.
  const count = reader.bounded(reader.u8count(), 1, "segment count");
  const segments: SegmentDescriptor[] = [];
  let records = 0;
  for (let index = 0; index < count; index++) {
    const raw = readCompressionSegment(reader, false);
    records += raw.records;
    if (
      raw.records < 0 ||
      !Number.isSafeInteger(records) ||
      records > rowCount ||
      raw.subsegmentRecords > rowCount
    )
      throw new VertipaqAlignmentError("segment records exceed the table rows");
    segments.push({
      records: raw.records,
      bitWidth: raw.bitWidth,
      countBitPacked: raw.hasSubsegment !== 0 ? raw.subsegmentRecords : 0,
      minDataId: raw.minDataId,
      hasNulls: raw.hasNulls,
      compressionClass: raw.compressionClass,
      subCompressionClass: raw.subCompressionClass,
      // A segment holding nulls is based at the reserved null id; any other is
      // based at its own minimum.
      minDataIdBase: raw.hasNulls ? XM_DATA_ID_NULL : raw.minDataId,
    });
  }
  return segments;
}

export interface IdfSegment {
  readonly dataValues: Uint32Array;
  readonly repeatValues: Uint32Array;
  readonly subWords: number;
  readonly sub: Uint8Array;
}

export function parseIdf(member: Uint8Array): IdfSegment[] {
  const reader = new Reader(member);
  const segments: IdfSegment[] = [];
  while (!reader.atEnd) {
    // Each entry is one data value and one repeat value, eight bytes together,
    // so the member's own remaining length is the ceiling. The two typed arrays
    // are created only after the declared count has passed it.
    const entries = reader.bounded(reader.u8count(), 8, "entry count");
    const dataValues = new Uint32Array(entries);
    const repeatValues = new Uint32Array(entries);
    for (let index = 0; index < entries; index++) {
      dataValues[index] = reader.u4();
      repeatValues[index] = reader.u4();
    }
    const subWords = reader.bounded(reader.u8count(), 8, "word count");
    const sub = reader.take(subWords * 8);
    segments.push({ dataValues, repeatValues, subWords, sub });
  }
  return segments;
}

/** Per run-length entry output counts, capped at the segment's declared records. */
function realRepeats(
  segment: IdfSegment,
  descriptor: SegmentDescriptor,
): { per: Uint32Array; total: number } {
  const cap = descriptor.records || 0;
  const per = new Uint32Array(segment.repeatValues.length);
  let accumulated = 0;
  for (let index = 0; index < segment.repeatValues.length; index++) {
    if (cap && accumulated >= cap) continue;
    let repeat = segment.repeatValues[index]!;
    if (cap && accumulated + repeat > cap) repeat = cap - accumulated;
    per[index] = repeat;
    accumulated += repeat;
  }
  // Without a declared record count there is no ceiling on the sum of four-byte
  // repeat values, so a segment that declares none is refused rather than
  // allowed to size a vector from two header fields.
  if (cap === 0 && accumulated > 0)
    throw new VertipaqError("segment declares no record count");
  return { per, total: accumulated };
}

/** Unpack bitWidth-wide little-endian values from 64-bit words, plus minDataId. */
function readBitPacked(
  sub: Uint8Array,
  bitWidth: number,
  minDataId: number,
  entries: number,
): Float64Array {
  // A width of zero would make perWord infinite and the allocation a RangeError.
  // The one caller checks it too; this keeps the function safe on its own.
  if (!Number.isInteger(bitWidth) || bitWidth < 1 || bitWidth > 64)
    throw new VertipaqError("bit width is inconsistent");
  const words = Math.floor(sub.length / 8);
  if (words === 0) return new Float64Array(0);
  const perWord = Math.floor(64 / bitWidth);
  const view = new DataView(sub.buffer, sub.byteOffset, sub.byteLength);
  // At one bit per value a word holds 64 values, so the words alone would
  // amplify 64-fold. The descriptor's declared subsegment count is the real
  // ceiling: nothing past it is ever read.
  const out = new Float64Array(Math.min(words * perWord, entries));
  const wrap = Math.pow(2, bitWidth);
  let cursor = 0;
  for (let word = 0; word < words; word++) {
    const low = view.getUint32(word * 8, true);
    const high = view.getUint32(word * 8 + 4, true);
    for (let slot = 0; slot < perWord; slot++) {
      const shift = slot * bitWidth;
      let value: number;
      if (shift + bitWidth <= 32) {
        value =
          bitWidth === 32 ? low >>> 0 : (low >>> shift) & ((1 << bitWidth) - 1);
      } else if (shift >= 32) {
        const inner = shift - 32;
        value =
          bitWidth === 32
            ? high >>> 0
            : (high >>> inner) & ((1 << bitWidth) - 1);
      } else {
        const lowBits = 32 - shift;
        const lowPart = low >>> shift;
        const highPart = high % Math.pow(2, bitWidth - lowBits);
        value = lowPart + highPart * Math.pow(2, lowBits);
      }
      if (cursor >= out.length) return out;
      out[cursor++] = (value < 0 ? value + wrap : value) + minDataId;
    }
  }
  return out;
}

/** One segment's run-length plus bit-packed hybrid, as data ids. */
export function decodeSegmentIds(
  segment: IdfSegment,
  descriptor: SegmentDescriptor,
): Float64Array {
  const { per, total } = realRepeats(segment, descriptor);
  const vector = new Float64Array(total);
  const entries = descriptor.countBitPacked;
  const bitWidth = descriptor.bitWidth;
  const base = descriptor.minDataIdBase;
  let packed: Float64Array | null = null;
  if (entries > 0) {
    if (bitWidth <= 0)
      // Without a known width a bit-packed subsegment cannot be read at all.
      throw new VertipaqError("bit-packed subsegment has no known width");
    const words = Math.floor(segment.sub.length / 8);
    const view = new DataView(
      segment.sub.buffer,
      segment.sub.byteOffset,
      segment.sub.byteLength,
    );
    const lastZero =
      words > 0 && view.getBigUint64((words - 1) * 8, true) === 0n;
    packed =
      lastZero && segment.subWords === 1
        ? new Float64Array(entries).fill(base)
        : readBitPacked(segment.sub, bitWidth, base, entries);
  }
  let position = 0;
  let packedOffset = 0;
  for (let index = 0; index < segment.dataValues.length; index++) {
    const count = per[index]!;
    if (count === 0) continue;
    const dataValue = segment.dataValues[index]!;
    if (dataValue + packedOffset === 0xffffffff) {
      if (packed === null)
        throw new VertipaqError("missing bit-packed subsegment");
      // An exhausted subsegment must not be filled with zeros: the column would
      // keep its row count and export fabricated ids as real values instead of
      // being excluded.
      if (packedOffset + count > packed.length)
        throw new VertipaqError("bit-packed subsegment is exhausted");
      for (let step = 0; step < count; step++)
        vector[position + step] = packed[packedOffset + step]!;
      packedOffset += count;
      position += count;
    } else {
      vector.fill(dataValue, position, position + count);
      position += count;
    }
  }
  return vector;
}

export type DictionaryValue = string | number | bigint;

export interface Dictionary {
  readonly values: readonly DictionaryValue[];
  readonly minId: number;
  readonly isString: boolean;
  /** Pages that used a non-Latin charset byte, an unverified reference path. */
  readonly nonLatinPages: number;
}

/** Null when the dictionary type is one the reader does not support. */
export function parseDictionary(
  member: Uint8Array,
  minDataId: number,
): Dictionary | null {
  const reader = new Reader(member);
  const dictionaryType = reader.s4();
  for (let index = 0; index < 6; index++) reader.s4(); // hash info
  if (dictionaryType === 2) return parseStringDictionary(reader, minDataId);
  if (dictionaryType === 0 || dictionaryType === 1) {
    const declared = reader.u8count();
    const elementSize = reader.u4();
    // Four-byte elements are read as an int32; every other width reads eight
    // bytes, which is the reference implementation's own rule.
    const count = reader.bounded(
      declared,
      elementSize === 4 ? 4 : 8,
      "value count",
    );
    const values: DictionaryValue[] = new Array<DictionaryValue>(count);
    for (let index = 0; index < count; index++) {
      if (elementSize === 4) values[index] = reader.s4();
      // A long dictionary keeps its bigint: an identifier past 2^53 must not
      // lose its low digits on the way to the workbook.
      else if (dictionaryType === 0) values[index] = reader.s8();
      else values[index] = reader.f8();
    }
    return { values, minId: minDataId, isString: false, nonLatinPages: 0 };
  }
  return null;
}

interface UncompressedPage {
  readonly compressed: false;
  readonly text: string;
}

interface CompressedPage {
  readonly compressed: true;
  readonly totalBits: number;
  readonly charsetTypeId: number;
  readonly charsetUsed: number;
  readonly encodeArray: Uint8Array;
  readonly buffer: Uint8Array;
}

function parseStringDictionary(reader: Reader, minDataId: number): Dictionary {
  reader.s8count(); // stored string count
  reader.u1(); // compressed flag
  reader.s8count(); // longest stored string
  // A page header is more than thirty bytes before its payload, so the bytes
  // still to be read bound the page count before the loop starts.
  const pageCount = reader.bounded(reader.s8count(), 30, "page count");
  const pages: (UncompressedPage | CompressedPage)[] = [];
  for (let index = 0; index < pageCount; index++) {
    reader.u8count(); // page mask
    reader.u1(); // page contains nulls
    reader.u8count(); // page start index
    reader.u8count(); // page string count
    const compressed = reader.u1();
    reader.take(4); // DD CC BB AA
    if (compressed === 0) {
      reader.u8count(); // remaining
      const usedChars = reader.u8count();
      const allocated = reader.u8count();
      const chars = reader.take(usedChars * 2);
      if (allocated < usedChars * 2)
        throw new VertipaqError("page allocation is inconsistent");
      reader.take(allocated - usedChars * 2);
      pages.push({ compressed: false, text: utf16.decode(chars) });
    } else {
      const totalBits = reader.u4();
      const charsetTypeId = reader.u4();
      const bufferLength = reader.u8count();
      let charsetUsed = 0;
      if (charsetTypeId === HUFFMAN_CHARSET_BASED) charsetUsed = reader.u1();
      reader.u4(); // decode bits
      const encodeArray = reader.take(128);
      reader.u8count(); // buffer size
      pages.push({
        compressed: true,
        totalBits,
        charsetTypeId,
        charsetUsed,
        encodeArray,
        buffer: reader.take(bufferLength),
      });
    }
    reader.take(4); // CD AB CD AB
  }
  reader.u8count(); // declared handle count, re-derived from the bytes below
  reader.take(4); // element size
  const handleCount = reader.bounded(
    Math.floor(reader.remaining / 8),
    8,
    "handle count",
  );
  const handleBytes = reader.take(handleCount * 8);
  const handles = new DataView(
    handleBytes.buffer,
    handleBytes.byteOffset,
    handleBytes.byteLength,
  );
  // Record handles grouped by page, keeping file order, which is data-id order.
  const byPage = new Map<number, number[]>();
  for (let index = 0; index < handleCount; index++) {
    const offset = handles.getUint32(index * 8, true);
    const page = handles.getUint32(index * 8 + 4, true);
    const bucket = byPage.get(page);
    if (bucket === undefined) byPage.set(page, [offset]);
    else bucket.push(offset);
  }

  const values: string[] = [];
  let nonLatinPages = 0;
  for (let pageId = 0; pageId < pages.length; pageId++) {
    const page = pages[pageId]!;
    if (!page.compressed) {
      // A zero-terminated run: the trailing empty split is padding, not a value.
      const parts = page.text.split("\0");
      parts.pop();
      for (const part of parts) values.push(part);
      continue;
    }
    const offsets = byPage.get(pageId);
    if (offsets === undefined) continue;
    const general = page.charsetTypeId === HUFFMAN_GENERAL;
    const charsetByte = general ? 0 : page.charsetUsed;
    if (charsetByte !== 0) nonLatinPages++;
    let runs: Uint8Array[];
    try {
      runs = decodePage(
        swapPairs(page.buffer),
        buildTable(expandEncodeArray(page.encodeArray)),
        offsets,
        page.totalBits,
        charsetByte,
      );
    } catch (error) {
      if (error instanceof HuffmanError)
        throw new VertipaqError("string page could not be decoded");
      throw error;
    }
    if (general)
      for (const run of runs)
        values.push(utf16.decode(run.subarray(0, run.length & ~1)));
    else if (charsetByte === 0)
      for (const run of runs) values.push(latin1.decode(run));
    else for (const run of runs) values.push(utf16.decode(run));
  }
  return { values, minId: minDataId, isString: true, nonLatinPages };
}

/**
 * True when a segment uses a compression class the width map does not know.
 * With no bit-packed subsegment the run-length path still decodes, which is an
 * unverified reference path; with one, the width is unknowable.
 */
export function unknownCompressionClass(
  descriptor: SegmentDescriptor,
): boolean {
  return descriptor.bitWidth === 0 && descriptor.compressionClass !== 0;
}
