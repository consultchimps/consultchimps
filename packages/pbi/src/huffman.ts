/**
 * Canonical Huffman decoding for VertiPaq string dictionary pages.
 *
 * Ported to TypeScript from `src/xmhuffman_kernel.c` of
 * https://github.com/Hugoberry/xmhuffman-cython, commit
 * 2f908c56f29ae526b3b35772186f095f0a5eaaec ("Decode at page grain for parallel
 * loads", 2026-05-15), MIT licence. No C from that repository is compiled here;
 * only the algorithm is reproduced. The licence text is reproduced in
 * THIRD-PARTY-LICENSES.md at the repository root.
 *
 * The alphabet is 256 symbols and the code lengths arrive nibble-packed in the
 * page's 128-byte encode array. VertiPaq writes the bitstream as 16-bit words,
 * so the buffer is byte-pair swapped before any code is read.
 */

const MAX_CODE_LENGTH = 15;

/** A corrupt or unusable code table. The caller turns this into a column exclusion. */
export class HuffmanError extends Error {}

/** 128 nibble-packed bytes become 256 code lengths. */
export function expandEncodeArray(packed: Uint8Array): Uint8Array {
  const lengths = new Uint8Array(256);
  for (let index = 0; index < 128; index++) {
    const byte = packed[index] ?? 0;
    lengths[2 * index] = byte & 0x0f;
    lengths[2 * index + 1] = (byte >> 4) & 0x0f;
  }
  return lengths;
}

/** Bytes 2k and 2k+1 swap; a trailing odd byte is copied through. */
export function swapPairs(buffer: Uint8Array): Uint8Array {
  const length = buffer.length;
  const out = new Uint8Array(length);
  const even = length & ~1;
  for (let index = 0; index < even; index += 2) {
    out[index] = buffer[index + 1]!;
    out[index + 1] = buffer[index]!;
  }
  if (length & 1) out[length - 1] = buffer[length - 1]!;
  return out;
}

export interface HuffmanTable {
  /** Flat canonical lookup: entry = (symbol << 8) | codeLength. */
  readonly table: Uint16Array;
  readonly maxLength: number;
}

export function buildTable(lengths: Uint8Array): HuffmanTable {
  const count = new Uint32Array(MAX_CODE_LENGTH + 1);
  let maxLength = 0;
  for (let symbol = 0; symbol < 256; symbol++) {
    const length = lengths[symbol]!;
    if (length === 0) continue;
    if (length > MAX_CODE_LENGTH)
      throw new HuffmanError("code length above the canonical maximum");
    count[length]!++;
    if (length > maxLength) maxLength = length;
  }
  if (maxLength === 0) return { table: new Uint16Array(0), maxLength: 0 };
  let left = 1;
  for (let length = 1; length <= maxLength; length++) {
    left <<= 1;
    if (count[length]! > left) throw new HuffmanError("over-subscribed code");
    left -= count[length]!;
  }
  if (left !== 0) throw new HuffmanError("incomplete code");
  const next = new Uint32Array(MAX_CODE_LENGTH + 2);
  let code = 0;
  for (let length = 1; length <= maxLength; length++) {
    code = (code + count[length - 1]!) << 1;
    next[length] = code;
  }
  const table = new Uint16Array(1 << maxLength);
  for (let symbol = 0; symbol < 256; symbol++) {
    const length = lengths[symbol]!;
    if (length === 0) continue;
    const assigned = next[length]!++;
    const pad = maxLength - length;
    const base = assigned << pad;
    const span = 1 << pad;
    const entry = ((symbol << 8) | length) & 0xffff;
    for (let step = 0; step < span; step++) table[base + step] = entry;
  }
  return { table, maxLength };
}

/**
 * Thirty-two big-endian bits from `byte`, zero-padded past the end. A code is
 * at most 15 bits and a bit offset at most 7, so this window always holds one
 * whole codeword.
 */
function bigEndian32(buffer: Uint8Array, length: number, byte: number): number {
  const b0 = byte < length ? buffer[byte]! : 0;
  const b1 = byte + 1 < length ? buffer[byte + 1]! : 0;
  const b2 = byte + 2 < length ? buffer[byte + 2]! : 0;
  const b3 = byte + 3 < length ? buffer[byte + 3]! : 0;
  return ((b0 << 24) | (b1 << 16) | (b2 << 8) | b3) >>> 0;
}

/**
 * Decode every string of one page into raw byte runs.
 *
 * `offsets` are ascending per-string start bits; the last string ends at
 * `totalBits`. A zero `charsetByte` emits one byte per symbol; a nonzero one
 * pairs each symbol with the charset byte so the run reads back as UTF-16LE.
 */
export function decodePage(
  swapped: Uint8Array,
  huffman: HuffmanTable,
  offsets: readonly number[],
  totalBits: number,
  charsetByte: number,
): Uint8Array[] {
  const count = offsets.length;
  const out: Uint8Array[] = new Array<Uint8Array>(count);
  const { table, maxLength } = huffman;
  if (maxLength === 0) {
    for (let index = 0; index < count; index++) out[index] = new Uint8Array(0);
    return out;
  }
  // A page cannot decode past its own bytes. The reference implementation
  // trusts the page's declared bit count, and `bigEndian32` zero-pads past the
  // end of the buffer, so a canonical table that assigns the all-zero codeword
  // to a real symbol would emit that symbol forever: four bytes of page can ask
  // for gigabytes of output. The page's own length in bits is the ceiling the
  // format already implies.
  const pageBits = swapped.length * 8;
  if (!Number.isSafeInteger(totalBits) || totalBits < 0 || totalBits > pageBits)
    throw new HuffmanError("page bit count runs past the page buffer");
  const shift = 32 - maxLength;
  const mask = (1 << maxLength) - 1;
  const length = swapped.length;
  const scratch = new Uint8Array(65536);
  const stride = charsetByte ? 2 : 1;
  for (let index = 0; index < count; index++) {
    const startBit = offsets[index]!;
    const endBit = index + 1 === count ? totalBits : offsets[index + 1]!;
    if (
      !Number.isSafeInteger(startBit) ||
      startBit < 0 ||
      endBit < startBit ||
      endBit > totalBits
    )
      throw new HuffmanError("page offsets are not ascending");
    // The shortest codeword is one bit, so this run emits at most one symbol
    // per bit of its own span. The buffer is bounded before it is grown.
    const limit = (endBit - startBit) * stride;
    let bit = startBit;
    let cursor = 0;
    let buffer = scratch;
    while (bit < endBit) {
      const byte = bit >>> 3;
      const offset = bit & 7;
      const window = bigEndian32(swapped, length, byte);
      const entry = table[(window >>> (shift - offset)) & mask]!;
      const codeLength = entry & 0xff;
      if (codeLength === 0) throw new HuffmanError("corrupt bitstream");
      // A codeword that runs past this string's end bit would be completed from
      // the next string's bits or from the zero padding past the page. The
      // symbol it produced would be invented, and the column would keep its
      // value count and export fabricated text.
      if (bit + codeLength > endBit)
        throw new HuffmanError("codeword crosses the string boundary");
      if (cursor + stride > limit) throw new HuffmanError("corrupt bitstream");
      if (cursor + stride > buffer.length) {
        const grown = new Uint8Array(Math.min(buffer.length * 2, limit));
        grown.set(buffer.subarray(0, cursor));
        buffer = grown;
      }
      buffer[cursor++] = entry >> 8;
      if (charsetByte) buffer[cursor++] = charsetByte;
      bit += codeLength;
    }
    out[index] = buffer.slice(0, cursor);
  }
  return out;
}
