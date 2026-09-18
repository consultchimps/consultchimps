import { unreadableModel } from "../errors.js";
import type { PipelineBudget } from "../budget.js";
import { loadXpress9 } from "./runtime.js";
import type { Xpress9Decoder, Xpress9RuntimeConfig } from "./runtime.js";

/**
 * The three markers a `DataModel` part can begin with. The two compression
 * signatures are UTF-16LE text in the first 102 bytes; the stream marker is the
 * uncompressed backup image itself, which carries a BOM first.
 */
const SIGNATURE_SINGLE = "This backup was created using XPress9 compression.";
const SIGNATURE_MULTITHREADED =
  "This backup was created using multithreaded XPrs9.";
const SIGNATURE_STREAM = "STREAM_STORAGE_SIGNATURE_)!@#$%^&*(";

/** Both compression signatures are 51 UTF-16 units inside a 102-byte header. */
const HEADER_BYTES = 102;
/** Five 64-bit counts follow a multithreaded header before the first chunk. */
const MULTITHREADED_COUNTS_BYTES = 40;
/** An eight-byte chunk header plus at least one byte of compressed payload. */
const MINIMUM_FRAMED_CHUNK = 9;

/**
 * What the module holds before either framing buffer exists: its own code,
 * stack, the LZ77 history window and the allocator's initial arena. It is the
 * build's own `-sINITIAL_MEMORY`, pinned in `scripts/build-wasm.mjs`.
 */
const RUNTIME_BASELINE_BYTES = 16 * 1024 * 1024;

/** WebAssembly grows linear memory a page at a time. */
const WASM_PAGE_BYTES = 64 * 1024;

/**
 * Emscripten's `MEMORY_GROWTH_GEOMETRIC_STEP`, pinned explicitly in
 * `scripts/build-wasm.mjs` so this prediction reads a flag rather than a
 * default that a toolchain upgrade could move. A heap that must grow does not
 * grow to what was asked for: it grows by this fraction on top.
 */
const MEMORY_GROWTH_GEOMETRIC_STEP = 0.2;

/**
 * dlmalloc's own bookkeeping around the two buffers: a fixed allowance for its
 * arena headers plus a fraction of what is being allocated, because a larger
 * request is split across more chunks.
 */
const ALLOCATOR_SLACK_BYTES = 64 * 1024;
const ALLOCATOR_SLACK_RATIO = 0.01;

/**
 * A conservative prediction of the linear memory the module will hold once it
 * carries buffers of these capacities.
 *
 * Summing the payloads under-counts, which is the defect this replaces: with
 * `ALLOW_MEMORY_GROWTH` the heap grows geometrically, rounds up to whole pages,
 * and carries allocator overhead, so the capacity that results is larger than
 * the bytes that were asked for. The corpus records 33,554,432 bytes of heap on
 * the largest fixture against about 20.2 MB of payload, and a `peakBytes`
 * between the two would have passed an allocation beyond the bound.
 */
export function predictedHeapBytes(
  baseline: number,
  sourceCapacity: number,
  destinationCapacity: number,
  currentHeap: number,
): number {
  const payload = sourceCapacity + destinationCapacity;
  const needed =
    baseline +
    payload +
    ALLOCATOR_SLACK_BYTES +
    Math.ceil(payload * ALLOCATOR_SLACK_RATIO);
  // A heap that already holds this much does not grow, so the geometric step
  // applies only to what growth would have to reach.
  const target = Math.max(
    currentHeap,
    Math.ceil(needed * (1 + MEMORY_GROWTH_GEOMETRIC_STEP)),
  );
  return Math.ceil(target / WASM_PAGE_BYTES) * WASM_PAGE_BYTES;
}

export type ModelCompression = "single" | "multithreaded" | "uncompressed";

function utf16le(bytes: Uint8Array, offset: number, units: number): string {
  let text = "";
  for (let index = 0; index < units; index++) {
    const low = bytes[offset + index * 2] ?? 0;
    const high = bytes[offset + index * 2 + 1] ?? 0;
    text += String.fromCharCode(low | (high << 8));
  }
  return text;
}

/** Which of the three forms a model part is, without touching the decoder. */
export function detectCompression(part: Uint8Array): ModelCompression | null {
  if (
    part.length >= 72 &&
    part[0] === 0xff &&
    part[1] === 0xfe &&
    utf16le(part, 2, SIGNATURE_STREAM.length) === SIGNATURE_STREAM
  )
    return "uncompressed";
  if (part.length < HEADER_BYTES) return null;
  const head = utf16le(part, 0, 51);
  if (head.startsWith(SIGNATURE_SINGLE)) return "single";
  if (head.startsWith(SIGNATURE_MULTITHREADED)) return "multithreaded";
  return null;
}

export interface ModelStream {
  /** The Analysis Services backup image. */
  readonly bytes: Uint8Array;
  readonly compression: ModelCompression;
  /** Chunks framed, across every thread group. */
  readonly chunks: number;
  /** The heap this stream's frames were charged for. */
  readonly reservedHeapBytes: number;
  /**
   * Frames whose real heap turned out larger than the prediction. It must be
   * zero: a nonzero count means the prediction is not an upper bound and the
   * budget was briefly below reality.
   */
  readonly heapUnderpredictions: number;
}

interface Frame {
  readonly uncompressedSize: number;
  readonly compressedSize: number;
  readonly start: number;
}

/**
 * Reads one chunk header and bounds-checks it against the model part. The
 * reference implementation reads these without any check; an untrusted header
 * would otherwise index past the buffer or walk backwards forever.
 */
function frameAt(view: DataView, length: number, cursor: number): Frame {
  if (cursor < 0 || cursor > length - 8) throw unreadableModel("xpress9");
  const uncompressedSize = view.getUint32(cursor, true);
  const compressedSize = view.getUint32(cursor + 4, true);
  const start = cursor + 8;
  if (
    uncompressedSize <= 0 ||
    compressedSize <= 0 ||
    compressedSize > length - start
  )
    throw unreadableModel("xpress9");
  return { uncompressedSize, compressedSize, start };
}

/**
 * Frame and decompress a `DataModel` part into the backup image it holds.
 *
 * A single-threaded stream is one decoder across every chunk, so the LZ77
 * history window carries. A multithreaded stream is a prefix group and a main
 * group, each thread of each group an independent stream with its own window,
 * so each gets its own decoder.
 */
export async function decompressModelPart(
  part: Uint8Array,
  budget: PipelineBudget,
  config: Xpress9RuntimeConfig | undefined,
  /** The loader, so a test can observe when the decoder is actually called. */
  load: (
    runtime?: Xpress9RuntimeConfig,
  ) => Promise<Xpress9Decoder> = loadXpress9,
): Promise<ModelStream> {
  const compression = detectCompression(part);
  if (compression === null) throw unreadableModel("xpress9");
  if (compression === "uncompressed")
    return {
      bytes: part,
      compression,
      chunks: 0,
      reservedHeapBytes: 0,
      heapUnderpredictions: 0,
    };

  const length = part.byteLength;
  const view = new DataView(part.buffer, part.byteOffset, length);
  const pieces: Uint8Array[] = [];
  let total = 0;
  let chunks = 0;
  let decoder: Xpress9Decoder | undefined;

  const openDecoder = async (): Promise<Xpress9Decoder> => {
    decoder?.close();
    const next = await load(config);
    decoder = next;
    return next;
  };

  // The runtime keeps one source and one destination buffer and only ever grows
  // them, so the capacities any frame of this stream has asked for determine
  // the heap the module ends up holding.
  let sourceCapacity = 0;
  let destinationCapacity = 0;
  let reservedHeap = RUNTIME_BASELINE_BYTES;
  let underpredictions = 0;

  const chargeHeap = (predicted: number): void => {
    if (predicted <= reservedHeap) return;
    budget.reserve("xpress9LinearMemory", predicted - reservedHeap, "xpress9");
    reservedHeap = predicted;
  };

  const emit = (session: Xpress9Decoder, frame: Frame): void => {
    // Check the budget before the allocation, never by catching an
    // out-of-memory. The prediction covers the linear memory the module will
    // hold once it carries buffers this frame's sizes imply, and it is charged
    // before the call that would allocate them.
    sourceCapacity = Math.max(sourceCapacity, frame.compressedSize);
    destinationCapacity = Math.max(destinationCapacity, frame.uncompressedSize);
    chargeHeap(
      predictedHeapBytes(
        RUNTIME_BASELINE_BYTES,
        sourceCapacity,
        destinationCapacity,
        session.memoryBytes(),
      ),
    );
    budget.decode(frame.uncompressedSize, "xpress9");
    budget.reserve("backupImage", frame.uncompressedSize, "xpress9");
    const bytes = session.decompress(
      part.subarray(frame.start, frame.start + frame.compressedSize),
      frame.uncompressedSize,
    );
    // The prediction is meant to be an upper bound. Reading the real figure
    // afterwards keeps the accounting honest if it ever is not, and the count
    // is reported so a test can insist the prediction held on every fixture.
    const actual = session.memoryBytes();
    if (actual > reservedHeap) {
      underpredictions++;
      chargeHeap(actual);
    }
    pieces.push(bytes);
    total += bytes.length;
    chunks++;
  };

  budget.reserve("xpress9LinearMemory", RUNTIME_BASELINE_BYTES, "xpress9");
  try {
    if (compression === "single") {
      const session = await openDecoder();
      let cursor = HEADER_BYTES;
      while (cursor < length) {
        const frame = frameAt(view, length, cursor);
        emit(session, frame);
        cursor = frame.start + frame.compressedSize;
      }
    } else {
      let cursor = HEADER_BYTES;
      if (cursor > length - MULTITHREADED_COUNTS_BYTES)
        throw unreadableModel("xpress9");
      const count = (): number => {
        const value = view.getBigUint64(cursor, true);
        cursor += 8;
        if (value > BigInt(Number.MAX_SAFE_INTEGER))
          throw unreadableModel("xpress9");
        return Number(value);
      };
      const mainChunks = count();
      const prefixChunks = count();
      const prefixThreads = count();
      const mainThreads = count();
      cursor += 8; // the group's uncompressed chunk size, not needed for framing
      // Every framed chunk costs at least an eight-byte header and one byte of
      // payload, so the bytes that remain bound how many chunks the whole
      // stream can hold. Without this a header declaring 2^53 threads would
      // instantiate WebAssembly modules for ever: a hang, with no allocation
      // for the budget to refuse and no error for a caller to see.
      const framable = Math.floor((length - cursor) / MINIMUM_FRAMED_CHUNK);
      const declared = prefixThreads * prefixChunks + mainThreads * mainChunks;
      if (!Number.isSafeInteger(declared) || declared > framable)
        throw unreadableModel("xpress9");
      const group = async (threads: number, chunks: number): Promise<void> => {
        // A group with no chunks produces nothing, so it needs no decoder.
        if (chunks === 0) return;
        for (let thread = 0; thread < threads; thread++) {
          const session = await openDecoder();
          for (let chunk = 0; chunk < chunks; chunk++) {
            const frame = frameAt(view, length, cursor);
            emit(session, frame);
            cursor = frame.start + frame.compressedSize;
          }
        }
      };
      await group(prefixThreads, prefixChunks);
      await group(mainThreads, mainChunks);
    }
  } finally {
    decoder?.close();
    budget.release(reservedHeap);
  }

  if (total === 0) throw unreadableModel("xpress9");
  // The concatenation is a second live copy of the image while it is built.
  budget.reserve("backupImage", total, "xpress9");
  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const piece of pieces) {
    bytes.set(piece, offset);
    offset += piece.length;
  }
  pieces.length = 0;
  budget.release(total);
  return {
    bytes,
    compression,
    chunks,
    reservedHeapBytes: reservedHeap,
    heapUnderpredictions: underpredictions,
  };
}
