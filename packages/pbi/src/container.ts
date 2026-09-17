import { ConsultChimpsError } from "@consultchimps/core";
import { checkLimit, unboundedContainer, validateLimits } from "./budget.js";
import type { ContainerLimits, PbiContainerOptions } from "./budget.js";
import { encryptedModel, invalidContainer, unreadableModel } from "./errors.js";

interface Part {
  start: number;
  size: number;
  decodedSize: number;
  crc: number;
  method: number;
  encrypted: boolean;
}

interface Directory {
  model: Part | undefined;
  peakBytes: number;
}

// A fixed scratch allowance plus an intentionally generous per-entry bound for
// decoded names, the duplicate-name set, and parsing objects. The retained input
// buffer (which can exceed the supplied view) is accounted separately.
const scratchBytes = 64 * 1024;
const entryBytes = 4096;
const decoder = new TextDecoder("utf-8", { fatal: true });

// A resizable ArrayBuffer can grow to its maximum while the reader holds it,
// so the retained-input term of the peak estimate uses that maximum.
function backingBytes(input: Uint8Array): number {
  const buffer = input.buffer as ArrayBuffer & { maxByteLength?: number };
  return Math.max(buffer.byteLength, buffer.maxByteLength ?? 0);
}

function directory(input: Uint8Array, limits: ContainerLimits): Directory {
  // One length is captured for the view and every bounds check. A
  // length-tracking view over a growable shared buffer can report a larger
  // byteLength later; the fixed-length view would then throw on reads that a
  // live-length check had approved.
  const length = input.byteLength;
  const view = new DataView(input.buffer, input.byteOffset, length);
  const fits = (offset: number, size: number): boolean =>
    offset >= 0 && size >= 0 && offset <= length - size;
  const u16 = (offset: number): number => view.getUint16(offset, true);
  const u32 = (offset: number): number => view.getUint32(offset, true);
  const u8 = (offset: number): number => view.getUint8(offset);
  // The end-of-central-directory record is the highest-offset candidate whose
  // comment reaches the end of input AND whose directory range ends exactly at
  // the record. A decoy signature inside an archive comment therefore does not
  // hide the real record; the scan continues past it.
  let end = length - 22;
  const minimum = Math.max(0, end - 65_535);
  let count = 0;
  let size = 0;
  let start = 0;
  let unbounded = false;
  while (end >= minimum) {
    if (u32(end) === 0x06054b50 && end + 22 + u16(end + 20) === length) {
      count = u16(end + 10);
      size = u32(end + 12);
      start = u32(end + 16);
      if (
        count === 0xffff ||
        size === 0xffffffff ||
        start === 0xffffffff ||
        u16(end + 4) !== 0 ||
        u16(end + 6) !== 0 ||
        u16(end + 8) !== count
      ) {
        // A ZIP64 or multi-disk candidate cannot be tested for consistency.
        // It only decides the outcome when no consistent record exists.
        unbounded = true;
      } else if (
        fits(start, size) &&
        start + size === end &&
        count * 46 <= size
      )
        break;
    }
    end--;
  }
  if (end < minimum) {
    if (unbounded) throw unboundedContainer();
    throw invalidContainer();
  }
  // Eight bytes per directory byte cover transient UTF-16 names and sets. No
  // model allocation occurs until the whole directory and its local headers pass.
  const peakBytes =
    backingBytes(input) + scratchBytes + count * entryBytes + size * 8;
  checkLimit(limits, "peakBytes", peakBytes, "container");
  const names = new Set<string>();
  const offsets = new Set<number>();
  const ranges: { start: number; end: number }[] = [];
  let model: Part | undefined;
  let cursor = start;
  for (let index = 0; index < count; index++) {
    if (!fits(cursor, 46) || u32(cursor) !== 0x02014b50)
      throw invalidContainer();
    const flags = u16(cursor + 8);
    const method = u16(cursor + 10);
    const crc = u32(cursor + 16);
    const compressedSize = u32(cursor + 20);
    const decodedSize = u32(cursor + 24);
    const nameSize = u16(cursor + 28);
    const extraSize = u16(cursor + 30);
    const recordSize = 46 + nameSize + extraSize + u16(cursor + 32);
    const local = u32(cursor + 42);
    if (
      compressedSize === 0xffffffff ||
      decodedSize === 0xffffffff ||
      local === 0xffffffff ||
      u16(cursor + 34) !== 0
    ) {
      throw unboundedContainer();
    }
    if (
      cursor + recordSize > end ||
      !fits(local, 30) ||
      u32(local) !== 0x04034b50
    ) {
      throw invalidContainer();
    }
    const nameBytes = input.subarray(cursor + 46, cursor + 46 + nameSize);
    let name: string;
    try {
      name = decoder.decode(nameBytes);
    } catch {
      throw invalidContainer();
    }
    if (
      name.length === 0 ||
      name.includes("\0") ||
      name.includes("\\") ||
      name.startsWith("/") ||
      name.split("/").some((part) => part === "." || part === "..") ||
      names.has(name) ||
      offsets.has(local)
    )
      throw invalidContainer();
    names.add(name);
    offsets.add(local);
    const localNameSize = u16(local + 26);
    const dataStart = local + 30 + localNameSize + u16(local + 28);
    if (
      localNameSize !== nameSize ||
      !fits(dataStart, compressedSize) ||
      dataStart + compressedSize > start ||
      u16(local + 6) !== flags ||
      u16(local + 8) !== method ||
      !nameBytes.every((byte, position) => byte === u8(local + 30 + position))
    )
      throw invalidContainer();
    if (
      (flags & 8) === 0 &&
      (u32(local + 14) !== crc ||
        u32(local + 18) !== compressedSize ||
        u32(local + 22) !== decodedSize)
    )
      throw invalidContainer();
    let rangeEnd = dataStart + compressedSize;
    if ((flags & 8) !== 0) {
      // Both legal descriptor forms are accepted. Sizes come from the checked
      // central directory, never from a scan for a signature inside model data.
      // The signed form is tried first; when the first word equals the
      // signature but the signed reading does not validate, the word is a
      // CRC-32 that happens to equal it and the unsigned reading is used.
      const validAt = (at: number): boolean =>
        fits(at, 12) &&
        at + 12 <= start &&
        u32(at) === crc &&
        u32(at + 4) === compressedSize &&
        u32(at + 8) === decodedSize;
      const descriptor =
        fits(rangeEnd, 4) &&
        u32(rangeEnd) === 0x08074b50 &&
        validAt(rangeEnd + 4)
          ? rangeEnd + 4
          : rangeEnd;
      if (!validAt(descriptor)) throw invalidContainer();
      rangeEnd = descriptor + 12;
    }
    ranges.push({ start: local, end: rangeEnd });
    if (name === "DataModel") {
      model = {
        start: dataStart,
        size: compressedSize,
        decodedSize,
        crc,
        method,
        encrypted: (flags & 0x41) !== 0,
      };
    }
    cursor += recordSize;
  }
  if (cursor !== end) throw invalidContainer();
  ranges.sort((a, b) => a.start - b.start);
  if (
    ranges.some(
      (range, index) => index > 0 && ranges[index - 1]!.end > range.start,
    )
  ) {
    throw invalidContainer();
  }
  // These are container markers, not claims about a file's connection mode.
  if (
    // Report parts moved between Desktop formats (Report/Layout in older
    // files, Report/definition/... in the enhanced format), so only the two
    // parts every sample carries are required.
    !["[Content_Types].xml", "Version"].every((name) => names.has(name))
  ) {
    throw invalidContainer();
  }
  return { model, peakBytes };
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) {
      crc = (crc >>> 1) ^ ((crc & 1) === 0 ? 0 : 0xedb88320);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * Read the stored DataModel ZIP part into an independent buffer. This low-level
 * reader does not decode XPress9, inspect the model catalog, or export tables.
 * No filesystem, runtime initialization, or network access is involved.
 */
export function readPbiModelPart(
  input: Uint8Array,
  options: PbiContainerOptions = {},
): Uint8Array {
  const limits = validateLimits(options);
  // A view whose buffer was transferred away (for example to a worker) has no
  // bytes to read. It is refused like any other unreadable input rather than
  // surfacing the runtime's own TypeError from the DataView constructor.
  if ((input.buffer as ArrayBuffer & { detached?: boolean }).detached === true)
    throw invalidContainer();
  checkLimit(limits, "inputBytes", input.byteLength, "container");
  checkLimit(
    limits,
    "peakBytes",
    backingBytes(input) + scratchBytes,
    "container",
  );
  const { model, peakBytes } = directory(input, limits);
  if (model === undefined) {
    throw new ConsultChimpsError(
      "PBI_NO_MODEL",
      "This Power BI container has no embedded model to export. Ask for a .pbix saved with imported data. Templates, live connections, and DirectQuery-only files may not contain imported rows.",
      { details: { stage: "container" } },
    );
  }
  if (model.encrypted) throw encryptedModel();
  if (model.size === 0) throw unreadableModel();
  // Power BI normally stores its already-compressed DataModel without another
  // ZIP compression layer. Refuse other methods until a bounded inflater exists:
  // JSZip's whole-buffer async API cannot enforce growth before allocation.
  if (model.method !== 0) throw unboundedContainer();
  if (model.size !== model.decodedSize) throw unreadableModel();
  checkLimit(limits, "decodedBytes", model.decodedSize, "model-part");
  checkLimit(limits, "peakBytes", peakBytes + model.decodedSize, "model-part");
  // The CRC is computed on the copy that is returned, so the guarantee holds
  // even when the caller's memory is shared and changes underneath the reader.
  const bytes = new Uint8Array(
    input.subarray(model.start, model.start + model.size),
  );
  if (crc32(bytes) !== model.crc) throw unreadableModel();
  return bytes;
}
