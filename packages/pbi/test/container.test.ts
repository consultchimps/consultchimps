import JSZip from "jszip";
import { crc32 } from "node:zlib";
import { describe, expect, it } from "vitest";
import { ConsultChimpsError } from "@consultchimps/core";
import { readPbiModelPart } from "../src/index.js";
import type { PbiContainerOptions } from "../src/index.js";

const confidential = "CONFIDENTIAL_SYNTHETIC_MARKER";
const model = new TextEncoder().encode(confidential);

async function fixture(
  options: {
    model?: Uint8Array | null;
    compression?: "STORE" | "DEFLATE";
    streamFiles?: boolean;
    connections?: boolean;
    omit?: string;
    comment?: string;
  } = {},
): Promise<Uint8Array> {
  const zip = new JSZip();
  for (const [name, data] of [
    ["[Content_Types].xml", "<Types />"],
    ["Version", "1.0"],
    ["Report/Layout", "{}"],
  ]) {
    if (name !== options.omit) zip.file(name!, data!);
  }
  if (options.model !== null) zip.file("DataModel", options.model ?? model);
  if (options.connections) zip.file("Connections", confidential);
  return zip.generateAsync({
    type: "uint8array",
    compression: options.compression ?? "STORE",
    streamFiles: options.streamFiles ?? false,
    ...(options.comment === undefined ? {} : { comment: options.comment }),
  });
}

/** Add `delta` to every directory offset at or beyond `at` (no archive comment). */
function shiftOffsets(bytes: Uint8Array, at: number, delta: number): void {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocd = bytes.length - 22;
  const count = view.getUint16(eocd + 10, true);
  let start = view.getUint32(eocd + 16, true);
  if (start >= at) {
    start += delta;
    view.setUint32(eocd + 16, start, true);
  }
  let cursor = start;
  for (let index = 0; index < count; index++) {
    const local = view.getUint32(cursor + 42, true);
    if (local >= at) view.setUint32(cursor + 42, local + delta, true);
    cursor +=
      46 +
      view.getUint16(cursor + 28, true) +
      view.getUint16(cursor + 30, true) +
      view.getUint16(cursor + 32, true);
  }
}

function localOf(
  input: Uint8Array,
  name: string,
): { local: number; data: number } {
  const view = new DataView(input.buffer, input.byteOffset, input.byteLength);
  const local = view.getUint32(central(input, name) + 42, true);
  const data =
    local +
    30 +
    view.getUint16(local + 26, true) +
    view.getUint16(local + 28, true);
  return { local, data };
}

function central(input: Uint8Array, name: string): number {
  const view = new DataView(input.buffer, input.byteOffset, input.byteLength);
  for (let offset = 0; offset <= input.byteLength - 46; offset++) {
    if (view.getUint32(offset, true) !== 0x02014b50) continue;
    const length = view.getUint16(offset + 28, true);
    if (
      new TextDecoder().decode(
        input.subarray(offset + 46, offset + 46 + length),
      ) === name
    )
      return offset;
  }
  throw new Error("Fixture entry missing");
}

function diagnostic(work: () => unknown): ConsultChimpsError {
  try {
    work();
  } catch (error) {
    expect(error).toBeInstanceOf(ConsultChimpsError);
    const result = error as ConsultChimpsError;
    expect(result.cause).toBeUndefined();
    expect(
      JSON.stringify({ ...result, message: result.message }),
    ).not.toContain(confidential);
    return result;
  }
  throw new Error("Expected refusal");
}

describe("readPbiModelPart", () => {
  it.each([false, true])(
    "reads a stored model with streamed ZIP = %s",
    async (streamFiles) => {
      const input = await fixture({ streamFiles, connections: true });
      const original = input.slice();
      const output = readPbiModelPart(input);
      expect(output).toEqual(model);
      expect(readPbiModelPart(input)).toEqual(output);
      output.fill(0);
      expect(input).toEqual(original);
      const buffer = Buffer.from(input);
      readPbiModelPart(buffer).fill(0);
      expect(buffer).toEqual(Buffer.from(original));
    },
  );

  it("reads an offset view and leaves unrelated backing bytes untouched", async () => {
    const input = await fixture();
    const backing = new Uint8Array(input.length + 30).fill(7);
    backing.set(input, 10);
    expect(readPbiModelPart(backing.subarray(10, 10 + input.length))).toEqual(
      model,
    );
    expect(backing[0]).toBe(7);
    expect(backing.at(-1)).toBe(7);
  });

  it("uses one no-model code even with a Connections part", async () => {
    const input = await fixture({ model: null, connections: true });
    expect(diagnostic(() => readPbiModelPart(input)).code).toBe("PBI_NO_MODEL");
  });

  it("reads enhanced-format containers that have no Report/Layout part", async () => {
    const zip = new JSZip();
    zip.file("Version", "1.0");
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("Report/definition/report.json", "{}");
    zip.file("Report/definition/pages/pages.json", "{}");
    zip.file("Connections", "{}");
    zip.file("DataModel", model);
    const input = await zip.generateAsync({
      type: "uint8array",
      compression: "STORE",
    });
    expect(readPbiModelPart(input)).toEqual(model);
  });

  it.each(["[Content_Types].xml", "Version"])(
    "requires the %s marker",
    async (omit) => {
      const input = await fixture({ omit });
      expect(diagnostic(() => readPbiModelPart(input)).code).toBe(
        "PBI_INVALID_CONTAINER",
      );
    },
  );

  it("refuses non-ZIP and truncated containers without leaking parser text", async () => {
    const input = await fixture();
    const transferred = input.slice();
    structuredClone(transferred.buffer, { transfer: [transferred.buffer] });
    for (const bytes of [
      model,
      new Uint8Array(),
      input.subarray(0, input.length - 1),
      transferred,
    ]) {
      expect(diagnostic(() => readPbiModelPart(bytes)).code).toBe(
        "PBI_INVALID_CONTAINER",
      );
    }
  });

  it("refuses an empty model and a model with a failed CRC", async () => {
    const empty = await fixture({ model: new Uint8Array() });
    expect(diagnostic(() => readPbiModelPart(empty)).code).toBe(
      "PBI_MODEL_UNREADABLE",
    );
    const input = await fixture();
    const view = new DataView(input.buffer);
    const entry = central(input, "DataModel");
    const local = view.getUint32(entry + 42, true);
    input[local + 30 + view.getUint16(local + 26, true)]! ^= 1;
    expect(diagnostic(() => readPbiModelPart(input)).code).toBe(
      "PBI_MODEL_UNREADABLE",
    );
  });

  it("refuses ZIP-encrypted models before reading their bytes", async () => {
    const input = await fixture();
    const view = new DataView(input.buffer);
    const entry = central(input, "DataModel");
    const local = view.getUint32(entry + 42, true);
    view.setUint16(entry + 8, 1, true);
    view.setUint16(local + 6, 1, true);
    expect(diagnostic(() => readPbiModelPart(input)).code).toBe(
      "PBI_MODEL_ENCRYPTED",
    );
  });

  it("refuses unsupported compression rather than attempting unbounded inflation", async () => {
    const input = await fixture({ compression: "DEFLATE" });
    expect(diagnostic(() => readPbiModelPart(input)).code).toBe(
      "PBI_EXPORT_LIMIT_EXCEEDED",
    );
  });

  it("refuses ZIP64 and multi-volume layouts", async () => {
    for (const offset of [4, 10]) {
      const input = await fixture();
      new DataView(input.buffer).setUint16(
        input.length - 22 + offset,
        0xffff,
        true,
      );
      expect(diagnostic(() => readPbiModelPart(input)).code).toBe(
        "PBI_EXPORT_LIMIT_EXCEEDED",
      );
    }
  });

  it("checks local names and directory ranges before reading model bytes", async () => {
    const input = await fixture();
    const view = new DataView(input.buffer);
    const entry = central(input, "DataModel");
    const local = view.getUint32(entry + 42, true);
    input[local + 30] = 0;
    expect(diagnostic(() => readPbiModelPart(input)).code).toBe(
      "PBI_INVALID_CONTAINER",
    );
    view.setUint32(entry + 42, input.length, true);
    expect(diagnostic(() => readPbiModelPart(input)).code).toBe(
      "PBI_INVALID_CONTAINER",
    );
  });

  it.each(["../DataModel", "./DataModel", "/DataModel", "a\\DataModel"])(
    "refuses the path alias %s in an otherwise valid container",
    async (alias) => {
      const zip = new JSZip();
      zip.file("[Content_Types].xml", "<Types />");
      zip.file("Version", "1.0");
      zip.file(alias, model, { createFolders: false });
      const input = await zip.generateAsync({
        type: "uint8array",
        compression: "STORE",
      });
      expect(diagnostic(() => readPbiModelPart(input)).code).toBe(
        "PBI_INVALID_CONTAINER",
      );
    },
  );

  it("refuses duplicate names and duplicate local offsets", async () => {
    const zip = new JSZip();
    zip.file("[Content_Types].xml", "<Types />");
    zip.file("Version", "1.0");
    zip.file("DataModel", model);
    zip.file("DataModeX", model);
    const bytes = await zip.generateAsync({
      type: "uint8array",
      compression: "STORE",
    });
    const view = new DataView(bytes.buffer);
    const entry = central(bytes, "DataModeX");
    bytes[entry + 46 + 8] = 108;
    bytes[view.getUint32(entry + 42, true) + 30 + 8] = 108;
    expect(diagnostic(() => readPbiModelPart(bytes)).code).toBe(
      "PBI_INVALID_CONTAINER",
    );
    const shared = await fixture();
    const sharedView = new DataView(shared.buffer);
    // DataModel's record comes after Version's, so the duplicate-offset check
    // is the one that fires; the name comparison would not reach it first.
    sharedView.setUint32(
      central(shared, "DataModel") + 42,
      sharedView.getUint32(central(shared, "Version") + 42, true),
      true,
    );
    expect(diagnostic(() => readPbiModelPart(shared)).code).toBe(
      "PBI_INVALID_CONTAINER",
    );
  });

  it("finds the real end-of-directory record behind a decoy in the archive comment", async () => {
    // A 22-byte comment that is itself a syntactically complete record whose
    // own comment length (0) also reaches the end of input.
    const decoy = "PK\u0005\u0006" + "\0".repeat(18);
    const input = await fixture({ comment: decoy });
    expect(readPbiModelPart(input)).toEqual(model);
    // A decoy carrying a multi-disk field is skipped the same way.
    new DataView(input.buffer).setUint16(input.length - 22 + 4, 1, true);
    expect(readPbiModelPart(input)).toEqual(model);
  });

  it("accepts a data descriptor without its optional signature", async () => {
    const streamed = await fixture({ streamFiles: true });
    const { data } = localOf(streamed, "DataModel");
    const descriptor = data + model.length;
    expect(new DataView(streamed.buffer).getUint32(descriptor, true)).toBe(
      0x08074b50,
    );
    const input = new Uint8Array(streamed.length - 4);
    input.set(streamed.subarray(0, descriptor));
    input.set(streamed.subarray(descriptor + 4), descriptor);
    shiftOffsets(input, descriptor, -4);
    expect(readPbiModelPart(input)).toEqual(model);
    new DataView(input.buffer).setUint32(
      descriptor + 4,
      model.length + 1,
      true,
    );
    expect(diagnostic(() => readPbiModelPart(input)).code).toBe(
      "PBI_INVALID_CONTAINER",
    );
  });

  it("reads an unsigned descriptor whose CRC equals the signature value", async () => {
    // Four trailing bytes force the model's CRC-32 to the descriptor signature,
    // so the first word of an unsigned descriptor is indistinguishable from a
    // signed one until the signed reading fails to validate.
    const table = Array.from({ length: 256 }, (_, n) => {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      return c >>> 0;
    });
    const target = 0x08074b50;
    const indexes: number[] = [];
    let state = (target ^ 0xffffffff) >>> 0;
    for (let i = 0; i < 4; i++) {
      const index = table.findIndex((t) => t >>> 24 === state >>> 24);
      indexes.unshift(index);
      state = ((state ^ table[index]!) << 8) >>> 0;
    }
    const collide = new Uint8Array(model.length + 4);
    collide.set(model);
    state = (crc32(model) ^ 0xffffffff) >>> 0;
    for (const [i, index] of indexes.entries()) {
      collide[model.length + i] = (state ^ index) & 0xff;
      state = ((state >>> 8) ^ table[index]!) >>> 0;
    }
    expect(crc32(collide)).toBe(target);
    const streamed = await fixture({ model: collide, streamFiles: true });
    const { data } = localOf(streamed, "DataModel");
    const descriptor = data + collide.length;
    const input = new Uint8Array(streamed.length - 4);
    input.set(streamed.subarray(0, descriptor));
    input.set(streamed.subarray(descriptor + 4), descriptor);
    shiftOffsets(input, descriptor, -4);
    expect(new DataView(input.buffer).getUint32(descriptor, true)).toBe(target);
    expect(readPbiModelPart(input)).toEqual(collide);
    expect(readPbiModelPart(streamed)).toEqual(collide);
  });

  it("refuses an entry whose bytes lie inside another entry's data", async () => {
    const big = new Uint8Array(256).fill(1);
    const input = await fixture({ model: big });
    const version = localOf(input, "Version");
    const versionBytes = input.subarray(version.local, version.data + 3);
    const { data } = localOf(input, "DataModel");
    input.set(versionBytes, data);
    new DataView(input.buffer).setUint32(
      central(input, "Version") + 42,
      data,
      true,
    );
    expect(diagnostic(() => readPbiModelPart(input)).code).toBe(
      "PBI_INVALID_CONTAINER",
    );
  });

  it.each([
    ["compressed size", 20, 0xffffffff, "PBI_EXPORT_LIMIT_EXCEEDED"],
    ["decoded size", 24, 0xffffffff, "PBI_EXPORT_LIMIT_EXCEEDED"],
    ["local offset", 42, 0xffffffff, "PBI_EXPORT_LIMIT_EXCEEDED"],
  ] as const)(
    "refuses a ZIP64 sentinel in the entry %s",
    async (_label, offset, value, code) => {
      const input = await fixture();
      new DataView(input.buffer).setUint32(
        central(input, "DataModel") + offset,
        value,
        true,
      );
      expect(diagnostic(() => readPbiModelPart(input)).code).toBe(code);
    },
  );

  it("refuses a nonzero entry disk number, a zero entry count, and STORE size mismatches", async () => {
    const disk = await fixture();
    new DataView(disk.buffer).setUint16(
      central(disk, "DataModel") + 34,
      1,
      true,
    );
    expect(diagnostic(() => readPbiModelPart(disk)).code).toBe(
      "PBI_EXPORT_LIMIT_EXCEEDED",
    );
    const empty = await fixture();
    const eocd = empty.length - 22;
    new DataView(empty.buffer).setUint16(eocd + 8, 0, true);
    new DataView(empty.buffer).setUint16(eocd + 10, 0, true);
    expect(diagnostic(() => readPbiModelPart(empty)).code).toBe(
      "PBI_INVALID_CONTAINER",
    );
    const mismatch = await fixture();
    const view = new DataView(mismatch.buffer);
    const entry = central(mismatch, "DataModel");
    const { local } = localOf(mismatch, "DataModel");
    view.setUint32(entry + 24, model.length + 1, true);
    view.setUint32(local + 22, model.length + 1, true);
    expect(diagnostic(() => readPbiModelPart(mismatch)).code).toBe(
      "PBI_MODEL_UNREADABLE",
    );
  });

  it("tolerates leading bytes when the directory offsets are file-absolute", async () => {
    const plain = await fixture();
    const input = new Uint8Array(plain.length + 16).fill(0x5a);
    input.set(plain, 16);
    shiftOffsets(input, 0, 16);
    expect(readPbiModelPart(input)).toEqual(model);
  });

  it("counts a resizable buffer's maximum size toward the peak estimate", async () => {
    const plain = await fixture();
    const Resizable = ArrayBuffer as unknown as new (
      length: number,
      options: { maxByteLength: number },
    ) => ArrayBuffer;
    const buffer = new Resizable(plain.length, {
      maxByteLength: plain.length * 4,
    });
    const input = new Uint8Array(buffer);
    input.set(plain);
    let bound = 1;
    for (;;) {
      try {
        readPbiModelPart(plain, { peakBytes: bound });
        break;
      } catch (error) {
        bound = (error as ConsultChimpsError).details!.required as number;
      }
    }
    expect(
      diagnostic(() => readPbiModelPart(input, { peakBytes: bound })).code,
    ).toBe("PBI_EXPORT_LIMIT_EXCEEDED");
    expect(
      readPbiModelPart(input, { peakBytes: bound + plain.length * 3 }),
    ).toEqual(model);
  });

  it("accepts exact inclusive input and decoded bounds, refuses one byte less", async () => {
    const input = await fixture();
    expect(
      readPbiModelPart(input, {
        inputBytes: input.length,
        decodedBytes: model.length,
      }),
    ).toEqual(model);
    for (const options of [
      { inputBytes: input.length - 1 },
      { decodedBytes: model.length - 1 },
    ]) {
      expect(diagnostic(() => readPbiModelPart(input, options)).code).toBe(
        "PBI_EXPORT_LIMIT_EXCEEDED",
      );
    }
  });

  it("bounds peak memory including a view's whole backing buffer", async () => {
    const input = await fixture();
    // The final reader bound is reported before the model output is allocated.
    let bound = 1;
    for (;;) {
      try {
        readPbiModelPart(input, { peakBytes: bound });
        break;
      } catch (error) {
        expect(error).toMatchObject({ code: "PBI_EXPORT_LIMIT_EXCEEDED" });
        bound = (error as ConsultChimpsError).details!.required as number;
      }
    }
    expect(readPbiModelPart(input, { peakBytes: bound })).toEqual(model);
    expect(
      diagnostic(() => readPbiModelPart(input, { peakBytes: bound - 1 })).code,
    ).toBe("PBI_EXPORT_LIMIT_EXCEEDED");
    const backing = new Uint8Array(bound * 2);
    backing.set(input);
    expect(
      diagnostic(() =>
        readPbiModelPart(backing.subarray(0, input.length), {
          peakBytes: bound,
        }),
      ).code,
    ).toBe("PBI_EXPORT_LIMIT_EXCEEDED");
  });

  it.each([
    0,
    -1,
    0.5,
    Number.MAX_SAFE_INTEGER + 1,
    NaN,
    Infinity,
    "secret",
    null,
    false,
  ])("aggregates invalid limits before accessing input: %s", (value) => {
    const input = new Proxy(new Uint8Array(), {
      get() {
        throw new Error("Input was read");
      },
    });
    const options = {
      inputBytes: value,
      decodedBytes: value,
      peakBytes: value,
    } as PbiContainerOptions;
    const error = diagnostic(() => readPbiModelPart(input, options));
    expect(error.code).toBe("PBI_INVALID_OPTIONS");
    expect(error.details).toEqual({
      stage: "options",
      invalidOptions: ["inputBytes", "decodedBytes", "peakBytes"].map(
        (path) => ({ path, requirement: "a positive safe-integer byte count" }),
      ),
    });
  });
});
