import { createHash } from "node:crypto";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { ConsultChimpsError } from "@consultchimps/core";
import { PipelineBudget, validateExportOptions } from "../src/budget.js";
import { readPbiModelPart } from "../src/container.js";
import {
  decompressModelPart,
  detectCompression,
  predictedHeapBytes,
} from "../src/xpress9/stream.js";
import { loadXpress9 } from "../src/xpress9/runtime.js";
import type { Xpress9Decoder } from "../src/xpress9/runtime.js";
import { readPbiTables } from "../src/pipeline.js";
import { FETCHED, FIXTURES } from "./oracle.js";

/** Section C, "XPress9 runtime and chunk framing". */

const container = new Uint8Array(
  readFileSync(path.join(FIXTURES, "a-2018-fuzzy.pbix")),
);
const modelPart = readPbiModelPart(container);

function budget(overrides: Record<string, number> = {}): PipelineBudget {
  return new PipelineBudget(validateExportOptions(overrides, true));
}

async function refusal(
  run: () => Promise<unknown>,
): Promise<ConsultChimpsError> {
  try {
    await run();
  } catch (error) {
    if (error instanceof ConsultChimpsError) return error;
    throw error;
  }
  throw new Error("expected a refusal");
}

const SIGNATURE_MULTITHREADED =
  "This backup was created using multithreaded XPrs9.";

/**
 * a-2018-fuzzy's model part is exactly one self-contained chunk, so a stream
 * carrying the multithreaded signature with one prefix thread and one main
 * thread, each holding that chunk, exercises the whole multithreaded framing
 * path through two genuinely independent decoder sessions.
 */
function syntheticMultithreaded(): Uint8Array {
  const view = new DataView(
    modelPart.buffer,
    modelPart.byteOffset,
    modelPart.byteLength,
  );
  const uncompressed = view.getUint32(102, true);
  const compressed = view.getUint32(106, true);
  const chunk = modelPart.subarray(110, 110 + compressed);
  const header = new Uint8Array(102);
  for (let index = 0; index < SIGNATURE_MULTITHREADED.length; index++) {
    const unit = SIGNATURE_MULTITHREADED.charCodeAt(index);
    header[index * 2] = unit & 0xff;
    header[index * 2 + 1] = unit >> 8;
  }
  const counts = new Uint8Array(40);
  const countsView = new DataView(counts.buffer);
  countsView.setBigUint64(0, 1n, true); // main chunks
  countsView.setBigUint64(8, 1n, true); // prefix chunks
  countsView.setBigUint64(16, 1n, true); // prefix threads
  countsView.setBigUint64(24, 1n, true); // main threads
  countsView.setBigUint64(32, BigInt(uncompressed), true);
  const body = new Uint8Array(8 + compressed);
  const bodyView = new DataView(body.buffer);
  bodyView.setUint32(0, uncompressed, true);
  bodyView.setUint32(4, compressed, true);
  body.set(chunk, 8);
  const stream = new Uint8Array(142 + body.length * 2);
  stream.set(header, 0);
  stream.set(counts, 102);
  stream.set(body, 142);
  stream.set(body, 142 + body.length);
  return stream;
}

describe("signature detection", () => {
  it("recognises the single-threaded signature the corpus carries", () => {
    expect(detectCompression(modelPart)).toBe("single");
  });

  it("recognises the multithreaded signature", () => {
    expect(detectCompression(syntheticMultithreaded())).toBe("multithreaded");
  });

  it("recognises an uncompressed backup image by its stream marker", () => {
    const marker = "STREAM_STORAGE_SIGNATURE_)!@#$%^&*(";
    const bytes = new Uint8Array(200);
    bytes[0] = 0xff;
    bytes[1] = 0xfe;
    for (let index = 0; index < marker.length; index++)
      bytes[2 + index * 2] = marker.charCodeAt(index);
    expect(detectCompression(bytes)).toBe("uncompressed");
  });

  it("recognises neither signature", () => {
    expect(detectCompression(new Uint8Array(200))).toBeNull();
    expect(detectCompression(new Uint8Array(4))).toBeNull();
  });
});

describe("chunk framing", () => {
  it("decodes the corpus stream to the independent reader's backup image", async () => {
    const stream = await decompressModelPart(modelPart, budget(), undefined);
    expect(stream.compression).toBe("single");
    expect(stream.chunks).toBe(1);
    expect(stream.bytes.byteLength).toBe(446_464);
    // The digest of the backup image pbixray extracted from the same file.
    expect(createHash("sha256").update(stream.bytes).digest("hex")).toBe(
      "57d6185cb0a6fe975638f7c62646a86694dc5a165b8bf1678a345abaa66bad2b",
    );
  }, 60_000);

  it("runs each multithreaded thread group on its own decoder", async () => {
    const synthetic = syntheticMultithreaded();
    const stream = await decompressModelPart(synthetic, budget(), undefined);
    expect(stream.compression).toBe("multithreaded");
    expect(stream.chunks).toBe(2);
    const single = await decompressModelPart(modelPart, budget(), undefined);
    // Both groups decode independently to the same backup image.
    expect(stream.bytes.byteLength).toBe(single.bytes.byteLength * 2);
    expect(
      Buffer.from(stream.bytes.subarray(0, single.bytes.byteLength)),
    ).toEqual(Buffer.from(single.bytes));
    expect(Buffer.from(stream.bytes.subarray(single.bytes.byteLength))).toEqual(
      Buffer.from(single.bytes),
    );
  }, 60_000);

  it("refuses a stream with neither signature", async () => {
    const error = await refusal(() =>
      decompressModelPart(new Uint8Array(200), budget(), undefined),
    );
    expect(error.code).toBe("PBI_MODEL_UNREADABLE");
    expect(error.details).toEqual({ stage: "xpress9" });
  });

  it("refuses a chunk header that runs past the model part", async () => {
    const truncated = modelPart.slice(0, modelPart.length - 1000);
    const error = await refusal(() =>
      decompressModelPart(truncated, budget(), undefined),
    );
    expect(error.code).toBe("PBI_MODEL_UNREADABLE");
  });

  it("refuses a framing walk that cannot advance", async () => {
    const zeroed = modelPart.slice();
    new DataView(zeroed.buffer).setUint32(102, 0, true);
    new DataView(zeroed.buffer).setUint32(106, 0, true);
    const error = await refusal(() =>
      decompressModelPart(zeroed, budget(), undefined),
    );
    expect(error.code).toBe("PBI_MODEL_UNREADABLE");
  });

  it("checks the decoded budget before each allocation", async () => {
    const error = await refusal(() =>
      decompressModelPart(modelPart, budget({ decodedBytes: 1000 }), undefined),
    );
    expect(error.code).toBe("PBI_EXPORT_LIMIT_EXCEEDED");
    expect(error.details).toMatchObject({
      stage: "xpress9",
      option: "decodedBytes",
      limit: 1000,
    });
  }, 60_000);

  it("checks the peak budget before each allocation", async () => {
    const error = await refusal(() =>
      decompressModelPart(
        modelPart,
        budget({ peakBytes: 1_000_000 }),
        undefined,
      ),
    );
    expect(error.code).toBe("PBI_EXPORT_LIMIT_EXCEEDED");
    expect(error.details).toMatchObject({
      stage: "xpress9",
      option: "peakBytes",
    });
  }, 60_000);
});

describe("the multithreaded warning", () => {
  it("is not raised for the single-threaded corpus stream", async () => {
    const model = await readPbiTables(container);
    expect(
      model.manifest.unverifiedPaths.some(
        (entry) => entry.code === "PBI_UNVERIFIED_XPRESS9_MULTITHREADED",
      ),
    ).toBe(false);
  }, 60_000);
});

describe("the runtime's own buffers are charged before it allocates them", () => {
  /** A single-chunk stream whose header declares `declared` output bytes. */
  function frame(declared: number, compressed = 16): Uint8Array {
    const signature = "This backup was created using XPress9 compression.";
    const stream = new Uint8Array(110 + compressed);
    for (let index = 0; index < signature.length; index++) {
      const unit = signature.charCodeAt(index);
      stream[index * 2] = unit & 0xff;
      stream[index * 2 + 1] = unit >> 8;
    }
    const view = new DataView(stream.buffer);
    view.setUint32(102, declared, true);
    view.setUint32(106, compressed, true);
    return stream;
  }

  /**
   * A decoder that records whether the runtime was ever asked to decompress,
   * and reports a heap that grows the way Emscripten's would.
   */
  function spy(heap = 16 * 1024 * 1024): {
    load: () => Promise<Xpress9Decoder>;
    calls: { input: number; outputSize: number }[];
  } {
    const calls: { input: number; outputSize: number }[] = [];
    let current = heap;
    return {
      calls,
      load: () =>
        Promise.resolve({
          decompress(input: Uint8Array, outputSize: number): Uint8Array {
            calls.push({ input: input.byteLength, outputSize });
            const needed = heap + input.byteLength + outputSize;
            if (needed > current)
              current = Math.ceil((needed * 1.2) / (64 * 1024)) * (64 * 1024);
            return new Uint8Array(outputSize);
          },
          memoryBytes: (): number => current,
          close(): void {},
        }),
    };
  }

  it("refuses a frame whose declared output passes peakBytes, before decompressing", async () => {
    const { load, calls } = spy();
    const error = await refusal(() =>
      decompressModelPart(
        frame(1_500_000_000),
        budget({ peakBytes: 64 * 1024 * 1024 }),
        undefined,
        load,
      ),
    );
    expect(error.code).toBe("PBI_EXPORT_LIMIT_EXCEEDED");
    expect(error.details).toMatchObject({
      stage: "xpress9",
      option: "peakBytes",
      limit: 64 * 1024 * 1024,
    });
    // The whole point: the runtime was never asked to allocate the buffer the
    // refusal describes. Charging a fixed corpus figure let this through.
    expect(calls).toEqual([]);
  });

  it("charges the destination buffer a frame actually needs", async () => {
    const { load, calls } = spy();
    const stream = await decompressModelPart(
      frame(4_000_000),
      budget(),
      undefined,
      load,
    );
    expect(stream.bytes.byteLength).toBe(4_000_000);
    expect(calls).toEqual([{ input: 16, outputSize: 4_000_000 }]);
  });

  it("charges the predicted heap, not the payload sum", async () => {
    const { load } = spy();
    const single = budget();
    const stream = await decompressModelPart(
      frame(2_000_000),
      single,
      undefined,
      load,
    );
    const predicted = predictedHeapBytes(
      16 * 1024 * 1024,
      16,
      2_000_000,
      16 * 1024 * 1024,
    );
    expect(stream.reservedHeapBytes).toBe(predicted);
    expect(single.terms().xpress9LinearMemory).toBe(predicted);
    // Well above the payload sum, which is what made a peakBytes between the
    // two figures let an over-limit allocation through.
    expect(predicted).toBeGreaterThan(16 * 1024 * 1024 + 2_000_000 + 16);
    expect(stream.heapUnderpredictions).toBe(0);
  });
});

describe("predictedHeapBytes", () => {
  const baseline = 16 * 1024 * 1024;
  const page = 64 * 1024;

  it("returns whole WebAssembly pages", () => {
    for (const [source, destination] of [
      [1, 1],
      [12_345, 67_891],
      [1_000_003, 7],
    ] as const) {
      const predicted = predictedHeapBytes(
        baseline,
        source,
        destination,
        baseline,
      );
      expect(predicted % page).toBe(0);
    }
  });

  it("adds the geometric growth step on top of what is needed", () => {
    // A payload far beyond the current heap must grow it, so the prediction is
    // the needed figure plus a fifth, rounded up to a page.
    const destination = 64 * 1024 * 1024;
    const predicted = predictedHeapBytes(baseline, 0, destination, baseline);
    // Two generations of the buffer, because a regrown one leaves the block it
    // replaced behind, plus the allocator's own slack.
    const needed = baseline + 2 * destination + 64 * 1024 + destination * 0.01;
    expect(predicted).toBeGreaterThanOrEqual(Math.ceil(needed * 1.2));
    expect(predicted).toBeLessThan(Math.ceil(needed * 1.2) + page);
  });

  it("never predicts below the heap the module already holds", () => {
    // A small frame against a large heap does not shrink anything.
    const current = 128 * 1024 * 1024;
    expect(predictedHeapBytes(baseline, 16, 1024, current)).toBe(current);
  });

  it("carries allocator slack for the buffers themselves", () => {
    // Two identical payloads, one split across both buffers: the slack term is
    // proportional, so neither prediction is merely the payload sum.
    const predicted = predictedHeapBytes(baseline, 0, 0, 0);
    expect(predicted).toBeGreaterThan(baseline);
  });
});

describe("the prediction against the real runtime", () => {
  const files = [path.join(FIXTURES, "a-2018-fuzzy.pbix")];
  for (const name of [
    "c-perf-analyzer.pbix",
    "g-2018-newformat.pbix",
    "l-2019-12.pbix",
    "m-2020-09.pbix",
    "n-2020-11.pbix",
  ]) {
    const candidate = path.join(FETCHED, name);
    if (existsSync(candidate)) files.push(candidate);
  }

  for (const file of files) {
    it(`bounds the heap on ${path.basename(file)}`, async () => {
      const running = budget({ peakBytes: 2 * 1024 * 1024 * 1024 });
      const stream = await decompressModelPart(
        readPbiModelPart(new Uint8Array(readFileSync(file))),
        running,
        undefined,
      );
      // The reservation is an upper bound on what the module actually took,
      // checked after every frame. A nonzero count means it was not.
      expect(stream.heapUnderpredictions).toBe(0);
      expect(stream.reservedHeapBytes).toBeGreaterThan(0);
    }, 300_000);
  }
});

describe("repeated regrowth of both cached buffers", () => {
  it("stays inside the prediction across many reallocations", async () => {
    // The review's probe: grow both buffers ten percent at a time, in the
    // runtime's free-then-allocate order, many steps. Every freed generation
    // stays in the heap because dlmalloc cannot serve the larger request from
    // the smaller block it just released, so a single-generation bound was
    // exceeded. The decode itself is irrelevant here; the allocation happens
    // first, and a damaged-model refusal after it is expected.
    const decoder = await loadXpress9();
    try {
      let sourceCapacity = 0;
      let destinationCapacity = 0;
      let worst = 0;
      let source = 4096;
      let destination = 1024 * 1024;
      for (let step = 0; step < 40; step++) {
        // Mirror the runtime's doubling, which is what the prediction reads.
        if (source > sourceCapacity)
          sourceCapacity = Math.max(source, sourceCapacity * 2);
        if (destination > destinationCapacity)
          destinationCapacity = Math.max(destination, destinationCapacity * 2);
        try {
          decoder.decompress(new Uint8Array(source), destination);
        } catch {
          // The bytes are not a real frame, so the decode refuses. The buffers
          // were already allocated by then, which is what this measures.
        }
        const predicted = predictedHeapBytes(
          16 * 1024 * 1024,
          sourceCapacity,
          destinationCapacity,
          16 * 1024 * 1024,
        );
        const actual = decoder.memoryBytes();
        worst = Math.max(worst, actual / predicted);
        expect(
          actual,
          `step ${step}: source ${sourceCapacity}, destination ${destinationCapacity}`,
        ).toBeLessThanOrEqual(predicted);
        source = Math.ceil(source * 1.1);
        destination = Math.ceil(destination * 1.1);
      }
      // The bound holds with room to spare rather than by a hair.
      expect(worst).toBeLessThan(1);
    } finally {
      decoder.close();
    }
  }, 300_000);
});
