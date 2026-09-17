import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { ConsultChimpsError } from "@consultchimps/core";
import { PipelineBudget, validateExportOptions } from "../src/budget.js";
import { readPbiModelPart } from "../src/container.js";
import {
  decompressModelPart,
  detectCompression,
} from "../src/xpress9/stream.js";
import { readPbiTables } from "../src/pipeline.js";
import { FIXTURES } from "./oracle.js";

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
