import { describe, expect, it, vi } from "vitest";
import type { ConsultChimpsError } from "@consultchimps/core";

/**
 * The decoder's two cached buffers. A `_x9_malloc` that returns zero must leave
 * no pointer behind: the old allocation has already been freed by then, and
 * `close()` handing that address to `_x9_free` a second time traps or corrupts
 * the allocator, burying the capacity refusal the caller should have seen.
 */

interface FakeModule {
  HEAPU8: Uint8Array;
  _x9_create: () => number;
  _x9_destroy: (context: number) => void;
  _x9_decompress: (
    context: number,
    source: number,
    sourceSize: number,
    destination: number,
    destinationSize: number,
  ) => number;
  _x9_last_error: (context: number) => number;
  _x9_malloc: (size: number) => number;
  _x9_free: (pointer: number) => void;
}

const freed: number[] = [];
const allocated: number[] = [];
/** Which allocation attempt returns zero, counted from one. */
let failAt = Number.MAX_SAFE_INTEGER;

vi.mock("../wasm/xpress9.mjs", () => {
  let next = 1000;
  const module: FakeModule = {
    HEAPU8: new Uint8Array(1 << 20),
    _x9_create: () => 1,
    _x9_destroy: () => undefined,
    _x9_decompress: (_c, _s, _ss, _d, size) => size,
    _x9_last_error: () => 0,
    _x9_malloc: (size) => {
      if (allocated.length + 1 === failAt) return 0;
      next += 4096;
      allocated.push(size);
      return next;
    },
    _x9_free: (pointer) => {
      freed.push(pointer);
    },
  };
  return { default: () => Promise.resolve(module) };
});

const { loadXpress9 } = await import("../src/xpress9/runtime.js");

function reset(fail: number): void {
  freed.length = 0;
  allocated.length = 0;
  failAt = fail;
}

/** Any valid module: the mocked factory ignores the bytes, validate does not. */
const wasmBinary = new Uint8Array([
  0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00,
]);

describe("growing a cached buffer", () => {
  it("leaves no pointer behind when the allocation fails", async () => {
    reset(Number.MAX_SAFE_INTEGER);
    const decoder = await loadXpress9({ wasmBinary });
    // A first frame allocates both buffers.
    decoder.decompress(new Uint8Array(16), 64);
    expect(allocated).toEqual([16, 64]);
    const beforeGrowth = freed.length;

    // The next frame needs a larger destination, so the old one is freed and
    // the replacement allocation fails.
    // The source buffer is already large enough, so the next allocation
    // attempt is the destination's replacement.
    failAt = allocated.length + 1;
    let failure: ConsultChimpsError | undefined;
    try {
      decoder.decompress(new Uint8Array(16), 1 << 30);
    } catch (error) {
      failure = error as ConsultChimpsError;
    }
    expect(failure).toBeDefined();
    expect(failure!.code).toBe("PBI_EXPORT_LIMIT_EXCEEDED");
    const releasedOnGrowth = freed.slice(beforeGrowth);
    expect(releasedOnGrowth).toHaveLength(1);

    // close() must not hand that address back a second time.
    decoder.close();
    const releasedOnClose = freed.slice(beforeGrowth + 1);
    expect(releasedOnClose).not.toContain(releasedOnGrowth[0]);
    expect(new Set(freed).size).toBe(freed.length);
  });

  it("frees both buffers exactly once on an ordinary close", async () => {
    reset(Number.MAX_SAFE_INTEGER);
    const decoder = await loadXpress9({ wasmBinary });
    decoder.decompress(new Uint8Array(16), 64);
    decoder.close();
    expect(freed).toHaveLength(2);
    expect(new Set(freed).size).toBe(2);
    // A second close releases nothing again.
    decoder.close();
    expect(freed).toHaveLength(2);
  });

  it("keeps a buffer that is already large enough", async () => {
    reset(Number.MAX_SAFE_INTEGER);
    const decoder = await loadXpress9({ wasmBinary });
    decoder.decompress(new Uint8Array(16), 64);
    decoder.decompress(new Uint8Array(8), 32);
    // Nothing was reallocated for the smaller frame.
    expect(allocated).toEqual([16, 64]);
    decoder.close();
  });
});
