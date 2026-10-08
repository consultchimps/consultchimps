import { describe, expect, it } from "vitest";

import { stagedFile, type StagedFileHandle } from "../src/split/staged-file.js";

/** A handle that takes at most `limit` bytes per write. */
function shortHandle(limit: number): {
  handle: StagedFileHandle;
  written: () => Uint8Array;
} {
  const parts: Uint8Array[] = [];
  return {
    handle: {
      write: (buffer, offset, length) => {
        const taken = Math.min(length, limit);
        parts.push(buffer.slice(offset, offset + taken));
        return Promise.resolve({ bytesWritten: taken });
      },
      close: () => Promise.resolve(),
    },
    written: () => Buffer.concat(parts),
  };
}

describe("stagedFile", () => {
  it("keeps writing until a short write has taken every byte", async () => {
    const { handle, written } = shortHandle(1000);
    const file = await stagedFile("out.xlsx", () => Promise.resolve(handle));
    const bytes = new Uint8Array(3 * 1024 * 1024).map(
      (_, index) => index % 251,
    );
    await file.write(bytes.subarray(0, 2 * 1024 * 1024));
    await file.write(bytes.subarray(2 * 1024 * 1024));
    await file.close();
    expect(Buffer.from(written()).equals(Buffer.from(bytes))).toBe(true);
  });

  it("fails rather than leave a short file when a write takes nothing", async () => {
    const { handle } = shortHandle(0);
    const file = await stagedFile("out.xlsx", () => Promise.resolve(handle));
    await file.write(new Uint8Array(10));
    await expect(file.close()).rejects.toThrow(/accepted no more data/u);
  });
});
