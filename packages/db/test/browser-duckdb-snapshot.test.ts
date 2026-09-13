import type { RandomAccessFile } from "@consultchimps/core";
import { expect, test, vi } from "vitest";

import { exportReadonlyDuckDbSnapshot } from "../src/browser-duckdb-snapshot.js";

const destination: RandomAccessFile = {
  name: "export.duckdb",
  size: 0,
  async readAt() {
    return new Uint8Array();
  },
  async writeAt() {},
  async truncate() {},
  async close() {},
};

test("copies through a prepared logical snapshot and removes it after release", async () => {
  const calls: string[] = [];

  await expect(
    exportReadonlyDuckDbSnapshot({
      destination,
      async allocate() {
        calls.push("allocate");
        return {
          name: ".consultchimps-export-snapshot-success.duckdb",
          path: "/success.duckdb",
          directory: "test-snapshots",
          async copyTo(received) {
            expect(received).toBe(destination);
            calls.push("copy-destination");
            return 4096;
          },
          async remove() {
            calls.push("remove");
          },
        };
      },
      async prepare() {
        calls.push("prepare");
      },
      async release() {
        calls.push("release");
      },
    }),
  ).resolves.toBe(4096);
  expect(calls).toEqual([
    "allocate",
    "prepare",
    "copy-destination",
    "release",
    "remove",
  ]);
});

test("observes cancellation after logical snapshot preparation before touching the destination", async () => {
  const controller = new AbortController();
  const copyTo = vi.fn(async () => 0);
  const release = vi.fn(async () => undefined);
  const remove = vi.fn(async () => undefined);

  await expect(
    exportReadonlyDuckDbSnapshot({
      destination,
      signal: controller.signal,
      async allocate() {
        return {
          name: ".consultchimps-export-snapshot-cancel.duckdb",
          path: "/cancel.duckdb",
          directory: "test-snapshots",
          copyTo,
          remove,
        };
      },
      async prepare() {
        controller.abort();
      },
      release,
    }),
  ).rejects.toMatchObject({ code: "OPERATION_ABORTED" });
  expect(copyTo).not.toHaveBeenCalled();
  expect(release).toHaveBeenCalledOnce();
  expect(remove).toHaveBeenCalledOnce();
});

test("rethrows an export failure after releasing snapshot storage", async () => {
  const failure = new Error("Injected destination failure");
  const release = vi.fn(async () => undefined);
  const remove = vi.fn(async () => undefined);

  await expect(
    exportReadonlyDuckDbSnapshot({
      destination,
      async allocate() {
        return {
          name: ".consultchimps-export-snapshot-failure.duckdb",
          path: "/failure.duckdb",
          directory: "test-snapshots",
          async copyTo() {
            throw failure;
          },
          remove,
        };
      },
      async prepare() {},
      release,
    }),
  ).rejects.toBe(failure);
  expect(release).toHaveBeenCalledOnce();
  expect(remove).toHaveBeenCalledOnce();
});

test("retains and names a snapshot when registered handles cannot be released", async () => {
  const remove = vi.fn(async () => undefined);

  await expect(
    exportReadonlyDuckDbSnapshot({
      destination,
      async allocate() {
        return {
          name: ".consultchimps-export-snapshot-retained.duckdb",
          path: "/retained.duckdb",
          directory: "test-snapshots",
          async copyTo() {
            return 4096;
          },
          remove,
        };
      },
      async prepare() {},
      async release() {
        throw new Error("Injected release failure");
      },
    }),
  ).rejects.toMatchObject({
    code: "DB_BROWSER_DUCKDB_SNAPSHOT_CLEANUP_REQUIRED",
    details: {
      directory: "test-snapshots",
      snapshotName: ".consultchimps-export-snapshot-retained.duckdb",
    },
  });
  expect(remove).not.toHaveBeenCalled();
});
