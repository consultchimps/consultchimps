import {
  mkdtemp,
  readFile,
  readdir,
  rename,
  rm,
  writeFile,
} from "node:fs/promises";
import type { open as openFile, rm as remove } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { afterEach, expect, test, vi } from "vitest";

const filesystemFailure = vi.hoisted(() => ({
  closeCalls: 0,
  closeFailures: 0,
  openGate: undefined as Promise<void> | undefined,
  openStarted: undefined as (() => void) | undefined,
  removeCalls: 0,
  removeFailures: 0,
}));

vi.mock("node:fs/promises", async (importOriginal) => {
  const actual = (await importOriginal()) as Record<string, unknown> & {
    open: typeof openFile;
    rm: typeof remove;
  };
  return {
    ...actual,
    async open(...args: Parameters<typeof actual.open>) {
      const filePath = String(args[0]);
      if (path.basename(filePath).startsWith("part-")) {
        filesystemFailure.openStarted?.();
        await filesystemFailure.openGate;
      }
      const handle = await actual.open(...args);
      const close = handle.close.bind(handle);
      handle.close = async () => {
        filesystemFailure.closeCalls += 1;
        if (filesystemFailure.closeFailures > 0) {
          filesystemFailure.closeFailures -= 1;
          throw new Error("synthetic scratch close failure");
        }
        await close();
      };
      return handle;
    },
    async rm(...args: Parameters<typeof actual.rm>) {
      const target = String(args[0]);
      if (path.basename(target).startsWith("cc-scratch-")) {
        filesystemFailure.removeCalls += 1;
        if (filesystemFailure.removeFailures > 0) {
          filesystemFailure.removeFailures -= 1;
          throw new Error("synthetic scratch removal failure");
        }
      }
      return actual.rm(...args);
    },
  };
});

import {
  createScratchDirectory,
  openRandomAccessSource,
} from "../src/index.js";

const directories: string[] = [];

async function fixture(): Promise<string> {
  const directory = await mkdtemp(path.join(os.tmpdir(), "cc-file-test-"));
  directories.push(directory);
  return directory;
}

function deferred(): { readonly promise: Promise<void>; resolve(): void } {
  let resolve!: () => void;
  const promise = new Promise<void>((done) => {
    resolve = done;
  });
  return { promise, resolve };
}

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
  filesystemFailure.closeCalls = 0;
  filesystemFailure.closeFailures = 0;
  filesystemFailure.openGate = undefined;
  filesystemFailure.openStarted = undefined;
  filesystemFailure.removeCalls = 0;
  filesystemFailure.removeFailures = 0;
});

test("reads file ranges without changing the source", async () => {
  const directory = await fixture();
  const input = path.join(directory, "source.xlsx");
  await writeFile(input, "abcdef");
  const source = await openRandomAccessSource(input);
  try {
    expect(source.name).toBe("source.xlsx");
    expect(source.size).toBe(6);
    expect(new TextDecoder().decode(await source.readAt(2, 3))).toBe("cde");
    expect(new TextDecoder().decode(await source.readAt(5, 3))).toBe("f");
    expect(await source.readAt(20, 2)).toHaveLength(0);
    await expect(source.readAt(-1, 2)).rejects.toMatchObject({
      code: "FILES_INVALID_RANGE",
    });
    await expect(
      source.readAt(0, 2, AbortSignal.abort()),
    ).rejects.toMatchObject({ code: "OPERATION_ABORTED" });
    await source.verifyUnchanged();
  } finally {
    await source.close();
  }
  expect(await readFile(input, "utf8")).toBe("abcdef");
});

test("detects edits and replacement of a captured source", async () => {
  const directory = await fixture();
  const input = path.join(directory, "source.xlsx");
  await writeFile(input, "before");
  const source = await openRandomAccessSource(input);
  try {
    await writeFile(input, "after");
    await expect(source.verifyUnchanged()).rejects.toMatchObject({
      code: "FILES_SOURCE_CHANGED",
    });
  } finally {
    await source.close();
  }
  const replacementSource = await openRandomAccessSource(input);
  try {
    await rename(input, path.join(directory, "moved.xlsx"));
    await writeFile(input, "after");
    await expect(replacementSource.verifyUnchanged()).rejects.toMatchObject({
      code: "FILES_SOURCE_CHANGED",
    });
  } finally {
    await replacementSource.close();
  }
});

test("temporary random-access storage cleans only its owned files", async () => {
  const directory = await fixture();
  const retained = path.join(directory, "keep.txt");
  await writeFile(retained, "keep");
  const scratch = await createScratchDirectory(directory);
  try {
    const file = await scratch.create();
    await file.writeAt(0, new TextEncoder().encode("abcdef"));
    await file.writeAt(2, new TextEncoder().encode("XY"));
    expect(file.size).toBe(6);
    expect(new TextDecoder().decode(await file.readAt(0, 6))).toBe("abXYef");
    await file.truncate(3);
    expect(file.size).toBe(3);
    expect(new TextDecoder().decode(await file.readAt(0, 6))).toBe("abX");
    await file.close();
    await file.close();
    await scratch.create();
  } finally {
    await scratch.close();
  }
  expect(await readdir(directory)).toEqual(["keep.txt"]);
  expect(await readFile(retained, "utf8")).toBe("keep");
  await expect(scratch.create()).rejects.toMatchObject({
    code: "FILES_SCRATCH_CLOSED",
  });
});

test("retries a failed scratch file close and shares concurrent attempts", async () => {
  const directory = await fixture();
  filesystemFailure.closeFailures = 1;
  const scratch = await createScratchDirectory(directory);
  const file = await scratch.create();
  const first = file.close();
  const concurrent = file.close();
  expect(concurrent).toBe(first);
  await expect(first).rejects.toThrow("synthetic scratch close failure");
  await expect(concurrent).rejects.toThrow("synthetic scratch close failure");

  await file.close();
  await scratch.close();
  expect(filesystemFailure.closeCalls).toBe(2);
});

test("retries directory cleanup without repeating successfully closed files", async () => {
  const directory = await fixture();
  filesystemFailure.closeFailures = 1;
  const scratch = await createScratchDirectory(directory);
  await scratch.create();
  await scratch.create();
  const first = scratch.close();
  const concurrent = scratch.close();
  expect(concurrent).toBe(first);
  await expect(first).rejects.toMatchObject({
    code: "FILES_SCRATCH_CLEANUP_FAILED",
    details: { directory: expect.stringContaining("cc-scratch-") },
  });
  await expect(concurrent).rejects.toMatchObject({
    code: "FILES_SCRATCH_CLEANUP_FAILED",
  });
  expect(filesystemFailure.closeCalls).toBe(2);
  await expect(scratch.create()).rejects.toMatchObject({
    code: "FILES_SCRATCH_CLOSED",
  });

  await scratch.close();
  expect(filesystemFailure.closeCalls).toBe(3);
  expect(await readdir(directory)).toEqual([]);
});

test("preserves every scratch handle failure from one directory close", async () => {
  const directory = await fixture();
  filesystemFailure.closeFailures = 2;
  const scratch = await createScratchDirectory(directory);
  await scratch.create();
  await scratch.create();

  let failure: unknown;
  try {
    await scratch.close();
  } catch (cause) {
    failure = cause;
  }
  expect(failure).toMatchObject({ code: "FILES_SCRATCH_CLEANUP_FAILED" });
  const cause = (failure as Error & { cause?: unknown }).cause;
  expect(cause).toBeInstanceOf(AggregateError);
  expect((cause as AggregateError).errors).toHaveLength(2);

  await scratch.close();
  expect(filesystemFailure.closeCalls).toBe(4);
});

test("retries directory removal without reclosing files", async () => {
  const directory = await fixture();
  filesystemFailure.removeFailures = 1;
  const scratch = await createScratchDirectory(directory);
  await scratch.create();

  await expect(scratch.close()).rejects.toMatchObject({
    code: "FILES_SCRATCH_CLEANUP_FAILED",
    cause: expect.objectContaining({
      message: "synthetic scratch removal failure",
    }),
  });
  expect(filesystemFailure.closeCalls).toBe(1);
  expect(filesystemFailure.removeCalls).toBe(1);

  await scratch.close();
  expect(filesystemFailure.closeCalls).toBe(1);
  expect(filesystemFailure.removeCalls).toBe(2);
});

test("waits for an in-flight create before closing the directory", async () => {
  const directory = await fixture();
  const openGate = deferred();
  const openStarted = deferred();
  filesystemFailure.openGate = openGate.promise;
  filesystemFailure.openStarted = openStarted.resolve;
  const scratch = await createScratchDirectory(directory);

  const creating = scratch.create();
  await openStarted.promise;
  const closing = scratch.close();
  let closeSettled = false;
  void closing.then(() => {
    closeSettled = true;
  });
  await Promise.resolve();
  expect(closeSettled).toBe(false);
  openGate.resolve();

  await expect(creating).rejects.toMatchObject({
    code: "FILES_SCRATCH_CLOSED",
  });
  await closing;
  expect(filesystemFailure.closeCalls).toBe(1);
  expect(await readdir(directory)).toEqual([]);
});
