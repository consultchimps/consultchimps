import {
  lstat,
  mkdir,
  mkdtemp,
  readFile,
  rm,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { ConsultChimpsError } from "@consultchimps/core";
import { afterEach, expect, test, vi } from "vitest";

import { failAfterNativeArtifactCleanup } from "../src/native-artifact-cleanup.js";
import { createDatabase } from "../src/node.js";

const directories: string[] = [];

afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

test("refuses a colliding DuckDB spill path without replacing either artifact", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-native-collision-"));
  directories.push(directory);
  const output = path.join(directory, "output.duckdb");
  const outputBytes = Buffer.from("prior destination bytes");
  await writeFile(output, outputBytes);
  const candidateId = "00000000-0000-4000-8000-000000000015";
  const spillPath = path.join(
    directory,
    `.output.duckdb.cc-create-${candidateId}.tmp`,
  );
  await mkdir(spillPath);
  vi.spyOn(globalThis.crypto, "randomUUID").mockReturnValueOnce(candidateId);

  await expect(
    createDatabase({
      path: output,
      format: "duckdb",
      overwrite: true,
    }),
  ).rejects.toMatchObject({
    code: "DB_OUTPUT_EXISTS",
    details: { path: spillPath },
  });
  expect(Buffer.from(await readFile(output)).equals(outputBytes)).toBe(true);
  expect((await lstat(spillPath)).isDirectory()).toBe(true);
});

const artifact = {
  operation: "prepare" as const,
  stage: "preparation" as const,
  kind: "prepared" as const,
  format: "sqlite" as const,
  temporaryPath: "/synthetic/.review.ccplan.cc-prepare-private",
  storagePaths: [
    "/synthetic/.review.ccplan.cc-prepare-private",
    "/synthetic/.review.ccplan.cc-prepare-private-wal",
    "/synthetic/.review.ccplan.cc-prepare-private-shm",
  ],
};

test("rethrows an undefined operation failure after successful cleanup", async () => {
  const close = vi.fn(async () => undefined);
  const remove = vi.fn(async () => undefined);

  await expect(
    failAfterNativeArtifactCleanup({
      ...artifact,
      cause: undefined,
      close: [close],
      remove,
    }),
  ).rejects.toBeUndefined();
  expect(close).toHaveBeenCalledOnce();
  expect(remove).toHaveBeenCalledOnce();
});

test.each([
  "DB_NATIVE_SQLITE_CLEANUP_REQUIRED",
  "DB_DUCKDB_EXPORT_CLEANUP_FAILED",
  "DB_DUCKDB_OPEN_CLEANUP_FAILED",
])(
  "does not remove storage retained by inaccessible handles from %s",
  async (code) => {
    const primary = new ConsultChimpsError(
      code,
      "The native engine could not release its private handle.",
    );
    const remove = vi.fn(async () => undefined);

    let failure: unknown;
    try {
      await failAfterNativeArtifactCleanup({
        ...artifact,
        cause: primary,
        remove,
      });
    } catch (error) {
      failure = error;
    }

    expect(remove).not.toHaveBeenCalled();
    expect(failure).toMatchObject({
      code: "DB_NATIVE_TEMPORARY_CLEANUP_REQUIRED",
      details: {
        primaryCode: code,
        closeFailed: true,
        removalFailed: false,
        removalSkipped: true,
      },
      cause: expect.objectContaining({ errors: [primary] }),
    });
    expect((failure as Error).message).toContain("Restart this process");
  },
);

test("skips removal after a synchronous close failure and still settles later closes", async () => {
  const laterClose = vi.fn(async () => undefined);
  const remove = vi.fn(async () => undefined);

  let failure: unknown;
  try {
    await failAfterNativeArtifactCleanup({
      ...artifact,
      cause: undefined,
      close: [
        () => {
          throw undefined;
        },
        laterClose,
      ],
      remove,
    });
  } catch (error) {
    failure = error;
  }

  expect(failure).toMatchObject({
    code: "DB_NATIVE_TEMPORARY_CLEANUP_REQUIRED",
    details: {
      operation: "prepare",
      stage: "preparation",
      kind: "prepared",
      format: "sqlite",
      temporaryPath: artifact.temporaryPath,
      storagePaths: artifact.storagePaths,
      closeFailed: true,
      removalFailed: false,
      removalSkipped: true,
    },
  });
  if (!(failure instanceof Error)) throw new Error("Expected cleanup failure");
  expect(failure.message).toContain(artifact.temporaryPath);
  expect(laterClose).toHaveBeenCalledOnce();
  expect(remove).not.toHaveBeenCalled();
  expect(failure.cause).toBeInstanceOf(AggregateError);
  expect((failure.cause as AggregateError).errors).toEqual([
    undefined,
    undefined,
  ]);
});

test("retains a published outcome from the primary failure", async () => {
  const primary = new ConsultChimpsError(
    "DB_NATIVE_PUBLISHED_OPEN_FAILED",
    "The published file could not be reopened.",
    { details: { published: true, output: "/synthetic/output.sqlite" } },
  );

  await expect(
    failAfterNativeArtifactCleanup({
      ...artifact,
      cause: primary,
      async remove() {
        throw new Error("Injected native file removal failure");
      },
    }),
  ).rejects.toMatchObject({
    code: "DB_NATIVE_TEMPORARY_CLEANUP_REQUIRED",
    details: {
      published: true,
      publishedPath: "/synthetic/output.sqlite",
      primaryCode: "DB_NATIVE_PUBLISHED_OPEN_FAILED",
    },
  });
});

test("leaves source and destination files unchanged when private removal fails", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-native-cleanup-"));
  directories.push(directory);
  const source = path.join(directory, "source.xlsx");
  const destination = path.join(directory, "review.ccplan");
  const temporary = path.join(directory, ".review.ccplan.cc-prepare-private");
  const sourceBytes = Buffer.from("synthetic source bytes");
  const destinationBytes = Buffer.from("prior destination bytes");
  const temporaryBytes = Buffer.from("private staged workbook values");
  await Promise.all([
    writeFile(source, sourceBytes),
    writeFile(destination, destinationBytes),
    writeFile(temporary, temporaryBytes),
  ]);
  const removalFailure = new Error("Injected private removal failure");
  const operationFailure = new Error("Injected preparation failure");

  let failure: unknown;
  try {
    await failAfterNativeArtifactCleanup({
      operation: "prepare",
      stage: "preparation",
      kind: "prepared",
      format: "sqlite",
      temporaryPath: temporary,
      storagePaths: [temporary, `${temporary}-wal`, `${temporary}-shm`],
      cause: operationFailure,
      async remove() {
        throw removalFailure;
      },
    });
  } catch (error) {
    failure = error;
  }
  expect(failure).toMatchObject({
    code: "DB_NATIVE_TEMPORARY_CLEANUP_REQUIRED",
    details: {
      temporaryPath: temporary,
      closeFailed: false,
      removalFailed: true,
    },
  });
  expect((failure as Error).cause).toBeInstanceOf(AggregateError);
  expect(((failure as Error).cause as AggregateError).errors).toEqual([
    operationFailure,
    removalFailure,
  ]);
  expect(Buffer.from(await readFile(source)).equals(sourceBytes)).toBe(true);
  expect(
    Buffer.from(await readFile(destination)).equals(destinationBytes),
  ).toBe(true);
  expect(Buffer.from(await readFile(temporary)).equals(temporaryBytes)).toBe(
    true,
  );
});
