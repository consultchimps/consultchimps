import { expect, test, vi } from "vitest";

import { useBrowserConversionArtifact } from "../src/browser-conversion-cleanup.js";

function artifact() {
  return {
    name: ".consultchimps-conversion-test.duckdb",
    format: "duckdb" as const,
    directory: "consultchimps/databases",
    storageNames: [
      ".consultchimps-conversion-test.duckdb",
      ".consultchimps-conversion-test.duckdb.wal",
    ],
    close: vi.fn(async () => undefined),
    remove: vi.fn(async () => undefined),
  };
}

test("removes a temporary conversion working copy after export", async () => {
  const temporary = artifact();

  await expect(
    useBrowserConversionArtifact(temporary, async () => 4096),
  ).resolves.toEqual({ value: 4096, warnings: [] });
  expect(temporary.close).toHaveBeenCalledOnce();
  expect(temporary.remove).toHaveBeenCalledOnce();
});

test("reports retained temporary data after a successful export", async () => {
  const temporary = artifact();
  const cleanupFailure = new Error("Injected browser removal failure");
  temporary.remove.mockRejectedValue(cleanupFailure);

  const result = await useBrowserConversionArtifact(
    temporary,
    async () => 4096,
  );
  expect(result).toMatchObject({ value: 4096 });
  expect(result.warnings).toEqual([
    expect.stringContaining(".consultchimps-conversion-test.duckdb.wal"),
  ]);
});

test("retains a temporary working copy when its handle cannot close", async () => {
  const temporary = artifact();
  const cleanupFailure = new Error("Injected browser close failure");
  temporary.close.mockRejectedValue(cleanupFailure);

  const result = await useBrowserConversionArtifact(
    temporary,
    async () => 4096,
  );
  expect(result).toMatchObject({ value: 4096 });
  expect(result.warnings).toHaveLength(1);
  expect(temporary.remove).not.toHaveBeenCalled();
});

test("describes SQLite cleanup using its logical pool name", async () => {
  const temporary = {
    ...artifact(),
    name: ".consultchimps-conversion-test.sqlite",
    format: "sqlite" as const,
    directory: "consultchimps/sqlite",
    storageNames: undefined,
  };
  temporary.remove.mockRejectedValue(
    new Error("Injected browser removal failure"),
  );

  const result = await useBrowserConversionArtifact(
    temporary,
    async () => 4096,
  );
  expect(result.warnings).toEqual([
    expect.stringContaining(
      'logical working copy ".consultchimps-conversion-test.sqlite" may remain in the SQLite SAH pool at "consultchimps/sqlite"',
    ),
  ]);
  expect(result.warnings[0]).not.toContain(".sqlite.wal");
});

test("preserves the export cause when conversion cleanup also fails", async () => {
  const temporary = artifact();
  const exportFailure = new Error("Injected export failure");
  temporary.remove.mockRejectedValue(
    new Error("Injected browser removal failure"),
  );

  let rejection: unknown;
  try {
    await useBrowserConversionArtifact(temporary, async () => {
      throw exportFailure;
    });
  } catch (error) {
    rejection = error;
  }
  expect(rejection).toMatchObject({
    code: "DB_BROWSER_CONVERSION_CLEANUP_REQUIRED",
    details: {
      directory: "consultchimps/databases",
      conversionName: ".consultchimps-conversion-test.duckdb",
      format: "duckdb",
    },
  });
  expect((rejection as Error).cause).toBeInstanceOf(AggregateError);
  expect(((rejection as Error).cause as AggregateError).errors).toContain(
    exportFailure,
  );
});

test("rethrows the original export failure after successful cleanup", async () => {
  const temporary = artifact();
  const exportFailure = new Error("Injected export failure");

  await expect(
    useBrowserConversionArtifact(temporary, async () => {
      throw exportFailure;
    }),
  ).rejects.toBe(exportFailure);
  expect(temporary.close).toHaveBeenCalledOnce();
  expect(temporary.remove).toHaveBeenCalledOnce();
});
