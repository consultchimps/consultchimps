import { expect, test } from "@playwright/test";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { build } from "tsup";

let bundle = "";
let bundleDirectory = "";

const workerSource = path.join(
  import.meta.dirname,
  "fixtures/workspace-export-recovery.worker.ts",
);

test.beforeAll(async () => {
  bundleDirectory = await mkdtemp(
    path.join(tmpdir(), "consultchimps-export-recovery-"),
  );
  await build({
    entry: [workerSource],
    outDir: bundleDirectory,
    format: ["esm"],
    platform: "browser",
    target: "es2022",
    splitting: false,
    clean: true,
    silent: true,
    noExternal: [/.*/u],
  });
  bundle = await readFile(
    path.join(bundleDirectory, "workspace-export-recovery.worker.js"),
    "utf8",
  );
});

test.afterAll(async () => {
  if (bundleDirectory !== "") {
    await rm(bundleDirectory, { recursive: true, force: true });
  }
});

test("restores caller export files across real browser database engines", async ({
  page,
}) => {
  await page.route(
    "**/__tests__/workspace-export-recovery.worker.js",
    (route) =>
      route.fulfill({
        body: bundle,
        contentType: "text/javascript",
      }),
  );
  await page.goto("/workspace");
  const result = await page.evaluate(async () => {
    const worker = new Worker(
      "/__tests__/workspace-export-recovery.worker.js",
      { type: "module" },
    );
    try {
      return await new Promise<unknown>((resolve, reject) => {
        const timeout = window.setTimeout(
          () => reject(new Error("Export recovery worker timed out")),
          55_000,
        );
        worker.addEventListener(
          "error",
          (event) => {
            window.clearTimeout(timeout);
            reject(new Error(event.message));
          },
          { once: true },
        );
        worker.addEventListener("message", (event: MessageEvent<unknown>) => {
          const message = event.data;
          if (typeof message !== "object" || message === null) {
            window.clearTimeout(timeout);
            reject(new Error("Export recovery worker returned invalid data"));
            return;
          }
          if (Reflect.get(message, "ok") !== true) {
            window.clearTimeout(timeout);
            reject(
              new Error(
                JSON.stringify(Reflect.get(message, "error") ?? message),
              ),
            );
            return;
          }
          window.clearTimeout(timeout);
          resolve(Reflect.get(message, "result"));
        });
        worker.postMessage({
          id: 1,
          type: "run",
          origin: window.location.origin,
          runId: crypto.randomUUID(),
        });
      });
    } finally {
      worker.terminate();
    }
  });

  const preservedFailure = (source: string, target: string, mode: string) => ({
    source,
    target,
    failureMode: mode,
    failure:
      mode === "cancel"
        ? { code: "OPERATION_ABORTED" }
        : { message: "Injected destination write failure" },
    destinationPreserved: true,
    destinationSize: 1024 * 1024 + 17,
    originalSize: 1024 * 1024 + 17,
    closeCalls: 0,
    maximumRead: 1024 * 1024,
    sourceFormat: source,
    sourceTables: ["SyntheticRecords"],
    sourceRowCount: "1",
  });
  const successfulConversion = (source: string, target: string) => ({
    source,
    target,
    failureMode: null,
    output: {
      size: expect.any(Number),
      inspectedFormat: target,
      tables: ["SyntheticRecords"],
      rowCount: "1",
    },
    destinationPreserved: false,
    closeCalls: 0,
    sourceFormat: source,
    sourceTables: ["SyntheticRecords"],
    sourceRowCount: "1",
  });

  expect(result).toMatchObject({
    recoveries: [
      preservedFailure("sqlite", "sqlite", "write"),
      preservedFailure("sqlite", "sqlite", "cancel"),
      preservedFailure("duckdb", "duckdb", "write"),
      preservedFailure("duckdb", "duckdb", "cancel"),
      successfulConversion("sqlite", "duckdb"),
      preservedFailure("sqlite", "duckdb", "write"),
      successfulConversion("duckdb", "sqlite"),
      preservedFailure("duckdb", "sqlite", "cancel"),
    ],
    concurrency: {
      mutationCompletedWhilePaused: false,
      snapshotColumns: ["value"],
      snapshotRowCount: "1",
      currentColumns: ["value", "reviewed"],
      currentRowCount: "1",
      destinationCloseCalls: 0,
    },
  });
});
