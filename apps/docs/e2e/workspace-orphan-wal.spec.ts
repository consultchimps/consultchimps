import { expect, test } from "@playwright/test";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { build } from "tsup";

let bundle = "";
let bundleDirectory = "";
const workerSource = path.join(
  import.meta.dirname,
  "fixtures/workspace-orphan-wal.worker.ts",
);

test.beforeAll(async () => {
  bundleDirectory = await mkdtemp(
    path.join(tmpdir(), "consultchimps-orphan-wal-"),
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
    path.join(bundleDirectory, "workspace-orphan-wal.worker.js"),
    "utf8",
  );
});

test.afterAll(async () => {
  if (bundleDirectory !== "")
    await rm(bundleDirectory, { recursive: true, force: true });
});

test("preserves orphaned DuckDB WAL files across browser create and import", async ({
  page,
}) => {
  await page.route("**/__tests__/workspace-orphan-wal.worker.js", (route) =>
    route.fulfill({ body: bundle, contentType: "text/javascript" }),
  );
  await page.goto("/tools/db");
  const result = await page.evaluate(async () => {
    const worker = new Worker("/__tests__/workspace-orphan-wal.worker.js", {
      type: "module",
    });
    try {
      return await new Promise<unknown>((resolve, reject) => {
        const timeout = window.setTimeout(
          () => reject(new Error("Orphan WAL worker timed out")),
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
          if (
            typeof message !== "object" ||
            message === null ||
            Reflect.get(message, "ok") !== true
          ) {
            window.clearTimeout(timeout);
            reject(new Error(JSON.stringify(message)));
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

  expect(result).toMatchObject({
    results: Array.from({ length: 8 }, () => ({
      failure: {
        code: "DB_BROWSER_INCOMPLETE_STORAGE",
        details: {
          format: "duckdb",
          missing: [expect.stringMatching(/\.duckdb$/u)],
          retained: [expect.stringMatching(/\.duckdb\.wal$/u)],
        },
      },
      mainExists: false,
      walPreserved: true,
    })),
    mainOnly: { mainExists: true, reopened: true },
  });
});
