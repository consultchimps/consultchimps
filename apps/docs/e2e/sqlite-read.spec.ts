import { expect, test } from "@playwright/test";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { build } from "tsup";

let directory = "";
let bundle = "";

test.beforeAll(async () => {
  directory = await mkdtemp(path.join(tmpdir(), "consultchimps-sqlite-read-"));
  // tsup treats entries as glob patterns, so Windows separators must not become
  // escape characters in the pattern.
  const entry = path
    .join(import.meta.dirname, "fixtures/sqlite-read.worker.ts")
    .split(path.sep)
    .join("/");
  await build({
    entry: [entry],
    outDir: directory,
    format: ["esm"],
    platform: "browser",
    target: "es2022",
    splitting: false,
    clean: true,
    silent: true,
    noExternal: [/.*/u],
  });
  bundle = await readFile(
    path.join(directory, "sqlite-read.worker.js"),
    "utf8",
  );
});

test.afterAll(async () => {
  if (directory) await rm(directory, { recursive: true, force: true });
});

test("reads ephemeral SQLite catalogs without storage or isolation headers", async ({
  context,
  page,
}) => {
  const wasm = await readFile(
    path.join(import.meta.dirname, "../public/database-wasm/sqlite3.wasm"),
  );
  // Requests issued from inside a worker are intercepted at the context level.
  await context.route("**/__tests__/sqlite-read.worker.js", (route) =>
    route.fulfill({ body: bundle, contentType: "text/javascript" }),
  );
  // A non-root deployment mapping must be used for every locator invocation.
  await context.route("**/nested/catalog-assets/sqlite3.wasm", (route) =>
    route.fulfill({ body: wasm, contentType: "application/wasm" }),
  );
  await page.goto("/docs/libraries");
  const result = await page.evaluate(async () => {
    const worker = new Worker("/__tests__/sqlite-read.worker.js", {
      type: "module",
    });
    try {
      return await new Promise<unknown>((resolve, reject) => {
        const timeout = setTimeout(
          () => reject(new Error("SQLite reader worker timed out")),
          55_000,
        );
        worker.onerror = () => {
          clearTimeout(timeout);
          reject(new Error("SQLite reader worker failed"));
        };
        worker.onmessage = (
          event: MessageEvent<{
            ok: boolean;
            result?: unknown;
            error?: string;
          }>,
        ) => {
          clearTimeout(timeout);
          if (event.data.ok) resolve(event.data.result);
          else reject(new Error(event.data.error));
        };
        worker.postMessage({
          wasmUrl: new URL(
            "/nested/catalog-assets/sqlite3.wasm",
            location.origin,
          ).href,
        });
      });
    } finally {
      worker.terminate();
    }
  });
  expect(result).toEqual({
    id: "9223372036854775807",
    readonly: "DB_SQLITE_READ_ONLY",
    rowLimit: "DB_SQLITE_READ_LIMIT_EXCEEDED",
    steps: "DB_SQLITE_READ_LIMIT_EXCEEDED",
    badRuntime: "DB_SQLITE_READ_RUNTIME_UNAVAILABLE",
    calls: ["sqlite3.wasm"],
    before: [],
    after: [],
    isolated: false,
    shared: "undefined",
    memoryAfterClose: 0,
  });
});
