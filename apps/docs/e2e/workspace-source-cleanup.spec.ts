import { expect, test } from "@playwright/test";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { build, type Options } from "tsup";

import { createWorkbookUpload } from "./fixtures";

let bundle = "";
let bundleDirectory = "";

const workerSource = path.join(
  import.meta.dirname,
  "../src/workers/workspace.worker.ts",
);
const databaseShim = path.join(
  import.meta.dirname,
  "fixtures/workspace-source-cleanup-db.ts",
);

test.beforeAll(async () => {
  bundleDirectory = await mkdtemp(
    path.join(tmpdir(), "consultchimps-source-cleanup-"),
  );
  const injectSourceCloseFailure: NonNullable<
    Options["esbuildPlugins"]
  >[number] = {
    name: "inject-workbook-source-close-failure",
    setup(buildContext) {
      buildContext.onResolve(
        // esbuild's Go regexp parser rejects the JavaScript Unicode flag.
        { filter: /^@consultchimps\/db$/ },
        (arguments_) =>
          path.resolve(arguments_.importer) === path.resolve(workerSource)
            ? { path: databaseShim }
            : undefined,
      );
    },
  };
  await build({
    entry: { "workspace-source-cleanup.worker": workerSource },
    outDir: bundleDirectory,
    format: ["esm"],
    platform: "browser",
    target: "es2022",
    splitting: false,
    clean: true,
    silent: true,
    noExternal: [/.*/u],
    esbuildPlugins: [injectSourceCloseFailure],
    define: { "process.env.NEXT_PUBLIC_BASE_PATH": '""' },
  });
  bundle = await readFile(
    path.join(bundleDirectory, "workspace-source-cleanup.worker.js"),
    "utf8",
  );
});

test.afterAll(async () => {
  if (bundleDirectory !== "") {
    await rm(bundleDirectory, { recursive: true, force: true });
  }
});

test("retries retained workbook cleanup before preparing another import", async ({
  page,
}) => {
  const seed = await createWorkbookUpload("seed.xlsx", [
    { name: "Inventory", rows: [["Region"], ["North"]] },
  ]);
  const next = await createWorkbookUpload("next.xlsx", [
    { name: "Inventory", rows: [["Region"], ["South"]] },
  ]);
  await page.route("**/__tests__/workspace-source-cleanup.worker.js", (route) =>
    route.fulfill({ body: bundle, contentType: "text/javascript" }),
  );
  await page.goto("/tools/db");

  const result = await page.evaluate(
    async ({ seedBase64, nextBase64, runId }) => {
      const worker = new Worker(
        "/__tests__/workspace-source-cleanup.worker.js",
        { type: "module" },
      );
      let nextId = 1;
      const request = (command: Record<string, unknown>) =>
        new Promise<Record<string, unknown>>((resolve, reject) => {
          const id = nextId;
          nextId += 1;
          const timeout = window.setTimeout(
            () => reject(new Error("Source cleanup worker timed out")),
            55_000,
          );
          const receive = (event: MessageEvent<Record<string, unknown>>) => {
            if (event.data["id"] !== id || event.data["type"] === "progress")
              return;
            window.clearTimeout(timeout);
            worker.removeEventListener("message", receive);
            resolve(event.data);
          };
          worker.addEventListener("message", receive);
          worker.addEventListener(
            "error",
            (event) => {
              window.clearTimeout(timeout);
              reject(new Error(event.message));
            },
            { once: true },
          );
          worker.postMessage({ ...command, id });
        });
      const file = (name: string, base64: string) => {
        const binary = atob(base64);
        const bytes = Uint8Array.from(binary, (character) =>
          character.charCodeAt(0),
        );
        return new File([bytes], name, {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
      };
      const source = (name: string, base64: string, role: string) => ({
        id: crypto.randomUUID(),
        file: file(name, base64),
        role,
        revision: "1",
      });
      const databaseName = `${runId}.sqlite`;

      try {
        const created = await request({
          type: "create",
          format: "sqlite",
          name: databaseName,
          schema: {
            version: 1,
            tables: [
              {
                name: "Inventory",
                recordId: { prefix: "INV", padding: 4 },
                columns: [{ name: "Region", type: "text" }],
              },
            ],
          },
        });
        const seedPrepared = await request({
          type: "prepareImport",
          sources: [source("seed.xlsx", seedBase64, "seed")],
        });
        const seedPlan = Reflect.get(seedPrepared, "plan") as {
          readonly id: string;
          readonly state: string;
        };
        const applied = await request({
          type: "applyImport",
          planId: seedPlan.id,
          delivery: {
            requestId: crypto.randomUUID(),
            vendor: "Synthetic vendor",
            entity: "Synthetic entity",
            phase: "Seed",
            coverage: "full",
            effectiveDate: null,
            receivedDate: null,
            note: "",
          },
        });

        const failed = await request({
          type: "prepareImport",
          sources: [source("next.xlsx", nextBase64, "failed close")],
        });
        const listedAfterFailure = await request({ type: "listImports" });

        const root = await navigator.storage.getDirectory();
        const scratchAfterFailure: string[] = [];
        for await (const name of root.keys()) {
          if (name.startsWith(".consultchimps-scratch-"))
            scratchAfterFailure.push(name);
        }

        const retried = await request({
          type: "prepareImport",
          sources: [source("next.xlsx", nextBase64, "retry")],
        });
        const listedAfterRetry = await request({ type: "listImports" });
        const scratchAfterRetry: string[] = [];
        for await (const name of root.keys()) {
          if (name.startsWith(".consultchimps-scratch-"))
            scratchAfterRetry.push(name);
        }

        await request({ type: "close" });
        const reopened = await request({ type: "reopen", name: databaseName });
        await request({ type: "close" });
        return {
          created,
          seedPrepared,
          applied,
          failed,
          listedAfterFailure,
          retried,
          listedAfterRetry,
          reopened,
          scratchAfterFailure,
          scratchAfterRetry,
        };
      } finally {
        worker.terminate();
      }
    },
    {
      seedBase64: seed.buffer.toString("base64"),
      nextBase64: next.buffer.toString("base64"),
      runId: `source-cleanup-${crypto.randomUUID()}`,
    },
  );

  expect(result.created).toMatchObject({ type: "ready" });
  expect(result.seedPrepared).toMatchObject({
    type: "importPrepared",
    plan: { state: "ready" },
  });
  expect(result.applied).toMatchObject({
    type: "importApplied",
    result: {
      appendedRows: 1,
      summary: { tables: [{ name: "Inventory", rowCount: 1 }] },
    },
  });
  expect(result.failed).toMatchObject({
    type: "error",
    code: "DB_BROWSER_IMPORT_CLEANUP_REQUIRED",
  });
  expect(result.listedAfterFailure).toMatchObject({
    type: "importsListed",
    plans: [{ id: expect.any(String), application: "applied" }],
    ignoredPlans: [],
  });
  expect(result.scratchAfterFailure.length).toBeGreaterThan(0);
  expect(result.retried).toMatchObject({
    type: "importPrepared",
    plan: { application: "pending" },
  });
  expect(result.listedAfterRetry).toMatchObject({
    type: "importsListed",
    plans: [
      { id: expect.any(String), application: "applied" },
      { id: expect.any(String), application: "pending" },
    ],
    ignoredPlans: [],
  });
  expect(result.scratchAfterRetry).toEqual([]);
  expect(result.reopened).toMatchObject({
    type: "ready",
    summary: { tables: [{ name: "Inventory", rowCount: 1 }] },
  });
});
