import { expect, test } from "@playwright/test";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { build, type Options } from "tsup";

import { inspectDatabase } from "@consultchimps/db";
import { openDatabase as openNativeDatabase } from "@consultchimps/db/node";

let bundle = "";
let bundleDirectory = "";

const workerSource = path.join(
  import.meta.dirname,
  "../src/workers/workspace.worker.ts",
);
const databaseShim = path.join(
  import.meta.dirname,
  "fixtures/workspace-schema-checkpoint-db.ts",
);

test.beforeAll(async () => {
  bundleDirectory = await mkdtemp(
    path.join(tmpdir(), "consultchimps-schema-checkpoint-"),
  );
  const injectCheckpointFailure: NonNullable<
    Options["esbuildPlugins"]
  >[number] = {
    name: "inject-schema-checkpoint-failure",
    setup(buildContext) {
      buildContext.onResolve(
        { filter: /^@consultchimps\/db$/ },
        (arguments_) =>
          path.resolve(arguments_.importer) === path.resolve(workerSource)
            ? { path: databaseShim }
            : undefined,
      );
    },
  };
  await build({
    entry: { "workspace-schema-checkpoint.worker": workerSource },
    outDir: bundleDirectory,
    format: ["esm"],
    platform: "browser",
    target: "es2022",
    splitting: false,
    clean: true,
    silent: true,
    noExternal: [/.*/u],
    esbuildPlugins: [injectCheckpointFailure],
    define: { "process.env.NEXT_PUBLIC_BASE_PATH": '""' },
  });
  bundle = await readFile(
    path.join(bundleDirectory, "workspace-schema-checkpoint.worker.js"),
    "utf8",
  );
});

test.afterAll(async () => {
  if (bundleDirectory !== "") {
    await rm(bundleDirectory, { recursive: true, force: true });
  }
});

test("reports a committed schema when its checkpoint fails", async ({
  page,
}, testInfo) => {
  await page.route(
    "**/__tests__/workspace-schema-checkpoint.worker.js",
    (route) => route.fulfill({ body: bundle, contentType: "text/javascript" }),
  );
  await page.addInitScript(() => {
    const NativeWorker = window.Worker;
    class CheckpointFailureWorker extends EventTarget {
      readonly worker = new NativeWorker(
        "/__tests__/workspace-schema-checkpoint.worker.js",
        { type: "module" },
      );

      constructor() {
        super();
        this.worker.addEventListener("message", (event) => {
          this.dispatchEvent(new MessageEvent("message", { data: event.data }));
        });
        this.worker.addEventListener("error", () => {
          this.dispatchEvent(new Event("error"));
        });
        this.worker.addEventListener("messageerror", () => {
          this.dispatchEvent(new MessageEvent("messageerror"));
        });
      }

      postMessage(
        message: unknown,
        options?: StructuredSerializeOptions | Transferable[],
      ): void {
        if (Array.isArray(options)) {
          this.worker.postMessage(message, options);
        } else {
          this.worker.postMessage(message, options);
        }
      }

      terminate(): void {
        this.worker.terminate();
      }
    }
    Object.defineProperty(window, "Worker", {
      configurable: true,
      value: CheckpointFailureWorker,
    });
  });

  const schema = {
    version: 1,
    tables: [
      {
        name: "Inventory",
        recordId: { prefix: "INV", padding: 4 },
        columns: [{ name: "Region", type: "text" }],
      },
    ],
  };
  await page.goto("/tools/db");
  await page
    .getByTestId("workspace-new-name")
    .fill(`checkpoint-${crypto.randomUUID()}.sqlite`);
  await page.getByTestId("workspace-new").click();
  await page.getByTestId("workspace-schema-input").fill(JSON.stringify(schema));
  await page.getByTestId("workspace-schema-plan").click();
  await page.getByTestId("workspace-schema-apply").click();

  await expect(page.getByTestId("workspace-table")).toContainText("Inventory");
  await expect(page.getByTestId("workspace-schema-review")).toHaveCount(0);
  await expect(page.getByTestId("workspace-error")).toContainText(
    "The schema changes were applied, but the browser could not finish saving the database",
  );
  await expect(page.getByTestId("workspace-error")).toContainText(
    "DB_BROWSER_SCHEMA_PERSISTENCE_REQUIRED",
  );
  await expect(page.getByTestId("workspace-notice")).toHaveCount(0);

  const downloadStarted = page.waitForEvent("download");
  await page.getByTestId("workspace-export-same").click();
  const download = await downloadStarted;
  const exportedPath = testInfo.outputPath("schema-checkpoint.sqlite");
  await download.saveAs(exportedPath);
  const exported = await openNativeDatabase({
    path: exportedPath,
    readonly: true,
  });
  try {
    const inspection = await inspectDatabase({ database: exported });
    expect(inspection.tables).toMatchObject([
      {
        name: "Inventory",
        schema: {
          recordId: { prefix: "INV", padding: 4 },
          columns: [{ name: "Region", type: "text" }],
        },
      },
    ]);
  } finally {
    await exported.close();
  }

  await page.getByTestId("workspace-schema-plan").click();
  await expect(page.getByTestId("workspace-schema-review")).toContainText(
    "No additive schema changes are needed",
  );
});
