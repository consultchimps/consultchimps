import { expect, test } from "@playwright/test";
import { execFile } from "node:child_process";
import { copyFile, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { build } from "tsup";

import { inspectDatabase, planConversion } from "@consultchimps/db";
import { createDatabase, openDatabase } from "@consultchimps/db/node";

let bundle = "";
let bundleDirectory = "";

const workerSource = path.join(
  import.meta.dirname,
  "fixtures/workspace-duckdb-wal-export.worker.ts",
);
const execFileAsync = promisify(execFile);

interface WalExportResult {
  readonly walSize: number;
  readonly mainOnly: { readonly rows: string; readonly views: string };
  readonly writablePair: { readonly rows: string; readonly views: string };
  readonly readonlyRows: string;
  readonly failedExportMessage: string;
  readonly failedDestinationSize: number;
  readonly sourceMainPreservedAfterFailure: boolean;
  readonly sourceWalPreservedAfterFailure: boolean;
  readonly snapshotFilesAfterFailure: readonly string[];
  readonly sourceMainPreserved: boolean;
  readonly sourceWalPreserved: boolean;
  readonly snapshotFiles: readonly string[];
  readonly exportedSize: number;
  readonly exportedBase64: string;
}

async function createNativeWalFixture(databasePath: string): Promise<{
  readonly main: Buffer;
  readonly wal: Buffer;
}> {
  const created = await createDatabase({
    path: databasePath,
    format: "duckdb",
    schema: {
      version: 1,
      tables: [
        {
          name: "Inventory",
          recordId: { prefix: "INV", padding: 6 },
          columns: [{ name: "value", type: "text" }],
        },
      ],
    },
  });
  await created.database.close();
  const baselineMain = await readFile(databasePath);
  const script = `
    import { DuckDBInstance } from "@duckdb/node-api";
    const instance = await DuckDBInstance.create(process.argv[1]);
    const connection = await instance.connect();
    await connection.run("SET checkpoint_threshold = '1 TB'");
    await connection.run("BEGIN TRANSACTION");
    await connection.run("INSERT INTO Inventory (record_id, value) VALUES ('INV-000001', 'WAL-backed')");
    await connection.run("UPDATE _consultchimps_tables SET next_record_id = 2 WHERE table_name = 'Inventory'");
    await connection.run("CREATE VIEW wal_inventory AS SELECT record_id, value FROM Inventory");
    await connection.run("COMMIT");
    const reader = await connection.runAndReadAll("SELECT count(*) AS value FROM Inventory");
    process.stdout.write(JSON.stringify(reader.getRowObjectsJson()), () => process.exit(0));
  `;
  const { stdout } = await execFileAsync(
    process.execPath,
    ["--input-type=module", "--eval", script, databasePath],
    {
      cwd: path.resolve(import.meta.dirname, "../../../packages/db"),
      encoding: "utf8",
    },
  );
  expect(JSON.parse(stdout)).toEqual([{ value: "1" }]);
  const main = await readFile(databasePath);
  const wal = await readFile(`${databasePath}.wal`);
  expect(main.equals(baselineMain)).toBe(true);
  expect(wal.byteLength).toBeGreaterThan(0);
  return { main, wal };
}

async function queryNativeView(databasePath: string): Promise<unknown> {
  const script = `
    import { DuckDBInstance } from "@duckdb/node-api";
    const instance = await DuckDBInstance.create(process.argv[1], { access_mode: "READ_ONLY" });
    const connection = await instance.connect();
    try {
      const reader = await connection.runAndReadAll(
        "SELECT record_id, value FROM wal_inventory ORDER BY record_id",
      );
      process.stdout.write(JSON.stringify(reader.getRowObjectsJson()));
    } finally {
      connection.closeSync();
      instance.closeSync();
    }
  `;
  const { stdout } = await execFileAsync(
    process.execPath,
    ["--input-type=module", "--eval", script, databasePath],
    {
      cwd: path.resolve(import.meta.dirname, "../../../packages/db"),
      encoding: "utf8",
    },
  );
  return JSON.parse(stdout) as unknown;
}

test.beforeAll(async () => {
  bundleDirectory = await mkdtemp(
    path.join(tmpdir(), "consultchimps-duckdb-wal-export-"),
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
    path.join(bundleDirectory, "workspace-duckdb-wal-export.worker.js"),
    "utf8",
  );
});

test.afterAll(async () => {
  if (bundleDirectory !== "") {
    await rm(bundleDirectory, { recursive: true, force: true });
  }
});

test("exports committed DuckDB WAL contents from a read-only workspace", async ({
  page,
}) => {
  const nativeSourcePath = path.join(bundleDirectory, "native-source.duckdb");
  const nativeSource = await createNativeWalFixture(nativeSourcePath);
  const recoveryPath = path.join(bundleDirectory, "native-recovery.duckdb");
  await Promise.all([
    copyFile(nativeSourcePath, recoveryPath),
    copyFile(`${nativeSourcePath}.wal`, `${recoveryPath}.wal`),
  ]);
  const recovered = await openDatabase({ path: recoveryPath, readonly: true });
  try {
    expect(
      (await inspectDatabase({ database: recovered })).tables,
    ).toMatchObject([{ name: "Inventory", rowCount: 1n }]);
  } finally {
    await recovered.close();
  }
  expect(await queryNativeView(recoveryPath)).toEqual([
    { record_id: "INV-000001", value: "WAL-backed" },
  ]);

  await page.route(
    "**/__tests__/workspace-duckdb-wal-export.worker.js",
    (route) => route.fulfill({ body: bundle, contentType: "text/javascript" }),
  );
  await page.goto("/workspace");
  const result = await page.evaluate(
    async ({ sourceMainBase64, sourceWalBase64 }) => {
      const worker = new Worker(
        "/__tests__/workspace-duckdb-wal-export.worker.js",
        { type: "module" },
      );
      try {
        return await new Promise<WalExportResult>((resolve, reject) => {
          const timeout = window.setTimeout(
            () => reject(new Error("DuckDB WAL export worker timed out")),
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
          worker.addEventListener(
            "message",
            (event: MessageEvent<unknown>) => {
              if (
                typeof event.data !== "object" ||
                event.data === null ||
                !("ok" in event.data)
              ) {
                window.clearTimeout(timeout);
                reject(
                  new Error(
                    "DuckDB WAL export worker returned an invalid response",
                  ),
                );
                return;
              }
              const message = event.data as {
                readonly ok: unknown;
                readonly error?: unknown;
                readonly result?: unknown;
              };
              if (message["ok"] !== true) {
                window.clearTimeout(timeout);
                reject(new Error(JSON.stringify(message["error"] ?? message)));
                return;
              }
              if (
                typeof message.result !== "object" ||
                message.result === null
              ) {
                window.clearTimeout(timeout);
                reject(
                  new Error(
                    "DuckDB WAL export worker returned an invalid result",
                  ),
                );
                return;
              }
              window.clearTimeout(timeout);
              resolve(message.result as WalExportResult);
            },
            { once: true },
          );
          worker.postMessage({
            id: 1,
            type: "run",
            origin: window.location.origin,
            runId: crypto.randomUUID(),
            sourceMainBase64,
            sourceWalBase64,
          });
        });
      } finally {
        worker.terminate();
      }
    },
    {
      sourceMainBase64: nativeSource.main.toString("base64"),
      sourceWalBase64: nativeSource.wal.toString("base64"),
    },
  );

  expect(result).toMatchObject({
    walSize: expect.any(Number),
    mainOnly: { rows: "0", views: "0" },
    writablePair: { rows: "1", views: "1" },
    readonlyRows: "1",
    failedExportMessage: "Injected WAL export write failure",
    failedDestinationSize: 0,
    sourceMainPreservedAfterFailure: true,
    sourceWalPreservedAfterFailure: true,
    snapshotFilesAfterFailure: [],
    sourceMainPreserved: true,
    sourceWalPreserved: true,
    snapshotFiles: [],
    exportedSize: expect.any(Number),
    exportedBase64: expect.any(String),
  });
  expect(result.walSize).toEqual(expect.any(Number));
  expect(result.walSize).toBeGreaterThan(0);
  expect(result.exportedSize).toBeGreaterThan(100);

  const exportPath = path.join(bundleDirectory, "wal-export.duckdb");
  await writeFile(exportPath, Buffer.from(result.exportedBase64, "base64"));
  const database = await openDatabase({ path: exportPath, readonly: true });
  try {
    const inspection = await inspectDatabase({ database });
    expect(inspection.tables).toMatchObject([
      { name: "Inventory", rowCount: 1n },
    ]);
    const conversion = await planConversion({ database, format: "sqlite" });
    expect(conversion.issues).toContainEqual(
      expect.objectContaining({
        kind: "unsupported-object",
        name: "wal_inventory",
        objectType: "VIEW",
      }),
    );
  } finally {
    await database.close();
  }
  expect(await queryNativeView(exportPath)).toEqual([
    { record_id: "INV-000001", value: "WAL-backed" },
  ]);
});
