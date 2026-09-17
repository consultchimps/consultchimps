import { execFile } from "node:child_process";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { promisify } from "node:util";

import { expect } from "vitest";
import { DuckDBInstance } from "@duckdb/node-api";
import * as XLSX from "xlsx";
import Sqlite from "better-sqlite3";
import JSZip from "jszip";

export const execute: typeof execFile.__promisify__ = promisify(execFile);
export const cli: string = fileURLToPath(
  new URL("../dist/index.js", import.meta.url),
);
const directories: string[] = [];

export async function directory(): Promise<string> {
  const value = await mkdtemp(path.join(tmpdir(), "cc-db-cli-test-"));
  directories.push(value);
  return value;
}

export async function cleanupDirectories(): Promise<void> {
  await Promise.all(
    directories
      .splice(0)
      .map((value) => rm(value, { recursive: true, force: true })),
  );
}

export async function run(
  args: string[],
  cwd?: string,
): Promise<Record<string, unknown>> {
  const result = await execute(
    process.execPath,
    [cli, "--json", "db", ...args],
    { encoding: "utf8", cwd },
  );
  expect(result.stderr).toBe("");
  expect(result.stdout.trim().split("\n")).toHaveLength(1);
  const envelope: unknown = JSON.parse(result.stdout);
  expect(envelope).toMatchObject({ ok: true });
  if (
    envelope === null ||
    typeof envelope !== "object" ||
    !("result" in envelope) ||
    envelope.result === null ||
    typeof envelope.result !== "object"
  )
    throw new Error("Missing command result.");
  return envelope.result as Record<string, unknown>;
}

export async function runHuman(args: string[]): Promise<string> {
  const result = await execute(process.execPath, [cli, "db", ...args], {
    encoding: "utf8",
  });
  expect(result.stderr).toBe("");
  expect(result.stdout.trim()).not.toMatch(/^\{/u);
  return result.stdout;
}

export async function runFailure(args: string[]): Promise<{
  readonly stdout: string;
  readonly stderr: string;
}> {
  try {
    await execute(process.execPath, [cli, ...args], { encoding: "utf8" });
  } catch (error) {
    if (
      error instanceof Error &&
      "stdout" in error &&
      typeof error.stdout === "string" &&
      "stderr" in error &&
      typeof error.stderr === "string"
    ) {
      return { stdout: error.stdout, stderr: error.stderr };
    }
    throw error;
  }
  throw new Error("The command unexpectedly succeeded.");
}

export async function closeFailureLoader(root: string): Promise<string> {
  const shim = path.join(root, "database-close-failure.mjs");
  const loader = path.join(root, "database-close-failure-loader.mjs");
  const databaseModule = pathToFileURL(
    fileURLToPath(new URL("../../db/dist/node.js", import.meta.url)),
  ).href;
  const coreModule = pathToFileURL(
    fileURLToPath(new URL("../../core/dist/index.js", import.meta.url)),
  ).href;
  await writeFile(
    shim,
    `export * from ${JSON.stringify(databaseModule)};
import { createDatabase as createRealDatabase } from ${JSON.stringify(databaseModule)};
import { ConsultChimpsError } from ${JSON.stringify(coreModule)};
export async function createDatabase(options) {
  const created = await createRealDatabase(options);
  return {
    ...created,
    database: new Proxy(created.database, {
      get(target, property) {
        if (property === "close") return async () => {
          await target.close();
          throw new ConsultChimpsError(
            "DB_INJECTED_CLOSE_FAILURE",
            "The injected database close failed after the database was created.",
          );
        };
        const value = Reflect.get(target, property, target);
        return typeof value === "function" ? value.bind(target) : value;
      },
    }),
  };
}
`,
  );
  await writeFile(
    loader,
    `const shim = ${JSON.stringify(pathToFileURL(shim).href)};
export async function resolve(specifier, context, nextResolve) {
  if (specifier === "@consultchimps/db/node") {
    return { shortCircuit: true, url: shim };
  }
  return nextResolve(specifier, context);
}
`,
  );
  return pathToFileURL(loader).href;
}

export function workbook(
  rows: readonly (readonly (string | number | boolean)[])[],
  hidden = false,
): Uint8Array {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    book,
    XLSX.utils.aoa_to_sheet(rows.map((row) => [...row])),
    "Inventory",
  );
  if (hidden) {
    book.Workbook = { Sheets: [{ Hidden: 1, name: "Inventory" }] };
  }
  return XLSX.write(book, { type: "buffer", bookType: "xlsx" }) as Uint8Array;
}

export async function workbookWithInvalidNumericCell(): Promise<Uint8Array> {
  const archive = await JSZip.loadAsync(workbook([["Name"], ["North"]]));
  const worksheet = archive.file("xl/worksheets/sheet1.xml");
  if (!worksheet) throw new Error("The generated workbook has no worksheet.");
  const xml = await worksheet.async("string");
  const invalid = xml.replace(
    '<c r="A2" t="str"><v>North</v></c>',
    '<c r="A2"><v>not-a-number</v></c>',
  );
  if (invalid === xml)
    throw new Error("The generated workbook cell was not found.");
  archive.file("xl/worksheets/sheet1.xml", invalid);
  return archive.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

export async function addUnsupportedObject(
  file: string,
  format: "sqlite" | "duckdb",
): Promise<void> {
  if (format === "sqlite") {
    const database = new Sqlite(file);
    try {
      database.exec("CREATE VIEW unsupported_view AS SELECT 1 AS value");
    } finally {
      database.close();
    }
    return;
  }

  const instance = await DuckDBInstance.create(file);
  const connection = await instance.connect();
  try {
    await connection.run("CREATE MACRO unsupported_macro(value) AS value + 1");
  } finally {
    connection.closeSync();
    instance.closeSync();
  }
}
