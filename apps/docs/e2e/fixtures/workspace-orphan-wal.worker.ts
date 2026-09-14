import type { RandomAccessFile } from "@consultchimps/core";
import {
  configureBrowserDatabaseRuntime,
  type BrowserDatabaseRuntime,
} from "@consultchimps/db/browser";
import type { DatabaseFormat } from "@consultchimps/db/schema";

interface RunRequest {
  readonly id: number;
  readonly type: "run";
  readonly origin: string;
  readonly runId: string;
}

class MemoryFile implements RandomAccessFile {
  #bytes = new Uint8Array();

  constructor(readonly name: string) {}

  get size(): number {
    return this.#bytes.byteLength;
  }

  async readAt(offset: number, length: number): Promise<Uint8Array> {
    return this.#bytes.slice(offset, offset + length);
  }

  async writeAt(offset: number, bytes: Uint8Array): Promise<void> {
    const end = offset + bytes.byteLength;
    if (end > this.#bytes.byteLength) {
      const expanded = new Uint8Array(end);
      expanded.set(this.#bytes);
      this.#bytes = expanded;
    }
    this.#bytes.set(bytes, offset);
  }

  async truncate(size: number): Promise<void> {
    const resized = new Uint8Array(size);
    resized.set(this.#bytes.subarray(0, size));
    this.#bytes = resized;
  }

  async close(): Promise<void> {}

  bytes(): Uint8Array<ArrayBuffer> {
    return this.#bytes.slice();
  }
}

async function bytesOf(handle: FileSystemFileHandle): Promise<Uint8Array> {
  return new Uint8Array(await (await handle.getFile()).arrayBuffer());
}

async function exists(
  directory: FileSystemDirectoryHandle,
  name: string,
): Promise<boolean> {
  try {
    await directory.getFileHandle(name);
    return true;
  } catch (error) {
    if (error instanceof DOMException && error.name === "NotFoundError")
      return false;
    throw error;
  }
}

async function seedWal(
  directory: FileSystemDirectoryHandle,
  name: string,
  bytes: Uint8Array<ArrayBuffer>,
): Promise<void> {
  const writable = await (
    await directory.getFileHandle(`${name}.wal`, { create: true })
  ).createWritable();
  await writable.write(bytes);
  await writable.truncate(bytes.byteLength);
  await writable.close();
}

function failureOf(error: unknown): unknown {
  return {
    code:
      typeof error === "object" && error !== null
        ? Reflect.get(error, "code")
        : undefined,
    message: error instanceof Error ? error.message : String(error),
    details:
      typeof error === "object" && error !== null
        ? Reflect.get(error, "details")
        : undefined,
  };
}

async function databaseSource(
  runtime: BrowserDatabaseRuntime,
  runId: string,
  format: DatabaseFormat,
): Promise<MemoryFile> {
  const name = `${runId}-source.${format}`;
  const created = await runtime.createDatabase({ name, format });
  const output = new MemoryFile(name);
  try {
    await runtime.exportDatabase({
      database: created.database,
      name,
      destination: output,
      format,
    });
  } finally {
    await created.database.close();
  }
  return output;
}

async function run(request: RunRequest) {
  const directoryName = `consultchimps-orphan-wal-${request.runId}`;
  const runtime = await configureBrowserDatabaseRuntime({
    sqlite: {
      wasmUrl: `${request.origin}/database-wasm/sqlite3.wasm`,
      directory: `/${directoryName}-sqlite`,
      initialCapacity: 16,
    },
    duckdb: {
      wasmUrl: `${request.origin}/database-wasm/duckdb-eh.wasm`,
      workerUrl: `${request.origin}/database-wasm/duckdb-browser-eh.worker.js`,
    },
    opfsDirectory: directoryName,
  });
  const root = await navigator.storage.getDirectory();
  const directory = await root.getDirectoryHandle(directoryName);
  const sources = {
    sqlite: await databaseSource(runtime, request.runId, "sqlite"),
    duckdb: await databaseSource(runtime, request.runId, "duckdb"),
  };
  const walBytes = new Uint8Array([7, 19, 31, 43, 59, 71]);
  const cases: Array<{
    readonly operation: "create" | "import";
    readonly format: DatabaseFormat;
    readonly overwrite: boolean;
  }> = [];
  for (const overwrite of [false, true]) {
    for (const format of ["sqlite", "duckdb"] as const) {
      cases.push({ operation: "create", format, overwrite });
      cases.push({ operation: "import", format, overwrite });
    }
  }
  const results = [];
  for (const [index, testCase] of cases.entries()) {
    const name = `${request.runId}-orphan-${String(index)}.duckdb`;
    await seedWal(directory, name, walBytes);
    let failure: unknown;
    try {
      const result =
        testCase.operation === "create"
          ? await runtime.createDatabase({
              name,
              format: testCase.format,
              overwrite: testCase.overwrite,
            })
          : await runtime.importDatabase({
              name,
              source: sources[testCase.format],
              overwrite: testCase.overwrite,
            });
      await ("database" in result ? result.database : result).close();
    } catch (error) {
      failure = failureOf(error);
    }
    const retained = await bytesOf(
      await directory.getFileHandle(`${name}.wal`),
    );
    results.push({
      ...testCase,
      failure,
      mainExists: await exists(directory, name),
      walPreserved:
        retained.length === walBytes.length &&
        retained.every((byte, byteIndex) => byte === walBytes[byteIndex]),
    });
  }

  const mainOnlyName = `${request.runId}-main-only.duckdb`;
  const mainOnlyBytes = sources.duckdb.bytes();
  const mainOnlyWritable = await (
    await directory.getFileHandle(mainOnlyName, { create: true })
  ).createWritable();
  await mainOnlyWritable.write(mainOnlyBytes);
  await mainOnlyWritable.truncate(mainOnlyBytes.byteLength);
  await mainOnlyWritable.close();
  const reopened = await runtime.openDatabase({ name: mainOnlyName });
  await reopened.close();

  return {
    results,
    mainOnly: {
      mainExists: await exists(directory, mainOnlyName),
      reopened: true,
    },
  };
}

const scope = globalThis as unknown as {
  onmessage: ((event: MessageEvent<RunRequest>) => void) | null;
  postMessage(message: unknown): void;
};
scope.onmessage = (event) => {
  if (event.data.type !== "run") return;
  void run(event.data).then(
    (result) => scope.postMessage({ id: event.data.id, ok: true, result }),
    (error: unknown) =>
      scope.postMessage({
        id: event.data.id,
        ok: false,
        error: failureOf(error),
      }),
  );
};
