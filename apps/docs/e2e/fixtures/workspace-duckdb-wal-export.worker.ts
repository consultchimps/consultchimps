import {
  AsyncDuckDB,
  DuckDBAccessMode,
  DuckDBDataProtocol,
  VoidLogger,
  type AsyncDuckDBConnection,
} from "@duckdb/duckdb-wasm";

import type { RandomAccessFile } from "@consultchimps/core";
import { inspectDatabase } from "@consultchimps/db";
import { configureBrowserDatabaseRuntime } from "@consultchimps/db/browser";

interface RunRequest {
  readonly id: number;
  readonly type: "run";
  readonly runId: string;
  readonly sourceMainBase64: string;
  readonly sourceWalBase64: string;
}

interface BrowserDirectoryHandle {
  readonly kind: "directory";
  readonly name: string;
  values(): AsyncIterableIterator<
    BrowserDirectoryHandle | (FileSystemFileHandle & { readonly kind: "file" })
  >;
  getFileHandle(
    name: string,
    options?: { readonly create?: boolean },
  ): Promise<FileSystemFileHandle>;
  removeEntry(name: string): Promise<void>;
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

  bytes(): Uint8Array {
    return this.#bytes.slice();
  }
}

class FailingFile extends MemoryFile {
  #failed = false;

  override async writeAt(offset: number, bytes: Uint8Array): Promise<void> {
    await super.writeAt(offset, bytes);
    if (!this.#failed) {
      this.#failed = true;
      throw new Error("Injected WAL export write failure");
    }
  }
}

interface RawDuckDb {
  readonly database: AsyncDuckDB;
  readonly connection: AsyncDuckDBConnection;
}

function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  return (
    left.byteLength === right.byteLength &&
    left.every((byte, index) => byte === right[index])
  );
}

function bytesToBase64(bytes: Uint8Array): string {
  const chunkSize = 32_768;
  let binary = "";
  for (let offset = 0; offset < bytes.byteLength; offset += chunkSize) {
    binary += String.fromCharCode(
      ...bytes.subarray(offset, offset + chunkSize),
    );
  }
  return btoa(binary);
}

function base64ToBytes(value: string): Uint8Array {
  const binary = atob(value);
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}

function queryScalar(
  table: Awaited<ReturnType<AsyncDuckDBConnection["query"]>>,
): unknown {
  const row = table.toArray()[0];
  if (row === undefined) return undefined;
  const value =
    typeof row === "object" && row !== null && "toJSON" in row
      ? (row as { toJSON(): unknown }).toJSON()
      : row;
  if (typeof value !== "object" || value === null) return undefined;
  return Reflect.get(value, "value");
}

async function directory(name: string): Promise<BrowserDirectoryHandle> {
  const root = await navigator.storage.getDirectory();
  return root.getDirectoryHandle(name, {
    create: true,
  }) as Promise<BrowserDirectoryHandle>;
}

async function fileBytes(
  storage: BrowserDirectoryHandle,
  name: string,
): Promise<Uint8Array> {
  const file = await (await storage.getFileHandle(name)).getFile();
  return new Uint8Array(await file.arrayBuffer());
}

async function replaceFile(
  storage: BrowserDirectoryHandle,
  name: string,
  bytes: Uint8Array,
): Promise<FileSystemFileHandle> {
  const handle = await storage.getFileHandle(name, { create: true });
  const writable = await handle.createWritable({ keepExistingData: false });
  try {
    const copy = new ArrayBuffer(bytes.byteLength);
    new Uint8Array(copy).set(bytes);
    await writable.write(copy);
  } finally {
    await writable.close();
  }
  return handle;
}

async function openRegisteredDuckDb(options: {
  readonly storage: BrowserDirectoryHandle;
  readonly name: string;
  readonly logicalName?: string | undefined;
  readonly wasmUrl: string;
  readonly workerUrl: string;
  readonly readonly: boolean;
}): Promise<RawDuckDb> {
  const worker = new Worker(options.workerUrl, { type: "classic" });
  const database = new AsyncDuckDB(new VoidLogger(), worker);
  const logicalName = options.logicalName ?? `/${options.name}`;
  try {
    await database.instantiate(options.wasmUrl);
    await database.registerFileHandle(
      logicalName,
      await options.storage.getFileHandle(options.name),
      DuckDBDataProtocol.BROWSER_FSACCESS,
      true,
    );
    await database.registerFileHandle(
      `${logicalName}.wal`,
      await options.storage.getFileHandle(`${options.name}.wal`, {
        create: true,
      }),
      DuckDBDataProtocol.BROWSER_FSACCESS,
      true,
    );
    await database.open({
      path: logicalName,
      accessMode: options.readonly
        ? DuckDBAccessMode.READ_ONLY
        : DuckDBAccessMode.READ_WRITE,
      useDirectIO: true,
      arrowLosslessConversion: true,
    });
    return { database, connection: await database.connect() };
  } catch (error) {
    await database.terminate().catch(() => undefined);
    throw error;
  }
}

async function inspectMainWithoutWal(options: {
  readonly storage: BrowserDirectoryHandle;
  readonly sourceMain: Uint8Array;
  readonly sourceLogicalName: string;
  readonly wasmUrl: string;
  readonly workerUrl: string;
  readonly runId: string;
}): Promise<{ readonly rows: string; readonly views: string }> {
  const name = `.consultchimps-main-only-${options.runId}.duckdb`;
  const walName = `${name}.wal`;
  await replaceFile(options.storage, name, options.sourceMain);
  await replaceFile(options.storage, walName, new Uint8Array());
  const raw = await openRegisteredDuckDb({
    storage: options.storage,
    name,
    logicalName: options.sourceLogicalName,
    wasmUrl: options.wasmUrl,
    workerUrl: options.workerUrl,
    readonly: true,
  });
  try {
    const rows = queryScalar(
      await raw.connection.query("SELECT count(*) AS value FROM Inventory"),
    );
    const views = queryScalar(
      await raw.connection.query(
        "SELECT count(*) AS value FROM duckdb_views() WHERE database_name = current_database() AND schema_name = 'main' AND view_name = 'wal_inventory'",
      ),
    );
    return { rows: String(rows), views: String(views) };
  } finally {
    await raw.connection.close().catch(() => undefined);
    await raw.database.terminate();
    await options.storage.removeEntry(walName).catch(() => undefined);
    await options.storage.removeEntry(name).catch(() => undefined);
  }
}

async function inspectWritablePair(options: {
  readonly storage: BrowserDirectoryHandle;
  readonly sourceMain: Uint8Array;
  readonly sourceWal: Uint8Array;
  readonly sourceLogicalName: string;
  readonly wasmUrl: string;
  readonly workerUrl: string;
  readonly runId: string;
}): Promise<{ readonly rows: string; readonly views: string }> {
  const name = `.consultchimps-pair-${options.runId}.duckdb`;
  const walName = `${name}.wal`;
  await replaceFile(options.storage, name, options.sourceMain);
  await replaceFile(options.storage, walName, options.sourceWal);
  const raw = await openRegisteredDuckDb({
    storage: options.storage,
    name,
    logicalName: options.sourceLogicalName,
    wasmUrl: options.wasmUrl,
    workerUrl: options.workerUrl,
    readonly: false,
  });
  try {
    const rows = queryScalar(
      await raw.connection.query("SELECT count(*) AS value FROM Inventory"),
    );
    const views = queryScalar(
      await raw.connection.query(
        "SELECT count(*) AS value FROM duckdb_views() WHERE database_name = current_database() AND schema_name = 'main' AND view_name = 'wal_inventory'",
      ),
    );
    return { rows: String(rows), views: String(views) };
  } finally {
    await raw.database.terminate();
    await options.storage.removeEntry(walName).catch(() => undefined);
    await options.storage.removeEntry(name).catch(() => undefined);
  }
}

async function storedNames(
  storage: BrowserDirectoryHandle,
): Promise<readonly string[]> {
  const names: string[] = [];
  for await (const entry of storage.values()) names.push(entry.name);
  return names.sort();
}

async function atStage<T>(
  name: string,
  operation: () => Promise<T>,
): Promise<T> {
  try {
    return await operation();
  } catch (cause) {
    const message = cause instanceof Error ? cause.message : String(cause);
    throw new Error(`${name}: ${message}`, { cause });
  }
}

async function run(request: RunRequest) {
  const wasmUrl = new URL(
    "/database-wasm/duckdb-eh.wasm",
    globalThis.location.href,
  ).href;
  const workerUrl = new URL(
    "/database-wasm/duckdb-browser-eh.worker.js",
    globalThis.location.href,
  ).href;
  const directoryName = `consultchimps-duckdb-wal-${request.runId}`;
  const name = `source-${request.runId}.duckdb`;
  const storage = await directory(directoryName);
  const beforeMain = base64ToBytes(request.sourceMainBase64);
  const beforeWal = base64ToBytes(request.sourceWalBase64);
  await replaceFile(storage, name, beforeMain);
  await replaceFile(storage, `${name}.wal`, beforeWal);
  const runtime = await configureBrowserDatabaseRuntime({
    sqlite: {
      wasmUrl: new URL("/database-wasm/sqlite3.wasm", globalThis.location.href)
        .href,
      directory: `/consultchimps-duckdb-wal-${request.runId}`,
      initialCapacity: 8,
    },
    duckdb: { wasmUrl, workerUrl },
    opfsDirectory: directoryName,
  });
  const mainOnly = await atStage("inspect main without WAL", () =>
    inspectMainWithoutWal({
      storage,
      sourceMain: beforeMain,
      sourceLogicalName: `/${name}`,
      wasmUrl,
      workerUrl,
      runId: request.runId,
    }),
  );
  const writablePair = await atStage("inspect writable main and WAL pair", () =>
    inspectWritablePair({
      storage,
      sourceMain: beforeMain,
      sourceWal: beforeWal,
      sourceLogicalName: `/${name}`,
      wasmUrl,
      workerUrl,
      runId: request.runId,
    }),
  );

  const readonly = await atStage("open public read-only database", () =>
    runtime.openDatabase({ name, readonly: true }),
  );
  const readonlyInspection = await atStage(
    "inspect public read-only database",
    () => inspectDatabase({ database: readonly }),
  );
  const failedDestination = new FailingFile(
    `failed-export-${request.runId}.duckdb`,
  );
  let failedExportMessage = "";
  try {
    await runtime.exportDatabase({
      database: readonly,
      name: failedDestination.name,
      destination: failedDestination,
      format: "duckdb",
    });
  } catch (error) {
    failedExportMessage =
      error instanceof Error ? error.message : String(error);
  }
  const afterFailureMain = await fileBytes(storage, name);
  const afterFailureWal = await fileBytes(storage, `${name}.wal`);
  const namesAfterFailure = await storedNames(storage);
  const destination = new MemoryFile(`export-${request.runId}.duckdb`);
  try {
    await runtime.exportDatabase({
      database: readonly,
      name: destination.name,
      destination,
      format: "duckdb",
    });
  } finally {
    await readonly.close();
  }

  const afterMain = await fileBytes(storage, name);
  const afterWal = await fileBytes(storage, `${name}.wal`);
  const names = await storedNames(storage);
  return {
    walSize: beforeWal.byteLength,
    mainOnly,
    writablePair,
    readonlyRows: String(readonlyInspection.tables[0]?.rowCount ?? -1n),
    failedExportMessage,
    failedDestinationSize: failedDestination.size,
    sourceMainPreservedAfterFailure: bytesEqual(beforeMain, afterFailureMain),
    sourceWalPreservedAfterFailure: bytesEqual(beforeWal, afterFailureWal),
    snapshotFilesAfterFailure: namesAfterFailure.filter((entry) =>
      entry.startsWith(".consultchimps-export-snapshot-"),
    ),
    sourceMainPreserved: bytesEqual(beforeMain, afterMain),
    sourceWalPreserved: bytesEqual(beforeWal, afterWal),
    snapshotFiles: names.filter((entry) =>
      entry.startsWith(".consultchimps-export-snapshot-"),
    ),
    exportedSize: destination.size,
    exportedBase64: bytesToBase64(destination.bytes()),
  };
}

function errorDetails(error: unknown): {
  readonly code?: string;
  readonly message: string;
} {
  const code =
    typeof error === "object" && error !== null
      ? Reflect.get(error, "code")
      : undefined;
  return {
    ...(typeof code === "string" ? { code } : {}),
    message: error instanceof Error ? error.message : String(error),
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
        error: errorDetails(error),
      }),
  );
};
