import sqlite3InitModule, { type Sqlite3Static } from "@sqlite.org/sqlite-wasm";

import {
  throwIfAborted,
  type OperationControlOptions,
  type OperationResult,
  type RandomAccessFile,
  type RandomAccessSource,
} from "@consultchimps/core";

import {
  createDatabaseHandle,
  engineOf,
  initializeSchema,
  openDatabaseHandle,
  type CreateDatabaseResult,
  type Database,
} from "./database.js";
import { executeConversion, planConversion } from "./conversion.js";
import { BrowserDuckDbEngine } from "./engines/duckdb/browser.js";
import { BrowserSqliteEngine } from "./engines/sqlite/browser.js";
import { databaseError } from "./errors.js";
import { inspectAppliedImportPlan } from "./import/history.js";
import {
  createPreparedImportHandle,
  openPreparedImportHandle,
  preparedRef,
  type PreparedImport,
} from "./prepared.js";
import type { DatabaseFormat, DatabaseSchema } from "./schema.js";
import type { ImportRecipe } from "./import/types.js";
import { readSchemaFingerprint } from "./records.js";

const SQLITE_HEADER = new TextEncoder().encode("SQLite format 3\0");
const COPY_CHUNK_BYTES = 1024 * 1024;
const SAH_HEADER_BYTES = 4096;
const SAH_PATH_BYTES = 512;

interface StoredDatabase {
  readonly name: string;
  readonly format: DatabaseFormat;
  readonly duckdb?: BrowserDuckDbEngine | undefined;
}

interface BrowserWritable {
  write(data: {
    readonly type: "write";
    readonly position: number;
    readonly data: Uint8Array;
  }): Promise<void>;
  truncate(size: number): Promise<void>;
  close(): Promise<void>;
}

interface BrowserFileHandle {
  readonly kind: "file";
  readonly name: string;
  getFile(): Promise<Blob>;
  createWritable(): Promise<BrowserWritable>;
}

interface BrowserDirectoryHandle {
  readonly kind: "directory";
  readonly name: string;
  values(): AsyncIterableIterator<BrowserDirectoryHandle | BrowserFileHandle>;
  getDirectoryHandle(
    name: string,
    options?: { readonly create?: boolean },
  ): Promise<BrowserDirectoryHandle>;
  getFileHandle(
    name: string,
    options?: { readonly create?: boolean },
  ): Promise<BrowserFileHandle>;
  removeEntry(name: string): Promise<void>;
}

function browserStorage(): {
  getDirectory(): Promise<BrowserDirectoryHandle>;
} {
  const navigatorValue = globalThis.navigator as unknown as {
    readonly storage?: {
      getDirectory?: () => Promise<BrowserDirectoryHandle>;
    };
  };
  if (navigatorValue.storage?.getDirectory === undefined) {
    throw databaseError(
      "DB_BROWSER_STORAGE_UNAVAILABLE",
      "This browser does not provide origin-private file storage.",
    );
  }
  return {
    getDirectory: navigatorValue.storage.getDirectory.bind(
      navigatorValue.storage,
    ),
  };
}

export interface BrowserDatabaseRuntimeOptions {
  readonly sqlite: {
    readonly wasmUrl: string;
    readonly directory: string;
    readonly initialCapacity?: number | undefined;
  };
  readonly duckdb: {
    readonly wasmUrl: string;
    readonly workerUrl: string;
  };
  readonly opfsDirectory?: string | undefined;
}

export interface CreateBrowserDatabaseOptions extends OperationControlOptions {
  readonly name: string;
  readonly format: DatabaseFormat;
  readonly schema?: DatabaseSchema | undefined;
  readonly overwrite?: boolean | undefined;
}

export interface ImportBrowserDatabaseOptions extends OperationControlOptions {
  readonly name: string;
  readonly source: RandomAccessSource;
  readonly overwrite?: boolean | undefined;
}

export interface CreateBrowserPreparedImportOptions extends OperationControlOptions {
  readonly name: string;
  readonly database: Database;
  readonly recipe: ImportRecipe;
  readonly baselineRevision: bigint;
  readonly overwrite?: boolean | undefined;
}

export interface ExportBrowserDatabaseOptions extends OperationControlOptions {
  readonly database: Database;
  readonly name: string;
  readonly destination: RandomAccessFile;
  readonly format: DatabaseFormat;
  readonly overwrite?: boolean | undefined;
}

export interface BrowserPreparedImportSummary {
  readonly name: string;
  readonly id: string;
  readonly databaseId: string;
  readonly application: "applied" | "pending";
}

export interface BrowserPreparedImportListing {
  readonly imports: readonly BrowserPreparedImportSummary[];
  readonly ignored: readonly string[];
}

export interface BrowserDatabaseRuntime {
  createDatabase(
    options: CreateBrowserDatabaseOptions,
  ): Promise<CreateDatabaseResult>;
  openDatabase(options: {
    readonly name: string;
    readonly readonly?: boolean | undefined;
  }): Promise<Database>;
  importDatabase(options: ImportBrowserDatabaseOptions): Promise<Database>;
  createPreparedImport(
    options: CreateBrowserPreparedImportOptions,
  ): Promise<PreparedImport>;
  openPreparedImport(options: {
    readonly name: string;
  }): Promise<PreparedImport>;
  listPreparedImports(options: {
    readonly database: Database;
  }): Promise<BrowserPreparedImportListing>;
  exportDatabase(
    options: ExportBrowserDatabaseOptions,
  ): Promise<OperationResult<"bytesWritten">>;
}

function storageName(name: string): string {
  const normalized = name.normalize("NFKC");
  if (
    normalized.length === 0 ||
    normalized === "." ||
    normalized === ".." ||
    normalized.includes("/") ||
    normalized.includes("\\") ||
    !/^[\p{L}\p{N}._-]+$/u.test(normalized)
  ) {
    throw databaseError(
      "DB_INVALID_STORAGE_NAME",
      "Choose a database name that uses letters, numbers, dots, underscores, or hyphens.",
      { name },
    );
  }
  return normalized;
}

function sqliteName(name: string): string {
  return `/${storageName(name)}`;
}

async function opfsDirectory(name: string): Promise<BrowserDirectoryHandle> {
  let directory = await browserStorage().getDirectory();
  for (const segment of name.split("/").filter((part) => part.length > 0)) {
    directory = await directory.getDirectoryHandle(storageName(segment), {
      create: true,
    });
  }
  return directory;
}

async function fileExists(
  directory: BrowserDirectoryHandle,
  name: string,
): Promise<boolean> {
  try {
    await directory.getFileHandle(name);
    return true;
  } catch (error) {
    if (error instanceof DOMException && error.name === "NotFoundError") {
      return false;
    }
    throw error;
  }
}

async function replaceDuckDbFiles(
  directory: BrowserDirectoryHandle,
  name: string,
  overwrite: boolean,
): Promise<{ file: BrowserFileHandle; wal: BrowserFileHandle }> {
  const exists = await fileExists(directory, name);
  if (exists && !overwrite) {
    throw databaseError(
      "DB_OUTPUT_EXISTS",
      "A browser database with this name already exists. Choose another name or allow replacement.",
      { name },
    );
  }
  if (exists) await directory.removeEntry(name);
  if (await fileExists(directory, `${name}.wal`)) {
    await directory.removeEntry(`${name}.wal`);
  }
  return {
    file: await directory.getFileHandle(name, { create: true }),
    wal: await directory.getFileHandle(`${name}.wal`, { create: true }),
  };
}

async function writeSource(
  source: RandomAccessSource,
  handle: BrowserFileHandle,
  signal?: AbortSignal,
  onProgress?: OperationControlOptions["onProgress"],
): Promise<void> {
  const writable = await handle.createWritable();
  try {
    for (let offset = 0; offset < source.size; offset += COPY_CHUNK_BYTES) {
      throwIfAborted(signal, "db.browser.import");
      const length = Math.min(COPY_CHUNK_BYTES, source.size - offset);
      const bytes = await source.readAt(offset, length, signal);
      if (bytes.length !== length) {
        throw databaseError(
          "DB_SOURCE_SHORT_READ",
          "The database source ended while it was being copied.",
          { offset, expected: length, actual: bytes.length },
        );
      }
      await writable.write({ type: "write", position: offset, data: bytes });
      onProgress?.({
        operation: "db.browser.import",
        stage: "copying",
        completed: offset + bytes.length,
        total: source.size,
        detail: source.name,
      });
    }
    await writable.truncate(source.size);
  } finally {
    await writable.close();
  }
}

async function sqlitePoolDirectory(
  path: string,
): Promise<BrowserDirectoryHandle> {
  let directory = await browserStorage().getDirectory();
  for (const segment of path.split("/").filter((part) => part.length > 0)) {
    directory = await directory.getDirectoryHandle(storageName(segment));
  }
  return directory.getDirectoryHandle(".opaque");
}

async function findSqlitePoolFile(
  directory: BrowserDirectoryHandle,
  path: string,
): Promise<Blob> {
  for await (const entry of directory.values()) {
    if (entry.kind !== "file") continue;
    const file = await entry.getFile();
    if (file.size <= SAH_HEADER_BYTES) continue;
    const pathBytes = new Uint8Array(
      await file.slice(0, SAH_PATH_BYTES).arrayBuffer(),
    );
    const end = pathBytes.indexOf(0);
    const associatedPath = new TextDecoder().decode(
      pathBytes.subarray(0, end < 0 ? pathBytes.length : end),
    );
    if (associatedPath === path) return file;
  }
  throw databaseError(
    "DB_BROWSER_STORAGE_MISSING",
    "The SQLite working database could not be found in browser storage.",
    { name: path },
  );
}

async function copySqlitePoolFile(options: {
  readonly poolDirectory: BrowserDirectoryHandle;
  readonly name: string;
  readonly destination: RandomAccessFile;
  readonly signal?: AbortSignal | undefined;
  readonly onProgress?: OperationControlOptions["onProgress"];
}): Promise<number> {
  const file = await findSqlitePoolFile(
    options.poolDirectory,
    sqliteName(options.name),
  );
  const size = file.size - SAH_HEADER_BYTES;
  if (size < 512 || size % 512 !== 0) {
    throw databaseError(
      "DB_CORRUPT_DATABASE",
      "The SQLite working database has an invalid file size.",
    );
  }
  await options.destination.truncate(0);
  for (let offset = 0; offset < size; offset += COPY_CHUNK_BYTES) {
    throwIfAborted(options.signal, "db.browser.export");
    const length = Math.min(COPY_CHUNK_BYTES, size - offset);
    const bytes = new Uint8Array(
      await file
        .slice(SAH_HEADER_BYTES + offset, SAH_HEADER_BYTES + offset + length)
        .arrayBuffer(),
    );
    if (
      offset === 0 &&
      !SQLITE_HEADER.every((byte, index) => bytes[index] === byte)
    ) {
      throw databaseError(
        "DB_CORRUPT_DATABASE",
        "The SQLite working database has an invalid file header.",
      );
    }
    await options.destination.writeAt(offset, bytes);
    options.onProgress?.({
      operation: "db.browser.export",
      stage: "copying",
      completed: offset + bytes.length,
      total: size,
      detail: options.name,
    });
  }
  await options.destination.truncate(size);
  return size;
}

async function detectSourceFormat(
  source: RandomAccessSource,
  signal?: AbortSignal,
): Promise<DatabaseFormat> {
  const header = await source.readAt(
    0,
    Math.min(SQLITE_HEADER.length, source.size),
    signal,
  );
  return header.length === SQLITE_HEADER.length &&
    header.every((byte, index) => byte === SQLITE_HEADER[index])
    ? "sqlite"
    : "duckdb";
}

export async function configureBrowserDatabaseRuntime(
  options: BrowserDatabaseRuntimeOptions,
): Promise<BrowserDatabaseRuntime> {
  const initialize = sqlite3InitModule as unknown as (options: {
    readonly locateFile: () => string;
  }) => Promise<Sqlite3Static>;
  const sqlite = await initialize({ locateFile: () => options.sqlite.wasmUrl });
  const pool = await sqlite.installOpfsSAHPoolVfs({
    directory: options.sqlite.directory,
    ...(options.sqlite.initialCapacity === undefined
      ? {}
      : { initialCapacity: options.sqlite.initialCapacity }),
  });
  if (options.sqlite.initialCapacity !== undefined) {
    await pool.reserveMinimumCapacity(options.sqlite.initialCapacity);
  }
  const duckDirectory = await opfsDirectory(
    options.opfsDirectory ?? "consultchimps-databases",
  );
  const sqliteDirectory = await sqlitePoolDirectory(options.sqlite.directory);
  const stored = new WeakMap<Database, StoredDatabase>();

  const sqliteEngine = (name: string): BrowserSqliteEngine =>
    new BrowserSqliteEngine(sqlite, new pool.OpfsSAHPoolDb(sqliteName(name)));

  const duckEngine = async (
    name: string,
    readonly = false,
  ): Promise<BrowserDuckDbEngine> => {
    const file = await duckDirectory.getFileHandle(name);
    const wal = await duckDirectory.getFileHandle(`${name}.wal`, {
      create: true,
    });
    return BrowserDuckDbEngine.open({
      ...options.duckdb,
      storageName: name,
      fileHandle: file,
      walHandle: wal,
      readonly,
    });
  };

  const openDatabase = async (open: {
    readonly name: string;
    readonly readonly?: boolean | undefined;
  }): Promise<Database> => {
    const name = storageName(open.name);
    const format: DatabaseFormat = pool
      .getFileNames()
      .includes(sqliteName(name))
      ? "sqlite"
      : "duckdb";
    const engine =
      format === "sqlite"
        ? sqliteEngine(name)
        : await duckEngine(name, open.readonly);
    try {
      const database = await openDatabaseHandle(engine);
      stored.set(database, {
        name,
        format,
        ...(format === "duckdb"
          ? { duckdb: engine as BrowserDuckDbEngine }
          : {}),
      });
      return database;
    } catch (error) {
      await engine.close().catch(() => undefined);
      throw error;
    }
  };

  const createDatabase = async (
    create: CreateBrowserDatabaseOptions,
  ): Promise<CreateDatabaseResult> => {
    throwIfAborted(create.signal, "db.browser.create");
    const name = storageName(create.name);
    let engine: BrowserSqliteEngine | BrowserDuckDbEngine;
    if (create.format === "sqlite") {
      const target = sqliteName(name);
      if (pool.getFileNames().includes(target) && create.overwrite !== true) {
        throw databaseError(
          "DB_OUTPUT_EXISTS",
          "A browser database with this name already exists. Choose another name or allow replacement.",
          { name },
        );
      }
      if (create.overwrite === true) pool.unlink(target);
      engine = sqliteEngine(name);
    } else {
      const handles = await replaceDuckDbFiles(
        duckDirectory,
        name,
        create.overwrite === true,
      );
      engine = await BrowserDuckDbEngine.open({
        ...options.duckdb,
        storageName: name,
        fileHandle: handles.file,
        walHandle: handles.wal,
      });
    }
    try {
      const database = await createDatabaseHandle(engine);
      const tablesCreated = await initializeSchema(database, create.schema);
      stored.set(database, {
        name,
        format: create.format,
        ...(create.format === "duckdb"
          ? { duckdb: engine as BrowserDuckDbEngine }
          : {}),
      });
      return {
        database,
        result: {
          operation: "db.create",
          artifacts: [],
          warnings: [],
          metrics: { tablesCreated },
        },
      };
    } catch (error) {
      await engine.close().catch(() => undefined);
      throw error;
    }
  };

  const exportStored = async (
    database: Database,
    source: StoredDatabase,
    destination: RandomAccessFile,
    controls: OperationControlOptions,
  ): Promise<number> => {
    if (source.format === "sqlite") {
      await database.checkpoint();
      return copySqlitePoolFile({
        poolDirectory: sqliteDirectory,
        name: source.name,
        destination,
        signal: controls.signal,
        onProgress: controls.onProgress,
      });
    }
    return source.duckdb!.copyTo(destination, controls.signal);
  };

  const removeStored = async (source: StoredDatabase): Promise<void> => {
    if (source.format === "sqlite") {
      pool.unlink(sqliteName(source.name));
      return;
    }
    if (await fileExists(duckDirectory, source.name)) {
      await duckDirectory.removeEntry(source.name);
    }
    if (await fileExists(duckDirectory, `${source.name}.wal`)) {
      await duckDirectory.removeEntry(`${source.name}.wal`);
    }
  };

  return {
    createDatabase,
    openDatabase,
    async importDatabase(importOptions) {
      const name = storageName(importOptions.name);
      const format = await detectSourceFormat(
        importOptions.source,
        importOptions.signal,
      );
      if (format === "sqlite") {
        const target = sqliteName(name);
        if (
          pool.getFileNames().includes(target) &&
          importOptions.overwrite !== true
        ) {
          throw databaseError(
            "DB_OUTPUT_EXISTS",
            "A browser database with this name already exists. Choose another name or allow replacement.",
            { name },
          );
        }
        if (importOptions.overwrite === true) pool.unlink(target);
        let offset = 0;
        await pool.importDb(target, async () => {
          if (offset >= importOptions.source.size) return undefined;
          throwIfAborted(importOptions.signal, "db.browser.import");
          const length = Math.min(
            COPY_CHUNK_BYTES,
            importOptions.source.size - offset,
          );
          const bytes = await importOptions.source.readAt(
            offset,
            length,
            importOptions.signal,
          );
          offset += bytes.length;
          importOptions.onProgress?.({
            operation: "db.browser.import",
            stage: "copying",
            completed: offset,
            total: importOptions.source.size,
            detail: importOptions.source.name,
          });
          return bytes;
        });
      } else {
        const handles = await replaceDuckDbFiles(
          duckDirectory,
          name,
          importOptions.overwrite === true,
        );
        await writeSource(
          importOptions.source,
          handles.file,
          importOptions.signal,
          importOptions.onProgress,
        );
      }
      return openDatabase({ name });
    },
    async createPreparedImport(create) {
      throwIfAborted(create.signal, "db.browser.plan");
      const name = sqliteName(create.name);
      if (pool.getFileNames().includes(name) && create.overwrite !== true) {
        throw databaseError(
          "DB_OUTPUT_EXISTS",
          "A browser import plan with this name already exists. Choose another name or allow replacement.",
          { name },
        );
      }
      if (create.overwrite === true) pool.unlink(name);
      const engine = new BrowserSqliteEngine(
        sqlite,
        new pool.OpfsSAHPoolDb(name),
      );
      try {
        return await createPreparedImportHandle({
          engine,
          databaseId: create.database.id,
          baselineRevision: create.baselineRevision,
          baselineSchemaFingerprint: await readSchemaFingerprint(
            engineOf(create.database),
            create.database.format,
          ),
          recipe: create.recipe,
        });
      } catch (error) {
        await engine.close().catch(() => undefined);
        pool.unlink(name);
        throw error;
      }
    },
    async openPreparedImport(open) {
      const engine = new BrowserSqliteEngine(
        sqlite,
        new pool.OpfsSAHPoolDb(sqliteName(open.name)),
      );
      try {
        return await openPreparedImportHandle(engine);
      } catch (error) {
        await engine.close().catch(() => undefined);
        throw error;
      }
    },
    async listPreparedImports(listOptions) {
      const names = pool
        .getFileNames()
        .filter(
          (name) =>
            name.startsWith("/.consultchimps-import-") &&
            name.endsWith(".sqlite"),
        )
        .map((name) => name.slice(1))
        .sort();
      const imports: BrowserPreparedImportSummary[] = [];
      const ignored: string[] = [];
      for (const name of names) {
        let engine: BrowserSqliteEngine | undefined;
        let prepared: PreparedImport | undefined;
        try {
          engine = new BrowserSqliteEngine(
            sqlite,
            new pool.OpfsSAHPoolDb(sqliteName(name)),
          );
          prepared = await openPreparedImportHandle(engine);
          const ref = await preparedRef(prepared);
          const applied =
            ref.databaseId === listOptions.database.id &&
            (await inspectAppliedImportPlan({
              database: listOptions.database,
              planId: ref.id,
              planRevision: ref.planRevision,
            })) !== null;
          imports.push({
            name,
            id: prepared.id,
            databaseId: prepared.databaseId,
            application: applied ? "applied" : "pending",
          });
        } catch {
          ignored.push(name);
        } finally {
          if (prepared !== undefined) await prepared.close();
          else await engine?.close().catch(() => undefined);
        }
      }
      return { imports, ignored };
    },
    async exportDatabase(exportOptions) {
      throwIfAborted(exportOptions.signal, "db.browser.export");
      const source = stored.get(exportOptions.database);
      if (source === undefined) {
        throw databaseError(
          "DB_UNKNOWN_BROWSER_DATABASE",
          "This database was not opened by the current browser runtime.",
        );
      }
      let bytesWritten: number;
      if (source.format === exportOptions.format) {
        bytesWritten = await exportStored(
          exportOptions.database,
          source,
          exportOptions.destination,
          exportOptions,
        );
      } else {
        const plan = await planConversion({
          database: exportOptions.database,
          format: exportOptions.format,
        });
        const temporaryName = `.consultchimps-conversion-${globalThis.crypto.randomUUID()}.${exportOptions.format}`;
        const created = await createDatabase({
          name: temporaryName,
          format: exportOptions.format,
          signal: exportOptions.signal,
        });
        const converted = stored.get(created.database);
        if (converted === undefined) {
          await created.database.close();
          throw databaseError(
            "DB_UNKNOWN_BROWSER_DATABASE",
            "The conversion database was not registered by the browser runtime.",
          );
        }
        try {
          await executeConversion({
            source: exportOptions.database,
            target: created.database,
            plan,
            signal: exportOptions.signal,
            onProgress: exportOptions.onProgress,
          });
          bytesWritten = await exportStored(
            created.database,
            converted,
            exportOptions.destination,
            exportOptions,
          );
        } finally {
          await created.database.close().catch(() => undefined);
          await removeStored(converted).catch(() => undefined);
        }
      }
      return {
        operation: "db.export",
        artifacts: [],
        warnings: [],
        metrics: { bytesWritten },
      };
    },
  };
}
