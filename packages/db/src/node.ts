import { lstat, rm, stat } from "node:fs/promises";
import path from "node:path";

import {
  ensureParentDirectory,
  openRandomAccessSource,
} from "@consultchimps/files";
import type {
  OperationControlOptions,
  OperationResult,
} from "@consultchimps/core";
import { isConsultChimpsError, throwIfAborted } from "@consultchimps/core";

import {
  createDatabaseHandle,
  engineOf,
  initializeSchema,
  openDatabaseHandle,
  type CreateDatabaseResult,
  type Database,
  type DatabaseId,
} from "./database.js";
import type { DatabaseEngine } from "./internal/engine.js";
import {
  executeConversion,
  planConversion,
  type ConversionPlan,
} from "./conversion.js";
import { databaseError } from "./errors.js";
import { NodeDuckDbEngine } from "./engines/duckdb/node.js";
import { NodeSqliteEngine } from "./engines/sqlite/node.js";
import {
  PREPARED_METADATA_TABLE,
  createImportBatchHandle,
  openImportBatchHandle,
  type ImportBatch,
} from "./prepared.js";
import { DATABASE_METADATA_TABLE } from "./metadata.js";
import type { DatabaseFormat, DatabaseSchema } from "./schema.js";
import { readSchemaFingerprint } from "./records.js";
import { prepareImport } from "./import/prepare.js";
import { NativeFileRegistry } from "./native-files.js";
import { failAfterNativeArtifactCleanup } from "./native-artifact-cleanup.js";
import type {
  ImportProfile,
  ImportSource,
  PrepareImportOutcome,
} from "./import/types.js";

const SQLITE_HEADER = new TextEncoder().encode("SQLite format 3\0");
const databasePaths = new WeakMap<Database, string>();
const nativeFiles = new NativeFileRegistry();

function retainInitializationFailure(filePath: string, cause: unknown): void {
  if (
    isConsultChimpsError(cause) &&
    (cause.code === "DB_NATIVE_SQLITE_CLEANUP_REQUIRED" ||
      cause.code === "DB_DUCKDB_OPEN_CLEANUP_FAILED")
  ) {
    nativeFiles.retainUnclosedOwner(filePath, cause);
    throw cause;
  }
}

async function closeNativeReadOwner(options: {
  readonly path: string;
  readonly owner: { close(): Promise<void> };
  readonly failure?: { readonly cause: unknown };
}): Promise<void> {
  try {
    await options.owner.close();
  } catch (cleanupFailure) {
    nativeFiles.retainUnclosedOwner(options.path, options.owner);
    throw databaseError(
      "DB_NATIVE_HANDLE_CLEANUP_REQUIRED",
      "The database or import batch could not be closed after opening or inspection. Restart the process before reopening or replacing the reported file. File replacement is blocked in this runtime.",
      { path: options.path, closeFailed: true },
      new AggregateError(
        [
          ...(options.failure === undefined ? [] : [options.failure.cause]),
          cleanupFailure,
        ],
        "Native file access and handle cleanup failed",
      ),
    );
  }
}

function nativeTemporaryPaths(
  temporary: string,
  format: DatabaseFormat,
): readonly string[] {
  return format === "sqlite"
    ? [temporary, `${temporary}-wal`, `${temporary}-shm`]
    : [temporary, `${temporary}.wal`, `${temporary}.tmp`];
}

async function removeNativeTemporary(paths: readonly string[]): Promise<void> {
  const removals = await Promise.allSettled(
    paths.map((filePath) =>
      rm(filePath, {
        force: true,
        ...(filePath.endsWith(".tmp") ? { recursive: true } : {}),
      }),
    ),
  );
  const failures = removals.flatMap((result) =>
    result.status === "rejected" ? [result.reason] : [],
  );
  if (failures.length > 0) {
    throw new AggregateError(failures, "Native temporary file removal failed");
  }
}

async function assertNativeTemporaryAvailable(
  paths: readonly string[],
): Promise<void> {
  for (const filePath of paths) {
    try {
      await lstat(filePath);
    } catch (error) {
      if (
        error instanceof Error &&
        "code" in error &&
        error.code === "ENOENT"
      ) {
        continue;
      }
      throw error;
    }
    throw databaseError(
      "DB_OUTPUT_EXISTS",
      "A private database working path already exists. Resolve or remove that path before retrying.",
      { path: filePath },
    );
  }
}

export type NodeDatabaseFileKind =
  | {
      readonly kind: "database";
      readonly format: DatabaseFormat;
      readonly databaseId: DatabaseId;
    }
  | {
      readonly kind: "prepared-import";
      readonly format: "sqlite";
      readonly databaseId: DatabaseId;
    }
  | {
      readonly kind: "unmanaged-database";
      readonly format: DatabaseFormat;
      readonly tables: readonly UnmanagedDatabaseTable[];
    };

export interface UnmanagedDatabaseColumn {
  readonly name: string;
  readonly storageType: string;
  readonly nullable: boolean;
}

export interface UnmanagedDatabaseTable {
  readonly name: string;
  readonly columns: readonly UnmanagedDatabaseColumn[];
}

async function inspectUnmanagedTables(
  engine: DatabaseEngine,
  format: DatabaseFormat,
): Promise<readonly UnmanagedDatabaseTable[]> {
  const tableRows = await engine.query(
    format === "sqlite"
      ? "SELECT name AS table_name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
      : "SELECT table_name FROM information_schema.tables WHERE table_schema = 'main' AND table_type = 'BASE TABLE' ORDER BY table_name",
  );
  const tables: UnmanagedDatabaseTable[] = [];
  for (const tableRow of tableRows) {
    const name = tableRow["table_name"];
    if (typeof name !== "string") continue;
    const columnRows =
      format === "sqlite"
        ? await engine.query(
            'SELECT name, type, "notnull" AS not_null, pk FROM pragma_table_info(?) ORDER BY cid',
            [name],
          )
        : await engine.query(
            "SELECT column_name, data_type, is_nullable FROM information_schema.columns WHERE table_schema = 'main' AND table_name = ? ORDER BY ordinal_position",
            [name],
          );
    tables.push({
      name,
      columns: columnRows.flatMap((row) => {
        const columnName =
          format === "sqlite" ? row["name"] : row["column_name"];
        const storageType =
          format === "sqlite" ? row["type"] : row["data_type"];
        if (typeof columnName !== "string" || typeof storageType !== "string") {
          return [];
        }
        const nullable =
          format === "sqlite"
            ? row["not_null"] !== 1n &&
              row["not_null"] !== 1 &&
              row["pk"] !== 1n &&
              row["pk"] !== 1
            : row["is_nullable"] === "YES";
        return [{ name: columnName, storageType, nullable }];
      }),
    });
  }
  return tables;
}

async function detectedFormat(filePath: string): Promise<DatabaseFormat> {
  const source = await openRandomAccessSource(filePath);
  try {
    const header = await source.readAt(0, SQLITE_HEADER.length);
    if (
      header.length === SQLITE_HEADER.length &&
      header.every((byte, index) => byte === SQLITE_HEADER[index])
    ) {
      return "sqlite";
    }
    // DuckDB's storage header is engine-versioned. Let DuckDB validate the file
    // rather than freezing an undocumented byte offset as our format contract.
    return "duckdb";
  } finally {
    await source.close();
  }
}

async function openEngine(
  filePath: string,
  format: DatabaseFormat,
  readonly = false,
) {
  try {
    return format === "sqlite"
      ? NodeSqliteEngine.open(filePath, readonly)
      : await NodeDuckDbEngine.create(filePath, readonly);
  } catch (cause) {
    retainInitializationFailure(filePath, cause);
    throw databaseError(
      "DB_OPEN_FAILED",
      "The database could not be opened. Check that it is a supported SQLite or DuckDB file and that you have access to it.",
      { path: filePath, format },
      cause,
    );
  }
}

async function inspectFileKindUnlocked(options: {
  readonly path: string;
}): Promise<NodeDatabaseFileKind> {
  const input = path.resolve(options.path);
  const format = await detectedFormat(input);
  let engine: DatabaseEngine | undefined;
  let owner: { close(): Promise<void> } | undefined;
  let failure: { readonly cause: unknown } | undefined;
  try {
    try {
      engine = await openEngine(input, format, true);
      owner = engine;
    } catch (cause) {
      retainInitializationFailure(input, cause);
      throw databaseError(
        "DB_UNSUPPORTED_FILE_FORMAT",
        "This file is not a supported ConsultChimps database or import batch.",
        { path: input },
        cause,
      );
    }
    const markerRows = await engine.query(
      format === "sqlite"
        ? "SELECT name AS table_name FROM sqlite_master WHERE type = 'table' AND name IN (?, ?)"
        : "SELECT table_name FROM information_schema.tables WHERE table_schema = 'main' AND table_name IN (?, ?)",
      [DATABASE_METADATA_TABLE, PREPARED_METADATA_TABLE],
    );
    const markers = new Set(
      markerRows.flatMap((row) =>
        typeof row["table_name"] === "string" ? [row["table_name"]] : [],
      ),
    );
    const isDatabase = markers.has(DATABASE_METADATA_TABLE);
    const isPrepared = markers.has(PREPARED_METADATA_TABLE);
    if (isDatabase && isPrepared) {
      throw databaseError(
        "DB_UNSUPPORTED_FILE_FORMAT",
        "This file has conflicting ConsultChimps database and import batch markers.",
        { path: input },
      );
    }
    if (!isDatabase && !isPrepared) {
      return {
        kind: "unmanaged-database",
        format,
        tables: await inspectUnmanagedTables(engine, format),
      };
    }
    if (isPrepared) {
      if (format !== "sqlite") {
        throw databaseError(
          "DB_INVALID_PREPARED_IMPORT",
          "A ConsultChimps import batch must use its supported SQLite storage format.",
        );
      }
      const prepared = await openImportBatchHandle(engine);
      owner = prepared;
      return {
        kind: "prepared-import",
        format: "sqlite",
        databaseId: prepared.databaseId,
      };
    }
    const database = await openDatabaseHandle(engine);
    owner = database;
    return {
      kind: "database",
      format: database.format,
      databaseId: database.id,
    };
  } catch (cause) {
    failure = { cause };
    throw cause;
  } finally {
    if (owner !== undefined) {
      await closeNativeReadOwner({
        path: input,
        owner,
        ...(failure === undefined ? {} : { failure }),
      });
    }
  }
}

export async function inspectFileKind(options: {
  readonly path: string;
}): Promise<NodeDatabaseFileKind> {
  return nativeFiles.inspect(() => inspectFileKindUnlocked(options));
}

export interface CreateNodeDatabaseOptions {
  readonly path: string;
  readonly format: DatabaseFormat;
  readonly schema?: DatabaseSchema | undefined;
  readonly overwrite?: boolean | undefined;
}

export async function createDatabase(
  options: CreateNodeDatabaseOptions,
): Promise<CreateDatabaseResult> {
  const output = await ensureParentDirectory(options.path);
  const publication = await nativeFiles.planPublication({
    output,
    inputs: [],
    overwrite: options.overwrite,
  });
  const temporary = path.join(
    path.dirname(output),
    `.${path.basename(output)}.cc-create-${globalThis.crypto.randomUUID()}`,
  );
  const storagePaths = nativeTemporaryPaths(temporary, options.format);
  await assertNativeTemporaryAvailable(storagePaths);
  let stage: "preparation" | "publication" = "preparation";
  let engine: DatabaseEngine | undefined;
  let database: Database | undefined;
  let candidateClosed = false;
  try {
    engine =
      options.format === "sqlite"
        ? NodeSqliteEngine.create(temporary)
        : await NodeDuckDbEngine.create(temporary);
    database = await createDatabaseHandle(engine);
    engine = undefined;
    const tablesCreated = await initializeSchema(database, options.schema);
    await database.checkpoint();
    await database.close();
    candidateClosed = true;
    database = undefined;
    stage = "publication";
    const opened = await nativeFiles.publishAndOpen({
      temporary,
      plan: publication,
      open: () => openDatabaseUnlocked({ path: output }),
    });
    return {
      database: opened,
      result: {
        operation: "db.create",
        artifacts: [
          {
            kind: "file",
            path: output,
            mediaType:
              options.format === "sqlite"
                ? "application/vnd.sqlite3"
                : "application/vnd.duckdb",
          },
        ],
        warnings: [],
        metrics: { tablesCreated },
      },
    };
  } catch (error) {
    const candidateDatabase = database;
    const candidateEngine = engine;
    return failAfterNativeArtifactCleanup({
      operation: "create",
      stage,
      kind: "database",
      format: options.format,
      temporaryPath: temporary,
      storagePaths,
      cause: error,
      close: candidateClosed
        ? []
        : candidateDatabase !== undefined
          ? [() => candidateDatabase.close()]
          : candidateEngine === undefined
            ? []
            : [() => candidateEngine.close()],
      remove: () => removeNativeTemporary(storagePaths),
    });
  }
}

async function openDatabaseUnlocked(options: {
  readonly path: string;
  readonly readonly?: boolean | undefined;
}): Promise<Database> {
  const input = path.resolve(options.path);
  const format = await detectedFormat(input);
  const engine = await openEngine(input, format, options.readonly);
  try {
    const database = await openDatabaseHandle(engine);
    databasePaths.set(database, input);
    return database;
  } catch (error) {
    await closeNativeReadOwner({
      path: input,
      owner: engine,
      failure: { cause: error },
    });
    throw error;
  }
}

export async function openDatabase(options: {
  readonly path: string;
  readonly readonly?: boolean | undefined;
}): Promise<Database> {
  const input = path.resolve(options.path);
  return nativeFiles.open(input, () =>
    openDatabaseUnlocked({ ...options, path: input }),
  );
}

export async function createImportBatch(options: {
  readonly path: string;
  readonly database: Database;
  readonly profile: ImportProfile;
  readonly baselineRevision: bigint;
  readonly overwrite?: boolean | undefined;
  readonly protectedInputPaths?: readonly string[] | undefined;
}): Promise<ImportBatch> {
  const output = await ensureParentDirectory(options.path);
  const databasePath = databasePaths.get(options.database);
  const publication = await nativeFiles.planPublication({
    output,
    inputs: [
      ...(databasePath === undefined ? [] : [databasePath]),
      ...(options.protectedInputPaths ?? []),
    ],
    overwrite: options.overwrite,
  });
  const temporary = path.join(
    path.dirname(output),
    `.${path.basename(output)}.cc-plan-${globalThis.crypto.randomUUID()}`,
  );
  const storagePaths = nativeTemporaryPaths(temporary, "sqlite");
  await assertNativeTemporaryAvailable(storagePaths);
  let stage: "preparation" | "publication" = "preparation";
  let engine: NodeSqliteEngine | undefined;
  let prepared: ImportBatch | undefined;
  let candidateClosed = false;
  try {
    engine = NodeSqliteEngine.create(temporary);
    prepared = await createImportBatchHandle({
      engine,
      databaseId: options.database.id,
      baselineRevision: options.baselineRevision,
      baselineSchemaFingerprint: await readSchemaFingerprint(
        engineOf(options.database),
        options.database.format,
      ),
      profile: options.profile,
    });
    engine = undefined;
    await prepared.close();
    candidateClosed = true;
    prepared = undefined;
    stage = "publication";
    return await nativeFiles.publishAndOpen({
      temporary,
      plan: publication,
      open: () => openImportBatchUnlocked({ path: output }),
    });
  } catch (error) {
    const candidatePrepared = prepared;
    const candidateEngine = engine;
    return failAfterNativeArtifactCleanup({
      operation: "plan",
      stage,
      kind: "prepared",
      format: "sqlite",
      temporaryPath: temporary,
      storagePaths,
      cause: error,
      close: candidateClosed
        ? []
        : candidatePrepared !== undefined
          ? [() => candidatePrepared.close()]
          : candidateEngine === undefined
            ? []
            : [() => candidateEngine.close()],
      remove: () => removeNativeTemporary(storagePaths),
    });
  }
}

export interface PrepareImportFileOptions extends OperationControlOptions {
  readonly path: string;
  readonly database: Database;
  readonly sources: readonly ImportSource[];
  readonly profile: ImportProfile;
  readonly baselineRevision: bigint;
  readonly overwrite?: boolean | undefined;
  readonly protectedInputPaths?: readonly string[] | undefined;
}

export async function prepareImportFile(
  options: PrepareImportFileOptions,
): Promise<PrepareImportOutcome> {
  throwIfAborted(options.signal, "db.import.prepare");
  const output = await ensureParentDirectory(options.path);
  const databasePath = databasePaths.get(options.database);
  const publication = await nativeFiles.planPublication({
    output,
    inputs: [
      ...(databasePath === undefined ? [] : [databasePath]),
      ...(options.protectedInputPaths ?? []),
    ],
    overwrite: options.overwrite,
  });
  const temporary = path.join(
    path.dirname(output),
    `.${path.basename(output)}.cc-prepare-${globalThis.crypto.randomUUID()}`,
  );
  const storagePaths = nativeTemporaryPaths(temporary, "sqlite");
  await assertNativeTemporaryAvailable(storagePaths);
  let stage: "preparation" | "publication" = "preparation";
  let engine: NodeSqliteEngine | undefined;
  let prepared: ImportBatch | undefined;
  let candidateClosed = false;
  try {
    await nativeFiles.planPublication({
      output: temporary,
      inputs: [
        ...(databasePath === undefined ? [] : [databasePath]),
        ...(options.protectedInputPaths ?? []),
      ],
      overwrite: false,
    });
    engine = NodeSqliteEngine.create(temporary);
    prepared = await createImportBatchHandle({
      engine,
      databaseId: options.database.id,
      baselineRevision: options.baselineRevision,
      baselineSchemaFingerprint: await readSchemaFingerprint(
        engineOf(options.database),
        options.database.format,
      ),
      profile: options.profile,
    });
    engine = undefined;
    const outcome = await prepareImport({
      database: options.database,
      prepared,
      sources: options.sources,
      profile: options.profile,
      signal: options.signal,
      onProgress: options.onProgress,
    });
    await prepared.close();
    candidateClosed = true;
    prepared = undefined;
    throwIfAborted(options.signal, "db.import.prepare");
    stage = "publication";
    await nativeFiles.publish({ temporary, plan: publication });
    return {
      prepared: outcome.prepared,
      result: {
        ...outcome.result,
        artifacts: [
          {
            kind: "file",
            path: output,
            mediaType: "application/vnd.consultchimps.import-plan",
          },
        ],
      },
    };
  } catch (error) {
    const candidatePrepared = prepared;
    const candidateEngine = engine;
    return failAfterNativeArtifactCleanup({
      operation: "prepare",
      stage,
      kind: "prepared",
      format: "sqlite",
      temporaryPath: temporary,
      storagePaths,
      cause: error,
      close: candidateClosed
        ? []
        : candidatePrepared !== undefined
          ? [() => candidatePrepared.close()]
          : candidateEngine === undefined
            ? []
            : [() => candidateEngine.close()],
      remove: () => removeNativeTemporary(storagePaths),
    });
  }
}

async function openImportBatchUnlocked(options: {
  readonly path: string;
  readonly readonly?: boolean | undefined;
}): Promise<ImportBatch> {
  const input = path.resolve(options.path);
  let engine: NodeSqliteEngine;
  try {
    engine = NodeSqliteEngine.open(input, options.readonly);
  } catch (cause) {
    retainInitializationFailure(input, cause);
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import batch could not be opened. Check that the file exists, that you have access to it, and that it is a saved ConsultChimps .ccplan file. Restore a verified copy or prepare the workbook again if it is damaged.",
      undefined,
      cause,
    );
  }
  try {
    const prepared = await openImportBatchHandle(engine);
    return prepared;
  } catch (error) {
    await closeNativeReadOwner({
      path: input,
      owner: engine,
      failure: { cause: error },
    });
    throw error;
  }
}

export async function openImportBatch(options: {
  readonly path: string;
  readonly readonly?: boolean | undefined;
}): Promise<ImportBatch> {
  const input = path.resolve(options.path);
  return nativeFiles.open(input, () =>
    openImportBatchUnlocked({ path: input, readonly: options.readonly }),
  );
}

export interface ExportNodeDatabaseOptions extends OperationControlOptions {
  readonly database: Database;
  readonly output: string;
  readonly format?: DatabaseFormat | undefined;
  readonly overwrite?: boolean | undefined;
  readonly protectedInputPaths?: readonly string[] | undefined;
}

export interface ExportNodeDatabaseResult extends OperationResult<
  "tablesConverted" | "rowsConverted" | "bytesWritten"
> {
  readonly plan: ConversionPlan;
}

export async function exportDatabase(
  options: ExportNodeDatabaseOptions,
): Promise<ExportNodeDatabaseResult> {
  throwIfAborted(options.signal, "db.export");
  const sourcePath = databasePaths.get(options.database);
  if (sourcePath === undefined) {
    throw databaseError(
      "DB_NATIVE_PATH_UNAVAILABLE",
      "This database was not opened by the Node.js path adapter.",
    );
  }
  const format = options.format ?? options.database.format;
  const plan = await planConversion({ database: options.database, format });
  if (plan.state !== "ready") {
    throw databaseError(
      "DB_CONVERSION_UNSUPPORTED",
      "This conversion has unsupported database objects. Review the conversion plan before exporting.",
      { issues: plan.issues },
    );
  }
  const output = await ensureParentDirectory(options.output);
  const publication = await nativeFiles.planPublication({
    output,
    inputs: [sourcePath, ...(options.protectedInputPaths ?? [])],
    overwrite: options.overwrite,
  });
  const temporary = path.join(
    path.dirname(output),
    `.${path.basename(output)}.cc-export-${globalThis.crypto.randomUUID()}`,
  );
  const storagePaths = nativeTemporaryPaths(temporary, format);
  await assertNativeTemporaryAvailable(storagePaths);
  let stage: "preparation" | "validation" | "publication" = "preparation";
  let targetEngine: DatabaseEngine | undefined;
  let target: Database | undefined;
  let validationEngine: DatabaseEngine | undefined;
  let validation: Database | undefined;
  let tablesConverted = 0;
  let rowsConverted = 0;
  try {
    if (format === options.database.format) {
      const source = engineOf(options.database);
      if (source instanceof NodeSqliteEngine) {
        await source.backupTo(temporary, options.signal);
      } else if (source instanceof NodeDuckDbEngine) {
        await source.copyTo(temporary, options.signal);
      } else {
        throw databaseError(
          "DB_NATIVE_ENGINE_REQUIRED",
          "This database is not using a native Node.js engine.",
        );
      }
    } else {
      targetEngine =
        format === "sqlite"
          ? NodeSqliteEngine.create(temporary)
          : await NodeDuckDbEngine.create(temporary);
      target = await createDatabaseHandle(targetEngine);
      targetEngine = undefined;
      const conversion = await executeConversion({
        source: options.database,
        target,
        plan,
        signal: options.signal,
        onProgress: options.onProgress,
      });
      tablesConverted = conversion.metrics.tablesConverted;
      rowsConverted = conversion.metrics.rowsConverted;
      await target.close();
      target = undefined;
    }
    throwIfAborted(options.signal, "db.export");
    stage = "validation";
    validationEngine = await openEngine(temporary, format, true);
    validation = await openDatabaseHandle(validationEngine);
    validationEngine = undefined;
    await validation.close();
    validation = undefined;
    throwIfAborted(options.signal, "db.export");
    const bytesWritten = Number((await stat(temporary)).size);
    throwIfAborted(options.signal, "db.export");
    stage = "publication";
    await nativeFiles.publish({
      temporary,
      plan: publication,
      signal: options.signal,
    });
    return {
      operation: "db.export",
      artifacts: [
        {
          kind: "file",
          path: output,
          mediaType:
            format === "sqlite"
              ? "application/vnd.sqlite3"
              : "application/vnd.duckdb",
        },
      ],
      warnings: [],
      metrics: { tablesConverted, rowsConverted, bytesWritten },
      plan,
    };
  } catch (error) {
    const candidateTarget = target;
    const candidateTargetEngine = targetEngine;
    const candidateValidation = validation;
    const candidateValidationEngine = validationEngine;
    return failAfterNativeArtifactCleanup({
      operation: "export",
      stage,
      kind: "database",
      format,
      temporaryPath: temporary,
      storagePaths,
      cause: error,
      close: [
        ...(candidateTarget !== undefined
          ? [() => candidateTarget.close()]
          : candidateTargetEngine === undefined
            ? []
            : [() => candidateTargetEngine.close()]),
        ...(candidateValidation !== undefined
          ? [() => candidateValidation.close()]
          : candidateValidationEngine === undefined
            ? []
            : [() => candidateValidationEngine.close()]),
      ],
      remove: () => removeNativeTemporary(storagePaths),
    });
  }
}
