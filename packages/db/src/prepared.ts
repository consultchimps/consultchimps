import type { DatabaseId } from "./database.js";
import { assertOpen, databaseError } from "./errors.js";
import type { DatabaseEngine } from "./internal/engine.js";
import { canonicalJson } from "./internal/json.js";
import {
  parseImportConflicts,
  parseImportDecisions,
  parseImportRecipe,
  validateImportRecipe,
} from "./validators.js";
import type {
  ImportConflict,
  ImportDecision,
  ImportRecipe,
  PreparedImportId,
  PreparedImportRef,
  ReadyImportRef,
} from "./import/types.js";

export const PREPARED_METADATA_TABLE = "_consultchimps_prepared";
export const PREPARED_CAPTURE_TABLE = "_consultchimps_prepared_captures";
export const PREPARED_BINDING_TABLE = "_consultchimps_prepared_bindings";
export const PREPARED_ROW_TABLE = "_consultchimps_prepared_rows";
export const PREPARED_FORMAT_VERSION = 1;

const engines = new WeakMap<PreparedImport, DatabaseEngine>();

function preparedBigInt(value: unknown, field: string): bigint {
  if (typeof value === "bigint") return value;
  if (typeof value === "number" && Number.isSafeInteger(value)) {
    return BigInt(value);
  }
  if (typeof value === "string" && /^-?\d+$/u.test(value)) {
    return BigInt(value);
  }
  throw databaseError(
    "DB_INVALID_PREPARED_IMPORT",
    `The import plan has an invalid ${field}.`,
    { field },
  );
}

export interface PreparedImport extends AsyncDisposable {
  readonly id: PreparedImportId;
  readonly databaseId: DatabaseId;
  readonly isOpen: boolean;
  close(): Promise<void>;
}

class ManagedPreparedImport implements PreparedImport {
  readonly id: PreparedImportId;
  readonly databaseId: DatabaseId;
  #open = true;

  constructor(id: PreparedImportId, databaseId: DatabaseId) {
    this.id = id;
    this.databaseId = databaseId;
  }

  get isOpen(): boolean {
    return this.#open;
  }

  async close(): Promise<void> {
    if (!this.#open) return;
    const engine = preparedEngineOf(this);
    try {
      await engine.close();
    } finally {
      this.#open = false;
      engines.delete(this);
    }
  }

  async [Symbol.asyncDispose](): Promise<void> {
    await this.close();
  }
}

export function preparedEngineOf(prepared: PreparedImport): DatabaseEngine {
  assertOpen(prepared.isOpen);
  const engine = engines.get(prepared);
  if (engine === undefined) {
    throw databaseError("DB_PREPARED_CLOSED", "The prepared import is closed.");
  }
  return engine;
}

function preparedId(): PreparedImportId {
  return `PLAN-${globalThis.crypto.randomUUID()}` as PreparedImportId;
}

export async function createPreparedImportHandle(options: {
  readonly engine: DatabaseEngine;
  readonly databaseId: DatabaseId;
  readonly baselineRevision: bigint;
  readonly baselineSchemaFingerprint: string;
  readonly recipe: ImportRecipe;
}): Promise<PreparedImport> {
  validateImportRecipe(options.recipe);
  const id = preparedId();
  await options.engine.transaction(async (transaction) => {
    await transaction.execute(
      `CREATE TABLE ${PREPARED_METADATA_TABLE} (format_version BIGINT NOT NULL, plan_id VARCHAR PRIMARY KEY, database_id VARCHAR NOT NULL, baseline_revision BIGINT NOT NULL, schema_fingerprint VARCHAR NOT NULL, plan_revision BIGINT NOT NULL, state VARCHAR NOT NULL, recipe_json VARCHAR NOT NULL, conflicts_json VARCHAR NOT NULL, decisions_json VARCHAR NOT NULL)`,
    );
    await transaction.execute(
      `INSERT INTO ${PREPARED_METADATA_TABLE} VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        BigInt(PREPARED_FORMAT_VERSION),
        id,
        options.databaseId,
        options.baselineRevision,
        options.baselineSchemaFingerprint,
        1n,
        "needs-review",
        canonicalJson(options.recipe),
        "[]",
        "[]",
      ],
    );
    await transaction.execute(
      `CREATE TABLE ${PREPARED_CAPTURE_TABLE} (capture_id VARCHAR PRIMARY KEY, source_file_id VARCHAR, source_key VARCHAR NOT NULL, display_name VARCHAR NOT NULL, selection_key VARCHAR NOT NULL, selection_label VARCHAR NOT NULL, reader_version VARCHAR NOT NULL, content_hash VARCHAR NOT NULL, byte_count BIGINT NOT NULL, reused BIGINT NOT NULL, row_count BIGINT NOT NULL, columns_json VARCHAR NOT NULL)`,
    );
    await transaction.execute(
      `CREATE TABLE ${PREPARED_BINDING_TABLE} (source_key VARCHAR NOT NULL, selection_key VARCHAR NOT NULL, capture_id VARCHAR NOT NULL, display_name VARCHAR NOT NULL, PRIMARY KEY(source_key, selection_key))`,
    );
    await transaction.execute(
      `CREATE TABLE ${PREPARED_ROW_TABLE} (capture_id VARCHAR NOT NULL, source_row BIGINT NOT NULL, values_json VARCHAR NOT NULL, PRIMARY KEY(capture_id, source_row))`,
    );
  });
  const prepared = new ManagedPreparedImport(id, options.databaseId);
  engines.set(prepared, options.engine);
  return prepared;
}

export async function openPreparedImportHandle(
  engine: DatabaseEngine,
): Promise<PreparedImport> {
  let rows;
  try {
    rows = await engine.query(
      `SELECT format_version, plan_id, database_id FROM ${PREPARED_METADATA_TABLE}`,
    );
  } catch (cause) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "This file is not a readable ConsultChimps import plan.",
      undefined,
      cause,
    );
  }
  const row = rows[0];
  if (
    row === undefined ||
    rows.length !== 1 ||
    typeof row["plan_id"] !== "string" ||
    typeof row["database_id"] !== "string"
  ) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "This file is not a readable ConsultChimps import plan.",
    );
  }
  const formatVersion =
    typeof row["format_version"] === "bigint"
      ? row["format_version"]
      : typeof row["format_version"] === "number" &&
          Number.isSafeInteger(row["format_version"])
        ? BigInt(row["format_version"])
        : typeof row["format_version"] === "string" &&
            /^\d+$/u.test(row["format_version"])
          ? BigInt(row["format_version"])
          : null;
  if (formatVersion !== BigInt(PREPARED_FORMAT_VERSION)) {
    throw databaseError(
      "DB_UNSUPPORTED_PREPARED_IMPORT_VERSION",
      `This import plan uses format version ${String(row["format_version"])}, but this build supports version ${PREPARED_FORMAT_VERSION}.`,
      {
        fileVersion: String(row["format_version"]),
        supportedVersion: PREPARED_FORMAT_VERSION,
      },
    );
  }
  const prepared = new ManagedPreparedImport(
    row["plan_id"] as PreparedImportId,
    row["database_id"] as DatabaseId,
  );
  engines.set(prepared, engine);
  return prepared;
}

export async function preparedRef(
  prepared: PreparedImport,
): Promise<PreparedImportRef | ReadyImportRef> {
  const rows = await preparedEngineOf(prepared).query(
    `SELECT plan_id, database_id, baseline_revision, schema_fingerprint, plan_revision, state FROM ${PREPARED_METADATA_TABLE}`,
  );
  const row = rows[0];
  if (
    row === undefined ||
    typeof row["plan_id"] !== "string" ||
    typeof row["database_id"] !== "string"
  ) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import plan metadata is missing.",
    );
  }
  const common = {
    id: row["plan_id"] as PreparedImportId,
    databaseId: row["database_id"] as DatabaseId,
    baselineRevision: preparedBigInt(
      row["baseline_revision"],
      "baseline revision",
    ),
    baselineSchemaFingerprint:
      typeof row["schema_fingerprint"] === "string"
        ? row["schema_fingerprint"]
        : (() => {
            throw databaseError(
              "DB_INVALID_PREPARED_IMPORT",
              "The import plan is missing its database schema fingerprint.",
            );
          })(),
    planRevision: preparedBigInt(row["plan_revision"], "plan revision"),
  };
  if (row["state"] === "ready") return { ...common, state: "ready" };
  if (row["state"] === "needs-review") {
    return { ...common, state: "needs-review" };
  }
  throw databaseError(
    "DB_INVALID_PREPARED_IMPORT",
    "The import plan has an invalid review state.",
    { state: row["state"] },
  );
}

export async function updatePreparedPlan(options: {
  readonly prepared: PreparedImport;
  readonly recipe: ImportRecipe;
  readonly conflicts: readonly ImportConflict[];
  readonly ready: boolean;
  readonly decisions?: readonly ImportDecision[] | undefined;
  readonly baselineRevision?: bigint | undefined;
  readonly baselineSchemaFingerprint?: string | undefined;
}): Promise<PreparedImportRef | ReadyImportRef> {
  validateImportRecipe(options.recipe);
  const engine = preparedEngineOf(options.prepared);
  const recipeJson = canonicalJson(options.recipe);
  const conflictsJson = canonicalJson(options.conflicts);
  const decisionsJson = canonicalJson(options.decisions ?? []);
  const state = options.ready ? "ready" : "needs-review";
  const current = await engine.query(
    `SELECT state, recipe_json, conflicts_json, decisions_json, baseline_revision, schema_fingerprint FROM ${PREPARED_METADATA_TABLE}`,
  );
  const baselineRevision =
    options.baselineRevision ??
    preparedBigInt(current[0]?.["baseline_revision"], "baseline revision");
  const baselineSchemaFingerprint =
    options.baselineSchemaFingerprint ?? current[0]?.["schema_fingerprint"];
  if (typeof baselineSchemaFingerprint !== "string") {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import plan is missing its database schema fingerprint.",
    );
  }
  if (
    current[0]?.["state"] === state &&
    current[0]?.["recipe_json"] === recipeJson &&
    current[0]?.["conflicts_json"] === conflictsJson &&
    current[0]?.["decisions_json"] === decisionsJson &&
    preparedBigInt(current[0]?.["baseline_revision"], "baseline revision") ===
      baselineRevision &&
    current[0]?.["schema_fingerprint"] === baselineSchemaFingerprint
  ) {
    return preparedRef(options.prepared);
  }
  await engine.execute(
    `UPDATE ${PREPARED_METADATA_TABLE} SET plan_revision = plan_revision + 1, baseline_revision = ?, schema_fingerprint = ?, state = ?, recipe_json = ?, conflicts_json = ?, decisions_json = ?`,
    [
      baselineRevision,
      baselineSchemaFingerprint,
      state,
      recipeJson,
      conflictsJson,
      decisionsJson,
    ],
  );
  return preparedRef(options.prepared);
}

export async function readPreparedRecipe(prepared: PreparedImport): Promise<{
  readonly recipe: ImportRecipe;
  readonly conflicts: readonly ImportConflict[];
  readonly decisions: readonly ImportDecision[];
}> {
  const rows = await preparedEngineOf(prepared).query(
    `SELECT recipe_json, conflicts_json, decisions_json FROM ${PREPARED_METADATA_TABLE}`,
  );
  const row = rows[0];
  if (
    row === undefined ||
    typeof row["recipe_json"] !== "string" ||
    typeof row["conflicts_json"] !== "string" ||
    typeof row["decisions_json"] !== "string"
  ) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import plan recipe is missing.",
    );
  }
  let recipeValue: unknown;
  let conflictsValue: unknown;
  let decisionsValue: unknown;
  try {
    recipeValue = JSON.parse(row["recipe_json"]);
    conflictsValue = JSON.parse(row["conflicts_json"]);
    decisionsValue = JSON.parse(row["decisions_json"]);
  } catch (cause) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import plan recipe or conflicts are not valid JSON.",
      undefined,
      cause,
    );
  }
  return {
    recipe: parseImportRecipe(recipeValue),
    conflicts: parseImportConflicts(conflictsValue),
    decisions: parseImportDecisions(decisionsValue),
  };
}
