import { isConsultChimpsError } from "@consultchimps/core";
import { sha256 } from "@noble/hashes/sha2.js";
import { bytesToHex } from "@noble/hashes/utils.js";

import type { DatabaseId } from "./database.js";
import { assertOpen, databaseError } from "./errors.js";
import type {
  DatabaseEngine,
  EngineRow,
  EngineTransaction,
  EngineValue,
} from "./internal/engine.js";
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
export const PREPARED_FORMAT_VERSION = 3;

const REQUIRED_PREPARED_SCHEMA = [
  {
    table: PREPARED_METADATA_TABLE,
    columns: [
      "format_version",
      "plan_id",
      "database_id",
      "baseline_revision",
      "schema_fingerprint",
      "plan_revision",
      "state",
      "recipe_json",
      "conflicts_json",
      "decisions_json",
      "review_fingerprint",
    ],
  },
  {
    table: PREPARED_CAPTURE_TABLE,
    columns: [
      "capture_id",
      "source_file_id",
      "source_key",
      "display_name",
      "selection_key",
      "selection_label",
      "reader_version",
      "content_hash",
      "byte_count",
      "reused",
      "row_count",
      "columns_json",
      "row_checksum",
    ],
  },
  {
    table: PREPARED_BINDING_TABLE,
    columns: ["source_key", "selection_key", "capture_id", "display_name"],
  },
  {
    table: PREPARED_ROW_TABLE,
    columns: ["capture_id", "source_row", "values_json"],
  },
] as const;

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

function reviewFingerprint(options: {
  readonly formatVersion: bigint;
  readonly planId: string;
  readonly databaseId: string;
  readonly baselineRevision: bigint;
  readonly schemaFingerprint: string;
  readonly planRevision: bigint;
  readonly state: string;
  readonly recipeJson: string;
  readonly conflictsJson: string;
  readonly decisionsJson: string;
  readonly captureDefinitionsJson: string;
  readonly bindingsJson: string;
}): string {
  return bytesToHex(
    sha256(
      new TextEncoder().encode(
        JSON.stringify([
          options.formatVersion.toString(),
          options.planId,
          options.databaseId,
          options.baselineRevision.toString(),
          options.schemaFingerprint,
          options.planRevision.toString(),
          options.state,
          options.recipeJson,
          options.conflictsJson,
          options.decisionsJson,
          options.captureDefinitionsJson,
          options.bindingsJson,
        ]),
      ),
    ),
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
  #closing: Promise<void> | undefined;

  constructor(id: PreparedImportId, databaseId: DatabaseId) {
    this.id = id;
    this.databaseId = databaseId;
  }

  get isOpen(): boolean {
    return this.#open;
  }

  close(): Promise<void> {
    if (this.#closing !== undefined) return this.#closing;
    if (!this.#open) return Promise.resolve();
    const closing = Promise.resolve()
      .then(() => preparedEngineOf(this).close())
      .then(() => {
        this.#open = false;
        engines.delete(this);
      })
      .finally(() => {
        if (this.#closing === closing) this.#closing = undefined;
      });
    this.#closing = closing;
    return closing;
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

function invalidPreparedImport(
  details?: Record<string, unknown>,
  cause?: unknown,
) {
  return databaseError(
    "DB_INVALID_PREPARED_IMPORT",
    "This import plan is incomplete or damaged. Prepare the workbook again or restore a verified plan copy.",
    details,
    cause,
  );
}

async function preparedQuery(
  engine: Pick<DatabaseEngine, "query">,
  sql: string,
  values?: readonly EngineValue[],
): Promise<readonly EngineRow[]> {
  try {
    return await engine.query(sql, values);
  } catch (cause) {
    if (isConsultChimpsError(cause)) throw cause;
    if (cause instanceof Error && cause.name === "AbortError") throw cause;
    throw invalidPreparedImport(undefined, cause);
  }
}

async function validatePreparedSchema(engine: DatabaseEngine): Promise<void> {
  const tables = await preparedQuery(
    engine,
    "SELECT name FROM sqlite_master WHERE type = 'table'",
  );
  const tableNames = new Set(
    tables.flatMap((row) =>
      typeof row["name"] === "string" ? [row["name"]] : [],
    ),
  );
  const missingTables = REQUIRED_PREPARED_SCHEMA.flatMap(({ table }) =>
    tableNames.has(table) ? [] : [table],
  );
  if (missingTables.length > 0) {
    throw invalidPreparedImport({ missingTables });
  }
  for (const required of REQUIRED_PREPARED_SCHEMA) {
    const columns = await preparedQuery(
      engine,
      "SELECT name FROM pragma_table_info(?) ORDER BY cid",
      [required.table],
    );
    const actualColumns = columns.map((row) =>
      typeof row["name"] === "string" ? row["name"] : null,
    );
    if (
      actualColumns.length !== required.columns.length ||
      actualColumns.some((column, index) => column !== required.columns[index])
    ) {
      throw invalidPreparedImport({
        table: required.table,
        expectedColumns: [...required.columns],
        actualColumns,
      });
    }
  }
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
  const formatVersion = BigInt(PREPARED_FORMAT_VERSION);
  const planRevision = 1n;
  const state = "needs-review";
  const recipeJson = canonicalJson(options.recipe);
  const conflictsJson = "[]";
  const decisionsJson = "[]";
  const fingerprint = reviewFingerprint({
    formatVersion,
    planId: id,
    databaseId: options.databaseId,
    baselineRevision: options.baselineRevision,
    schemaFingerprint: options.baselineSchemaFingerprint,
    planRevision,
    state,
    recipeJson,
    conflictsJson,
    decisionsJson,
    captureDefinitionsJson: "[]",
    bindingsJson: "[]",
  });
  await options.engine.transaction(async (transaction) => {
    await transaction.execute(
      `CREATE TABLE ${PREPARED_METADATA_TABLE} (format_version BIGINT NOT NULL, plan_id VARCHAR PRIMARY KEY, database_id VARCHAR NOT NULL, baseline_revision BIGINT NOT NULL, schema_fingerprint VARCHAR NOT NULL, plan_revision BIGINT NOT NULL, state VARCHAR NOT NULL, recipe_json VARCHAR NOT NULL, conflicts_json VARCHAR NOT NULL, decisions_json VARCHAR NOT NULL, review_fingerprint VARCHAR NOT NULL)`,
    );
    await transaction.execute(
      `INSERT INTO ${PREPARED_METADATA_TABLE} VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        formatVersion,
        id,
        options.databaseId,
        options.baselineRevision,
        options.baselineSchemaFingerprint,
        planRevision,
        state,
        recipeJson,
        conflictsJson,
        decisionsJson,
        fingerprint,
      ],
    );
    await transaction.execute(
      `CREATE TABLE ${PREPARED_CAPTURE_TABLE} (capture_id VARCHAR PRIMARY KEY, source_file_id VARCHAR, source_key VARCHAR NOT NULL, display_name VARCHAR NOT NULL, selection_key VARCHAR NOT NULL, selection_label VARCHAR NOT NULL, reader_version VARCHAR NOT NULL, content_hash VARCHAR NOT NULL, byte_count BIGINT NOT NULL, reused BIGINT NOT NULL, row_count BIGINT NOT NULL, columns_json VARCHAR NOT NULL, row_checksum VARCHAR NOT NULL)`,
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
  const rows = await preparedQuery(
    engine,
    `SELECT format_version, plan_id, database_id FROM ${PREPARED_METADATA_TABLE}`,
  );
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
      `This import plan uses format version ${String(row["format_version"])}, but this build supports version ${PREPARED_FORMAT_VERSION}. Regenerate the plan from its original sources with this build.`,
      {
        fileVersion: String(row["format_version"]),
        supportedVersion: PREPARED_FORMAT_VERSION,
      },
    );
  }
  await validatePreparedSchema(engine);
  const review = await readStoredPreparedReview(engine);
  const prepared = new ManagedPreparedImport(
    review.prepared.id,
    review.prepared.databaseId,
  );
  engines.set(prepared, engine);
  return prepared;
}

export interface PreparedReview {
  readonly prepared: PreparedImportRef | ReadyImportRef;
  readonly recipe: ImportRecipe;
  readonly conflicts: readonly ImportConflict[];
  readonly decisions: readonly ImportDecision[];
}

interface StoredPreparedReview extends PreparedReview {
  readonly raw: {
    readonly state: "needs-review" | "ready";
    readonly recipeJson: string;
    readonly conflictsJson: string;
    readonly decisionsJson: string;
    readonly captureDefinitionsJson: string;
    readonly bindingsJson: string;
  };
}

function preparedString(value: unknown, field: string): string {
  if (typeof value === "string") return value;
  throw databaseError(
    "DB_INVALID_PREPARED_IMPORT",
    `The import plan has an invalid ${field}.`,
    { field },
  );
}

function requiredPreparedString(value: unknown, field: string): string {
  if (typeof value === "string" && value.trim().length > 0) return value;
  throw databaseError(
    "DB_INVALID_PREPARED_IMPORT",
    `The import plan has an invalid ${field}.`,
    { field },
  );
}

function nonNegativePreparedBigInt(value: unknown, field: string): bigint {
  const parsed = preparedBigInt(value, field);
  if (parsed >= 0n) return parsed;
  throw databaseError(
    "DB_INVALID_PREPARED_IMPORT",
    `The import plan has an invalid ${field}.`,
    { field },
  );
}

function positivePreparedBigInt(value: unknown, field: string): bigint {
  const parsed = preparedBigInt(value, field);
  if (parsed > 0n) return parsed;
  throw databaseError(
    "DB_INVALID_PREPARED_IMPORT",
    `The import plan has an invalid ${field}.`,
    { field },
  );
}

function compareStoredTuple(
  left: readonly (string | null)[],
  right: readonly (string | null)[],
): number {
  const leftKey = JSON.stringify(left);
  const rightKey = JSON.stringify(right);
  return leftKey < rightKey ? -1 : leftKey > rightKey ? 1 : 0;
}

async function readPreparedCaptureMetadata(
  engine: Pick<DatabaseEngine, "query">,
): Promise<{
  readonly captureDefinitionsJson: string;
  readonly bindingsJson: string;
}> {
  const captureRows = await preparedQuery(
    engine,
    `SELECT capture_id, source_file_id, source_key, display_name, selection_key, selection_label, reader_version, content_hash, byte_count, reused, row_count, columns_json, row_checksum FROM ${PREPARED_CAPTURE_TABLE}`,
  );
  const captures = captureRows.map(
    (row) =>
      [
        requiredPreparedString(row["capture_id"], "capture ID"),
        row["source_file_id"] === null
          ? null
          : requiredPreparedString(row["source_file_id"], "source file ID"),
        requiredPreparedString(row["source_key"], "source key"),
        preparedString(row["display_name"], "source display name"),
        requiredPreparedString(row["selection_key"], "selection key"),
        preparedString(row["selection_label"], "selection label"),
        requiredPreparedString(row["reader_version"], "reader version"),
        requiredPreparedString(row["content_hash"], "content hash"),
        nonNegativePreparedBigInt(row["byte_count"], "byte count").toString(),
        nonNegativePreparedBigInt(row["reused"], "reuse marker").toString(),
        nonNegativePreparedBigInt(row["row_count"], "row count").toString(),
        preparedString(row["columns_json"], "captured columns"),
        requiredPreparedString(row["row_checksum"], "captured row checksum"),
      ] satisfies readonly (string | null)[],
  );
  const bindingRows = await preparedQuery(
    engine,
    `SELECT source_key, selection_key, capture_id, display_name FROM ${PREPARED_BINDING_TABLE}`,
  );
  const bindings = bindingRows.map(
    (row) =>
      [
        requiredPreparedString(row["source_key"], "source key"),
        requiredPreparedString(row["selection_key"], "selection key"),
        requiredPreparedString(row["capture_id"], "capture ID"),
        preparedString(row["display_name"], "source display name"),
      ] satisfies readonly (string | null)[],
  );
  captures.sort(compareStoredTuple);
  bindings.sort(compareStoredTuple);
  return {
    captureDefinitionsJson: JSON.stringify(captures),
    bindingsJson: JSON.stringify(bindings),
  };
}

async function readStoredPreparedReview(
  engine: Pick<DatabaseEngine, "query">,
): Promise<StoredPreparedReview> {
  const rows = await preparedQuery(
    engine,
    `SELECT format_version, plan_id, database_id, baseline_revision, schema_fingerprint, plan_revision, state, recipe_json, conflicts_json, decisions_json, review_fingerprint FROM ${PREPARED_METADATA_TABLE}`,
  );
  const row = rows[0];
  if (row === undefined || rows.length !== 1) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import plan metadata is missing.",
    );
  }
  const formatVersion = nonNegativePreparedBigInt(
    row["format_version"],
    "format version",
  );
  if (formatVersion !== BigInt(PREPARED_FORMAT_VERSION)) {
    throw databaseError(
      "DB_UNSUPPORTED_PREPARED_IMPORT_VERSION",
      `This import plan uses format version ${formatVersion.toString()}, but this build supports version ${PREPARED_FORMAT_VERSION}. Regenerate the plan from its original sources with this build.`,
      {
        fileVersion: formatVersion.toString(),
        supportedVersion: PREPARED_FORMAT_VERSION,
      },
    );
  }
  const planId = requiredPreparedString(row["plan_id"], "plan ID");
  const databaseId = requiredPreparedString(row["database_id"], "database ID");
  const baselineRevision = nonNegativePreparedBigInt(
    row["baseline_revision"],
    "baseline revision",
  );
  const schemaFingerprint = requiredPreparedString(
    row["schema_fingerprint"],
    "database schema fingerprint",
  );
  const planRevision = positivePreparedBigInt(
    row["plan_revision"],
    "plan revision",
  );
  if (row["state"] !== "ready" && row["state"] !== "needs-review") {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import plan has an invalid review state.",
      { state: row["state"] },
    );
  }
  const state = row["state"];
  const recipeJson = requiredPreparedString(row["recipe_json"], "recipe");
  if (
    typeof row["conflicts_json"] !== "string" ||
    typeof row["decisions_json"] !== "string"
  ) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import plan recipe or conflicts are missing.",
    );
  }
  const conflictsJson = row["conflicts_json"];
  const decisionsJson = row["decisions_json"];
  let recipeValue: unknown;
  let conflictsValue: unknown;
  let decisionsValue: unknown;
  try {
    recipeValue = JSON.parse(recipeJson);
    conflictsValue = JSON.parse(conflictsJson);
    decisionsValue = JSON.parse(decisionsJson);
  } catch (cause) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import plan recipe or conflicts are not valid JSON.",
      undefined,
      cause,
    );
  }
  const recipe = parseImportRecipe(recipeValue);
  const conflicts = parseImportConflicts(conflictsValue);
  const decisions = parseImportDecisions(decisionsValue);
  if (state === "ready" && conflicts.length > 0) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import plan is marked ready but still contains conflicts. Review the plan again before applying it.",
    );
  }
  const storedFingerprint = requiredPreparedString(
    row["review_fingerprint"],
    "review fingerprint",
  );
  const { captureDefinitionsJson, bindingsJson } =
    await readPreparedCaptureMetadata(engine);
  const expectedFingerprint = reviewFingerprint({
    formatVersion,
    planId,
    databaseId,
    baselineRevision,
    schemaFingerprint,
    planRevision,
    state,
    recipeJson,
    conflictsJson,
    decisionsJson,
    captureDefinitionsJson,
    bindingsJson,
  });
  if (
    !/^[0-9a-f]{64}$/u.test(storedFingerprint) ||
    storedFingerprint !== expectedFingerprint
  ) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import plan review metadata changed outside the review workflow. Regenerate the plan from its original sources or restore a verified plan copy.",
    );
  }
  const common = {
    id: planId as PreparedImportId,
    databaseId: databaseId as DatabaseId,
    baselineRevision,
    baselineSchemaFingerprint: schemaFingerprint,
    planRevision,
    reviewFingerprint: storedFingerprint,
  };
  return {
    prepared:
      state === "ready"
        ? { ...common, state: "ready" }
        : { ...common, state: "needs-review" },
    recipe,
    conflicts,
    decisions,
    raw: {
      state,
      recipeJson,
      conflictsJson,
      decisionsJson,
      captureDefinitionsJson,
      bindingsJson,
    },
  };
}

export async function readPreparedReview(
  prepared: PreparedImport,
): Promise<PreparedReview> {
  return readPreparedReviewSnapshot(prepared, (review) =>
    Promise.resolve(review),
  );
}

export async function readPreparedReviewSnapshot<T>(
  prepared: PreparedImport,
  read: (review: PreparedReview, transaction: EngineTransaction) => Promise<T>,
): Promise<T> {
  return preparedEngineOf(prepared).readTransaction(async (transaction) => {
    const review = await readStoredPreparedReview(transaction);
    return read(review, transaction);
  });
}

export async function preparedRef(
  prepared: PreparedImport,
): Promise<PreparedImportRef | ReadyImportRef> {
  return (await readPreparedReview(prepared)).prepared;
}

export async function updatePreparedPlan(options: {
  readonly prepared: PreparedImport;
  readonly recipe: ImportRecipe;
  readonly conflicts: readonly ImportConflict[];
  readonly ready: boolean;
  readonly decisions?: readonly ImportDecision[] | undefined;
  readonly baselineRevision?: bigint | undefined;
  readonly baselineSchemaFingerprint?: string | undefined;
  readonly expectedReviewFingerprint?: string | undefined;
}): Promise<PreparedImportRef | ReadyImportRef> {
  validateImportRecipe(options.recipe);
  if (options.ready && options.conflicts.length > 0) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "An import plan with unresolved conflicts cannot be marked ready.",
    );
  }
  const engine = preparedEngineOf(options.prepared);
  const recipeJson = canonicalJson(options.recipe);
  const conflictsJson = canonicalJson(options.conflicts);
  const decisionsJson = canonicalJson(options.decisions ?? []);
  const state = options.ready ? "ready" : "needs-review";
  return engine.transaction(async (transaction) => {
    const current = await readStoredPreparedReview(transaction);
    if (
      options.expectedReviewFingerprint !== undefined &&
      current.prepared.reviewFingerprint !== options.expectedReviewFingerprint
    ) {
      throw databaseError(
        "DB_STALE_IMPORT_PLAN",
        "The import plan changed while it was being reviewed. Inspect and approve its latest revision.",
      );
    }
    const baselineRevision =
      options.baselineRevision ?? current.prepared.baselineRevision;
    const baselineSchemaFingerprint =
      options.baselineSchemaFingerprint ??
      current.prepared.baselineSchemaFingerprint;
    if (
      current.raw.state === state &&
      current.raw.recipeJson === recipeJson &&
      current.raw.conflictsJson === conflictsJson &&
      current.raw.decisionsJson === decisionsJson &&
      current.prepared.baselineRevision === baselineRevision &&
      current.prepared.baselineSchemaFingerprint === baselineSchemaFingerprint
    ) {
      return current.prepared;
    }
    const planRevision = current.prepared.planRevision + 1n;
    const fingerprint = reviewFingerprint({
      formatVersion: BigInt(PREPARED_FORMAT_VERSION),
      planId: current.prepared.id,
      databaseId: current.prepared.databaseId,
      baselineRevision,
      schemaFingerprint: baselineSchemaFingerprint,
      planRevision,
      state,
      recipeJson,
      conflictsJson,
      decisionsJson,
      captureDefinitionsJson: current.raw.captureDefinitionsJson,
      bindingsJson: current.raw.bindingsJson,
    });
    await transaction.execute(
      `UPDATE ${PREPARED_METADATA_TABLE} SET plan_revision = ?, baseline_revision = ?, schema_fingerprint = ?, state = ?, recipe_json = ?, conflicts_json = ?, decisions_json = ?, review_fingerprint = ?`,
      [
        planRevision,
        baselineRevision,
        baselineSchemaFingerprint,
        state,
        recipeJson,
        conflictsJson,
        decisionsJson,
        fingerprint,
      ],
    );
    const common = {
      id: current.prepared.id,
      databaseId: current.prepared.databaseId,
      planRevision,
      baselineRevision,
      baselineSchemaFingerprint,
      reviewFingerprint: fingerprint,
    };
    return state === "ready"
      ? { ...common, state: "ready" }
      : { ...common, state: "needs-review" };
  });
}

export async function updatePreparedCaptureMetadata(
  prepared: PreparedImport,
  mutate: (transaction: EngineTransaction) => Promise<void>,
): Promise<PreparedImportRef> {
  const engine = preparedEngineOf(prepared);
  return engine.transaction(async (transaction) => {
    const current = await readStoredPreparedReview(transaction);
    await mutate(transaction);
    const { captureDefinitionsJson, bindingsJson } =
      await readPreparedCaptureMetadata(transaction);
    const planRevision =
      current.prepared.state === "ready"
        ? current.prepared.planRevision + 1n
        : current.prepared.planRevision;
    const state = "needs-review";
    const fingerprint = reviewFingerprint({
      formatVersion: BigInt(PREPARED_FORMAT_VERSION),
      planId: current.prepared.id,
      databaseId: current.prepared.databaseId,
      baselineRevision: current.prepared.baselineRevision,
      schemaFingerprint: current.prepared.baselineSchemaFingerprint,
      planRevision,
      state,
      recipeJson: current.raw.recipeJson,
      conflictsJson: current.raw.conflictsJson,
      decisionsJson: current.raw.decisionsJson,
      captureDefinitionsJson,
      bindingsJson,
    });
    await transaction.execute(
      `UPDATE ${PREPARED_METADATA_TABLE} SET plan_revision = ?, state = ?, review_fingerprint = ?`,
      [planRevision, state, fingerprint],
    );
    return {
      id: current.prepared.id,
      databaseId: current.prepared.databaseId,
      baselineRevision: current.prepared.baselineRevision,
      baselineSchemaFingerprint: current.prepared.baselineSchemaFingerprint,
      planRevision,
      reviewFingerprint: fingerprint,
      state,
    };
  });
}

export async function readPreparedRecipe(prepared: PreparedImport): Promise<{
  readonly recipe: ImportRecipe;
  readonly conflicts: readonly ImportConflict[];
  readonly decisions: readonly ImportDecision[];
}> {
  const review = await readPreparedReview(prepared);
  return {
    recipe: review.recipe,
    conflicts: review.conflicts,
    decisions: review.decisions,
  };
}
