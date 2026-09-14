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
import { assertNoSqliteTableTriggers } from "./internal/database-layout.js";
import { canonicalJson } from "./internal/json.js";
import {
  PREPARED_BINDING_TABLE,
  PREPARED_CAPTURE_TABLE,
  PREPARED_METADATA_TABLE,
  PREPARED_ROW_TABLE,
  PREPARED_STORAGE,
} from "./internal/storage-layouts.js";
import { RetryableClose } from "./internal/retryable-close.js";
import {
  createInternalTables,
  insertInternalRow,
  validateInternalTables,
} from "./internal/storage-schema.js";
import {
  parseImportConflicts,
  parseImportDecisions,
  parseImportProfile,
  validateImportProfile,
} from "./validators.js";
import type {
  ImportConflict,
  ImportDecision,
  ImportProfile,
  ImportBatchId,
  ImportBatchRef,
  ReadyImportBatchRef,
} from "./import/types.js";

export {
  PREPARED_BINDING_TABLE,
  PREPARED_CAPTURE_TABLE,
  PREPARED_METADATA_TABLE,
  PREPARED_ROW_TABLE,
} from "./internal/storage-layouts.js";
export const PREPARED_FORMAT_VERSION = 3;

export async function assertImportBatchWritable(
  transaction: EngineTransaction,
): Promise<void> {
  await assertNoSqliteTableTriggers(
    transaction,
    "sqlite",
    [
      PREPARED_METADATA_TABLE,
      PREPARED_CAPTURE_TABLE,
      PREPARED_BINDING_TABLE,
      PREPARED_ROW_TABLE,
    ],
    {
      code: "DB_SCHEMA_DRIFT",
      message:
        "A trigger was added to a prepared-batch table outside the import operations. Remove the trigger or restore a verified batch copy before writing.",
    },
  );
}

const engines = new WeakMap<ImportBatch, DatabaseEngine>();

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
    `The import batch has an invalid ${field}.`,
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

export interface ImportBatch extends AsyncDisposable {
  readonly id: ImportBatchId;
  readonly databaseId: DatabaseId;
  readonly isOpen: boolean;
  close(): Promise<void>;
}

class ManagedImportBatch implements ImportBatch {
  readonly id: ImportBatchId;
  readonly databaseId: DatabaseId;
  readonly #lifecycle: RetryableClose;

  constructor(
    id: ImportBatchId,
    databaseId: DatabaseId,
    engine: DatabaseEngine,
  ) {
    this.id = id;
    this.databaseId = databaseId;
    engines.set(this, engine);
    this.#lifecycle = new RetryableClose(async () => {
      await engine.close();
      engines.delete(this);
    });
  }

  get isOpen(): boolean {
    return this.#lifecycle.isOpen;
  }

  close(): Promise<void> {
    return this.#lifecycle.close();
  }

  async [Symbol.asyncDispose](): Promise<void> {
    await this.close();
  }
}

export function preparedEngineOf(prepared: ImportBatch): DatabaseEngine {
  assertOpen(prepared.isOpen);
  const engine = engines.get(prepared);
  if (engine === undefined) {
    throw databaseError("DB_PREPARED_CLOSED", "The import batch is closed.");
  }
  return engine;
}

function preparedId(): ImportBatchId {
  return `PLAN-${globalThis.crypto.randomUUID()}` as ImportBatchId;
}

function invalidImportBatch(
  details?: Record<string, unknown>,
  cause?: unknown,
) {
  return databaseError(
    "DB_INVALID_PREPARED_IMPORT",
    "This import batch is incomplete or damaged. Prepare the workbook again or restore a verified batch copy.",
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
    throw invalidImportBatch(undefined, cause);
  }
}

async function validatePreparedSchema(engine: DatabaseEngine): Promise<void> {
  await validateInternalTables({
    query: (sql, values) => preparedQuery(engine, sql, values),
    format: engine.format,
    schema: PREPARED_STORAGE,
    invalid: invalidImportBatch,
  });
}

async function validatePreparedRowOwners(
  engine: Pick<DatabaseEngine, "query">,
): Promise<void> {
  const captureRows = await preparedQuery(
    engine,
    `SELECT capture_id FROM ${PREPARED_CAPTURE_TABLE}`,
  );
  const captureIds = new Set(
    captureRows.map((row) =>
      requiredPreparedString(row["capture_id"], "capture ID"),
    ),
  );
  let cursor: string | undefined;
  for (;;) {
    const rows = await preparedQuery(
      engine,
      cursor === undefined
        ? `SELECT capture_id FROM ${PREPARED_ROW_TABLE} ORDER BY capture_id LIMIT 1`
        : `SELECT capture_id FROM ${PREPARED_ROW_TABLE} WHERE capture_id > ? ORDER BY capture_id LIMIT 1`,
      cursor === undefined ? undefined : [cursor],
    );
    const row = rows[0];
    if (row === undefined) return;
    const captureId = requiredPreparedString(
      row["capture_id"],
      "row capture ID",
    );
    if (!captureIds.has(captureId)) {
      throw databaseError(
        "DB_INVALID_PREPARED_IMPORT",
        "The import batch contains captured rows without their capture definition. Regenerate the batch from its original sources or restore a verified batch copy.",
        { captureId },
      );
    }
    cursor = captureId;
  }
}

export async function createImportBatchHandle(options: {
  readonly engine: DatabaseEngine;
  readonly databaseId: DatabaseId;
  readonly baselineRevision: bigint;
  readonly baselineSchemaFingerprint: string;
  readonly profile: ImportProfile;
}): Promise<ImportBatch> {
  assertBaselineRevision(options.baselineRevision);
  validateImportProfile(options.profile);
  const id = preparedId();
  const formatVersion = BigInt(PREPARED_FORMAT_VERSION);
  const planRevision = 1n;
  const state = "needs-review";
  const recipeJson = canonicalJson(options.profile);
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
    await createInternalTables(transaction, PREPARED_STORAGE, "sqlite");
    await insertInternalRow(transaction, PREPARED_STORAGE.tables.metadata, {
      format_version: formatVersion,
      plan_id: id,
      database_id: options.databaseId,
      baseline_revision: options.baselineRevision,
      schema_fingerprint: options.baselineSchemaFingerprint,
      plan_revision: planRevision,
      state,
      recipe_json: recipeJson,
      conflicts_json: conflictsJson,
      decisions_json: decisionsJson,
      review_fingerprint: fingerprint,
    });
  });
  return new ManagedImportBatch(id, options.databaseId, options.engine);
}

export async function openImportBatchHandle(
  engine: DatabaseEngine,
): Promise<ImportBatch> {
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
      "This file is not a readable ConsultChimps import batch.",
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
      `This import batch uses format version ${String(row["format_version"])}, but this build supports version ${PREPARED_FORMAT_VERSION}. Regenerate the batch from its original sources with this build.`,
      {
        fileVersion: String(row["format_version"]),
        supportedVersion: PREPARED_FORMAT_VERSION,
      },
    );
  }
  await validatePreparedSchema(engine);
  const review = await readStoredPreparedReview(engine);
  await validatePreparedRowOwners(engine);
  const prepared = new ManagedImportBatch(
    review.prepared.id,
    review.prepared.databaseId,
    engine,
  );
  return prepared;
}

export interface PreparedReview {
  readonly prepared: ImportBatchRef | ReadyImportBatchRef;
  readonly profile: ImportProfile;
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
    `The import batch has an invalid ${field}.`,
    { field },
  );
}

function requiredPreparedString(value: unknown, field: string): string {
  if (typeof value === "string" && value.trim().length > 0) return value;
  throw databaseError(
    "DB_INVALID_PREPARED_IMPORT",
    `The import batch has an invalid ${field}.`,
    { field },
  );
}

function nonNegativePreparedBigInt(value: unknown, field: string): bigint {
  const parsed = preparedBigInt(value, field);
  if (parsed >= 0n) return parsed;
  throw databaseError(
    "DB_INVALID_PREPARED_IMPORT",
    `The import batch has an invalid ${field}.`,
    { field },
  );
}

function assertBaselineRevision(value: unknown): asserts value is bigint {
  if (typeof value === "bigint" && value >= 0n) return;
  throw databaseError(
    "DB_INVALID_PREPARED_IMPORT",
    "The import batch has an invalid baseline revision.",
    { field: "baseline revision" },
  );
}

function positivePreparedBigInt(value: unknown, field: string): bigint {
  const parsed = preparedBigInt(value, field);
  if (parsed > 0n) return parsed;
  throw databaseError(
    "DB_INVALID_PREPARED_IMPORT",
    `The import batch has an invalid ${field}.`,
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
      "The import batch metadata is missing.",
    );
  }
  const formatVersion = nonNegativePreparedBigInt(
    row["format_version"],
    "format version",
  );
  if (formatVersion !== BigInt(PREPARED_FORMAT_VERSION)) {
    throw databaseError(
      "DB_UNSUPPORTED_PREPARED_IMPORT_VERSION",
      `This import batch uses format version ${formatVersion.toString()}, but this build supports version ${PREPARED_FORMAT_VERSION}. Regenerate the batch from its original sources with this build.`,
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
    "batch revision",
  );
  if (row["state"] !== "ready" && row["state"] !== "needs-review") {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import batch has an invalid review state.",
      { state: row["state"] },
    );
  }
  const state = row["state"];
  const recipeJson = requiredPreparedString(row["recipe_json"], "profile");
  if (
    typeof row["conflicts_json"] !== "string" ||
    typeof row["decisions_json"] !== "string"
  ) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import batch profile or conflicts are missing.",
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
      "The import batch profile or conflicts are not valid JSON.",
      undefined,
      cause,
    );
  }
  const profile = parseImportProfile(recipeValue);
  const conflicts = parseImportConflicts(conflictsValue);
  const decisions = parseImportDecisions(decisionsValue);
  if (state === "ready" && conflicts.length > 0) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import batch is marked ready but still contains conflicts. Review the batch again before applying it.",
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
      "The import batch review metadata changed outside the review workflow. Regenerate the batch from its original sources or restore a verified batch copy.",
    );
  }
  const common = {
    id: planId as ImportBatchId,
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
    profile,
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
  prepared: ImportBatch,
): Promise<PreparedReview> {
  return readPreparedReviewSnapshot(prepared, (review) =>
    Promise.resolve(review),
  );
}

export async function readPreparedReviewSnapshot<T>(
  prepared: ImportBatch,
  read: (review: PreparedReview, transaction: EngineTransaction) => Promise<T>,
): Promise<T> {
  return preparedEngineOf(prepared).readTransaction(async (transaction) => {
    const review = await readStoredPreparedReview(transaction);
    return read(review, transaction);
  });
}

export async function preparedRef(
  prepared: ImportBatch,
): Promise<ImportBatchRef | ReadyImportBatchRef> {
  return (await readPreparedReview(prepared)).prepared;
}

export async function updatePreparedPlan(options: {
  readonly prepared: ImportBatch;
  readonly profile: ImportProfile;
  readonly conflicts: readonly ImportConflict[];
  readonly ready: boolean;
  readonly decisions?: readonly ImportDecision[] | undefined;
  readonly baselineRevision?: bigint | undefined;
  readonly baselineSchemaFingerprint?: string | undefined;
  readonly expectedReviewFingerprint?: string | undefined;
}): Promise<ImportBatchRef | ReadyImportBatchRef> {
  if (options.baselineRevision !== undefined) {
    assertBaselineRevision(options.baselineRevision);
  }
  validateImportProfile(options.profile);
  if (options.ready && options.conflicts.length > 0) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "An import batch with unresolved conflicts cannot be marked ready.",
    );
  }
  const engine = preparedEngineOf(options.prepared);
  const recipeJson = canonicalJson(options.profile);
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
        "The import batch changed while it was being reviewed. Inspect and approve its latest revision.",
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
    await assertImportBatchWritable(transaction);
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
  prepared: ImportBatch,
  mutate: (transaction: EngineTransaction) => Promise<void>,
): Promise<ImportBatchRef> {
  const engine = preparedEngineOf(prepared);
  return engine.transaction(async (transaction) => {
    const current = await readStoredPreparedReview(transaction);
    await assertImportBatchWritable(transaction);
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

export async function readPreparedRecipe(prepared: ImportBatch): Promise<{
  readonly profile: ImportProfile;
  readonly conflicts: readonly ImportConflict[];
  readonly decisions: readonly ImportDecision[];
}> {
  const review = await readPreparedReview(prepared);
  return {
    profile: review.profile,
    conflicts: review.conflicts,
    decisions: review.decisions,
  };
}
