import { throwIfAborted } from "@consultchimps/core";

import {
  engineOf,
  parseStoredTableSchema,
  valueAsBigInt,
  valueAsPositiveBigInt,
  valueAsString,
} from "../database.js";
import { databaseError } from "../errors.js";
import {
  assertMetadataAllocationCounters,
  assertRecordAllocationAvailable,
  assertRowAllocationCounters,
} from "../internal/allocation-counters.js";
import type { EngineTransaction, EngineValue } from "../internal/engine.js";
import { assertNoManagedDatabaseTriggers } from "../internal/database-layout.js";
import { canonicalJson } from "../internal/json.js";
import { DATABASE_STORAGE } from "../internal/storage-layouts.js";
import { insertInternalRow } from "../internal/storage-schema.js";
import {
  CAPTURE_TABLE,
  COUNTERS_TABLE,
  DATABASE_METADATA_TABLE,
  DELIVERY_TABLE,
  IMPORT_REQUEST_TABLE,
  PLAN_TABLE,
  SOURCE_CONTENT_TABLE,
  SOURCE_FILE_TABLE,
  CAPTURE_ROW_TABLE,
  TABLE_REGISTRY_TABLE,
} from "../metadata.js";
import {
  PREPARED_ROW_TABLE,
  preparedEngineOf,
  readPreparedReviewSnapshot,
} from "../prepared.js";
import {
  createManagedTable,
  readSchemaFingerprint,
  sortTablesByReferences,
} from "../records.js";
import { formatRecordId, identifierKey, type TableSchema } from "../schema.js";
import { parseBatchContext } from "../validators.js";
import { parseStoredBatchContext } from "../internal/stored-delivery.js";
import { parseImportCellsJson, valueForColumn } from "./inference.js";
import {
  effectiveApplicationKey,
  inspectApplicationIdentity,
} from "./application-identity.js";
import {
  orderRoutesByReferences,
  preparedCapturesFromEngine,
  routeColumns,
  routeKey,
} from "./planning.js";
import { assertCaptureRowCount, readSourceRowPage } from "./source-rows.js";
import { createCaptureRowChecksum } from "./row-checksum.js";
import {
  assertReceiptDeliveryMemberships,
  receiptIds,
  validatedReceiptRowCount,
} from "./receipt.js";
import type { ApplyImportOptions, ImportResult } from "./types.js";
import type { DatabaseWriteResult } from "../write-completion.js";

const CAPTURE_BATCH_ROWS = 2_000;

async function allocate(
  transaction: EngineTransaction,
  counter: string,
  prefix: string,
  padding = 6,
): Promise<string> {
  const rows = await transaction.query(
    `SELECT next_value FROM ${COUNTERS_TABLE} WHERE counter_name = ?`,
    [counter],
  );
  const value = valueAsPositiveBigInt(
    rows[0]?.["next_value"],
    `${counter} counter`,
  );
  await transaction.execute(
    `UPDATE ${COUNTERS_TABLE} SET next_value = ? WHERE counter_name = ?`,
    [value + 1n, counter],
  );
  return `${prefix}-${value.toString().padStart(padding, "0")}`;
}

async function registeredSchema(
  transaction: EngineTransaction,
  table: string,
): Promise<TableSchema | null> {
  const rows = await transaction.query(
    `SELECT schema_json FROM ${TABLE_REGISTRY_TABLE} WHERE table_name = ?`,
    [table],
  );
  const stored = rows[0]?.["schema_json"];
  return typeof stored === "string"
    ? parseStoredTableSchema(stored, table)
    : null;
}

function assertCaptureRowChecksum(options: {
  readonly owner: "prepared-import" | "database";
  readonly captureId: string;
  readonly expected: string;
  readonly actual: string;
}): void {
  if (options.actual === options.expected) return;
  throw databaseError(
    options.owner === "prepared-import"
      ? "DB_INVALID_PREPARED_IMPORT"
      : "DB_CORRUPT_DATABASE",
    options.owner === "prepared-import"
      ? "The import batch's captured rows do not match the reviewed row checksum. Regenerate the batch from its original sources or restore a verified batch copy."
      : "A saved capture's rows do not match the reviewed row checksum. Restore a verified database copy or regenerate the import batch from its original sources.",
    { captureId: options.captureId },
  );
}

export async function applyImport(
  options: ApplyImportOptions,
): Promise<ImportResult> {
  throwIfAborted(options.signal, "db.import.apply");
  if (options.requestId.trim().length === 0) {
    throw databaseError(
      "DB_IMPORT_REQUEST_ID_REQUIRED",
      "Give the import a request ID so retrying it cannot apply the same batch twice.",
    );
  }
  const delivery =
    options.batchContext === undefined
      ? undefined
      : parseBatchContext(options.batchContext);
  if (options.approved.state !== "ready") {
    throw databaseError(
      "DB_IMPORT_NEEDS_REVIEW",
      "The import batch still needs review.",
    );
  }
  const preparedSnapshot = await readPreparedReviewSnapshot(
    options.prepared,
    async (review, transaction) => ({
      review,
      captures: await preparedCapturesFromEngine(transaction),
    }),
  );
  const preparedPlan = preparedSnapshot.review;
  const actual = preparedPlan.prepared;
  if (
    actual.state !== "ready" ||
    actual.id !== options.approved.id ||
    actual.databaseId !== options.approved.databaseId ||
    actual.planRevision !== options.approved.planRevision ||
    actual.baselineRevision !== options.approved.baselineRevision ||
    actual.baselineSchemaFingerprint !==
      options.approved.baselineSchemaFingerprint ||
    actual.reviewFingerprint !== options.approved.reviewFingerprint
  ) {
    throw databaseError(
      "DB_STALE_IMPORT_PLAN",
      "The approved import batch is no longer current. Inspect and approve its latest revision.",
    );
  }
  const approved = actual;
  if (approved.databaseId !== options.database.id) {
    throw databaseError(
      "DB_STALE_IMPORT_PLAN",
      "The approved import batch belongs to another database. Prepare it again for this database before applying.",
    );
  }
  const target = engineOf(options.database);
  const preparedEngine = preparedEngineOf(options.prepared);
  let profile = preparedPlan.profile;
  const { conflicts, decisions } = preparedPlan;
  const captures = preparedSnapshot.captures;
  let rowsImported = 0;
  let rowsReused = 0;
  let tablesCreated = 0;
  let batchId: string | undefined;
  let batchesRecorded = 0;
  let databaseWrite: DatabaseWriteResult["databaseWrite"] = "unchanged";
  const importIds: string[] = [];
  const appliedCaptureIds = new Set<string>();
  let captureIds: string[] = [];
  await target.transaction(async (transaction) => {
    const existingRequest = await transaction.query(
      `SELECT plan_id, plan_revision, import_ids_json, capture_ids_json, row_count FROM ${IMPORT_REQUEST_TABLE} WHERE request_id = ?`,
      [options.requestId],
    );
    if (existingRequest[0] !== undefined) {
      const samePlan =
        valueAsString(existingRequest[0]["plan_id"], "plan ID") ===
          approved.id &&
        valueAsBigInt(existingRequest[0]["plan_revision"], "plan revision") ===
          approved.planRevision;
      if (!samePlan) {
        throw databaseError(
          "DB_REQUEST_ID_CONFLICT",
          "This request ID was already used for a different import. Choose a new request ID.",
          { requestId: options.requestId },
        );
      }
      const savedImportIds = receiptIds(existingRequest[0]["import_ids_json"]);
      const savedCaptureIds = receiptIds(
        existingRequest[0]["capture_ids_json"],
      );
      const receiptRowCount = await validatedReceiptRowCount({
        transaction,
        importIds: savedImportIds,
        captureIds: savedCaptureIds,
        storedRowCount: existingRequest[0]["row_count"],
        profile,
        planId: approved.id,
        planRevision: approved.planRevision,
      });
      importIds.push(...savedImportIds);
      for (const captureId of savedCaptureIds) appliedCaptureIds.add(captureId);
      rowsReused = Number(receiptRowCount);
      const deliveries = await transaction.query(
        `SELECT delivery_id, context_json FROM ${DELIVERY_TABLE} WHERE request_id = ?`,
        [options.requestId],
      );
      if (deliveries[0] !== undefined) {
        batchId = valueAsString(deliveries[0]["delivery_id"], "batch ID");
        const storedDelivery = parseStoredBatchContext(
          deliveries[0]["context_json"],
        );
        await assertReceiptDeliveryMemberships({
          transaction,
          batchId,
          captureIds: savedCaptureIds,
        });
        if (
          delivery === undefined ||
          canonicalJson(storedDelivery) !== canonicalJson(delivery)
        ) {
          throw databaseError(
            "DB_REQUEST_ID_CONFLICT",
            "This request ID was already used with different batch details. Choose a new request ID.",
            { requestId: options.requestId },
          );
        }
      } else if (delivery !== undefined) {
        throw databaseError(
          "DB_REQUEST_ID_CONFLICT",
          "This request ID was already used without batch details. Choose a new request ID.",
          { requestId: options.requestId },
        );
      }
      captureIds = [...appliedCaptureIds].sort();
      return;
    }
    const standaloneDelivery = await transaction.query(
      `SELECT request_id FROM ${DELIVERY_TABLE} WHERE request_id = ? LIMIT 1`,
      [options.requestId],
    );
    if (standaloneDelivery.length > 0) {
      throw databaseError(
        "DB_REQUEST_ID_CONFLICT",
        "This request ID was already used for a batch. Choose a new request ID for the import.",
        { requestId: options.requestId },
      );
    }
    await assertNoManagedDatabaseTriggers(transaction, options.database.format);
    const revisionRows = await transaction.query(
      `SELECT revision FROM ${DATABASE_METADATA_TABLE}`,
    );
    const revision = valueAsBigInt(
      revisionRows[0]?.["revision"],
      "database revision",
    );
    const schemaFingerprint = await readSchemaFingerprint(
      transaction,
      options.database.format,
    );
    if (
      revision !== approved.baselineRevision ||
      schemaFingerprint !== approved.baselineSchemaFingerprint
    ) {
      throw databaseError(
        "DB_STALE_IMPORT_PLAN",
        "The database changed after this import was prepared. Prepare it again before applying.",
      );
    }
    await assertMetadataAllocationCounters(transaction);
    const registeredRows = await transaction.query(
      `SELECT table_name, schema_json FROM ${TABLE_REGISTRY_TABLE}`,
    );
    const canonicalTableNames = new Map(
      registeredRows.map((row) => {
        const name = valueAsString(row["table_name"], "table name");
        return [identifierKey(name), name] as const;
      }),
    );
    for (const route of profile.routes) {
      if (route.destination.kind !== "new-table") continue;
      const name = route.destination.schema.name;
      if (!canonicalTableNames.has(identifierKey(name))) {
        canonicalTableNames.set(identifierKey(name), name);
      }
    }
    profile = {
      ...profile,
      routes: profile.routes.map((route) => {
        if (route.destination.kind === "existing-table") {
          return {
            ...route,
            destination: {
              kind: "existing-table" as const,
              table:
                canonicalTableNames.get(
                  identifierKey(route.destination.table),
                ) ?? route.destination.table,
            },
          };
        }
        if (route.destination.kind === "new-table") {
          const name =
            canonicalTableNames.get(
              identifierKey(route.destination.schema.name),
            ) ?? route.destination.schema.name;
          return {
            ...route,
            destination: {
              ...route.destination,
              schema: { ...route.destination.schema, name },
            },
          };
        }
        return route;
      }),
    };
    const captureIdByBinding = new Map<string, string>();
    const verifiedCaptureIds = new Set<string>();
    const publishedCaptures = new Map<
      string,
      { readonly targetCaptureId: string; readonly sourceFileId: string }
    >();
    for (const capture of captures) {
      const published = publishedCaptures.get(capture.captureId);
      if (published !== undefined) {
        await insertInternalRow(
          transaction,
          DATABASE_STORAGE.tables.sourceNames,
          {
            source_file_id: published.sourceFileId,
            display_name: capture.displayName,
          },
          { onConflict: "do-nothing" },
        );
        captureIdByBinding.set(
          routeKey(capture.sourceKey, capture.selectionKey),
          published.targetCaptureId,
        );
        continue;
      }
      let targetCaptureId =
        capture.sourceFileId === null ? null : capture.captureId;
      let sourceFileId = capture.sourceFileId;
      if (!capture.reused) {
        const contentRows = await transaction.query(
          `SELECT content_hash FROM ${SOURCE_CONTENT_TABLE} WHERE content_hash = ?`,
          [capture.contentHash],
        );
        if (contentRows.length === 0) {
          await insertInternalRow(
            transaction,
            DATABASE_STORAGE.tables.sourceContents,
            {
              content_hash: capture.contentHash,
              byte_count: capture.byteCount,
            },
          );
        }
        const fileRows = await transaction.query(
          `SELECT source_file_id FROM ${SOURCE_FILE_TABLE} WHERE content_hash = ?`,
          [capture.contentHash],
        );
        sourceFileId =
          fileRows[0] === undefined
            ? await allocate(transaction, "source_file", "SRC")
            : valueAsString(fileRows[0]["source_file_id"], "source file ID");
        if (fileRows.length === 0) {
          await insertInternalRow(
            transaction,
            DATABASE_STORAGE.tables.sourceFiles,
            {
              source_file_id: sourceFileId,
              content_hash: capture.contentHash,
              display_name: capture.displayName,
            },
          );
        }
        const existingCapture = await transaction.query(
          `SELECT capture_id FROM ${CAPTURE_TABLE} WHERE source_file_id = ? AND selection_key = ? AND reader_version = ?`,
          [sourceFileId, capture.selectionKey, capture.readerVersion],
        );
        targetCaptureId =
          existingCapture[0] === undefined
            ? await allocate(transaction, "capture", "CAP")
            : valueAsString(existingCapture[0]["capture_id"], "capture ID");
        if (existingCapture.length === 0) {
          await insertInternalRow(
            transaction,
            DATABASE_STORAGE.tables.captures,
            {
              capture_id: targetCaptureId,
              source_file_id: sourceFileId,
              source_key: capture.sourceKey,
              selection_key: capture.selectionKey,
              selection_label: capture.selectionLabel,
              reader_version: capture.readerVersion,
              state: "completed",
              row_count: capture.rowCount,
              columns_json: JSON.stringify(capture.columns),
            },
          );
          let sourceRowCursor: bigint | undefined;
          let copiedRows = 0n;
          const rowChecksum = createCaptureRowChecksum();
          while (true) {
            const page = await readSourceRowPage({
              engine: preparedEngine,
              table: PREPARED_ROW_TABLE,
              captureId: capture.captureId,
              cursor: sourceRowCursor,
              limit: CAPTURE_BATCH_ROWS,
              owner: "prepared-import",
            });
            throwIfAborted(options.signal, "db.import.apply");
            if (page.rows.length === 0) break;
            const rows = page.rows.map(({ row, sourceRow }) => {
              const valuesJson = valueAsString(
                row["values_json"],
                "captured values",
              );
              rowChecksum.update(sourceRow, valuesJson);
              return [targetCaptureId, sourceRow, valuesJson] as const;
            });
            await transaction.bulkInsert({
              table: CAPTURE_ROW_TABLE,
              columns: ["capture_id", "source_row", "values_json"],
              rows,
              signal: options.signal,
            });
            copiedRows += BigInt(rows.length);
            sourceRowCursor = page.cursor;
          }
          assertCaptureRowCount({
            owner: "prepared-import",
            captureId: capture.captureId,
            expected: capture.rowCount,
            actual: copiedRows,
          });
          assertCaptureRowChecksum({
            owner: "prepared-import",
            captureId: capture.captureId,
            expected: capture.rowChecksum,
            actual: rowChecksum.digest(),
          });
          verifiedCaptureIds.add(targetCaptureId);
        }
      }
      if (targetCaptureId === null || sourceFileId === null) {
        throw databaseError(
          "DB_CORRUPT_PREPARED_IMPORT",
          "A prepared capture is missing its target identity.",
        );
      }
      await insertInternalRow(
        transaction,
        DATABASE_STORAGE.tables.sourceNames,
        {
          source_file_id: sourceFileId,
          display_name: capture.displayName,
        },
        { onConflict: "do-nothing" },
      );
      captureIdByBinding.set(
        routeKey(capture.sourceKey, capture.selectionKey),
        targetCaptureId,
      );
      publishedCaptures.set(capture.captureId, {
        targetCaptureId,
        sourceFileId,
      });
      appliedCaptureIds.add(targetCaptureId);
    }
    const recipeJson = canonicalJson(profile);
    const conflictsJson = canonicalJson(conflicts);
    const decisionsJson = canonicalJson(decisions);
    const bindingsJson = canonicalJson(
      captures.map((capture) => ({
        source: capture.sourceKey,
        displayName: capture.displayName,
        selection: capture.selectionKey,
        label: capture.selectionLabel,
        captureId: captureIdByBinding.get(
          routeKey(capture.sourceKey, capture.selectionKey),
        ),
      })),
    );
    const savedPlan = await transaction.query(
      `SELECT baseline_revision, state, recipe_json, conflicts_json, decisions_json, bindings_json FROM ${PLAN_TABLE} WHERE plan_id = ? AND plan_revision = ?`,
      [approved.id, approved.planRevision],
    );
    if (savedPlan[0] === undefined) {
      await insertInternalRow(
        transaction,
        DATABASE_STORAGE.tables.importPlans,
        {
          plan_id: approved.id,
          plan_revision: approved.planRevision,
          baseline_revision: approved.baselineRevision,
          state: "applied",
          recipe_json: recipeJson,
          conflicts_json: conflictsJson,
          decisions_json: decisionsJson,
          bindings_json: bindingsJson,
        },
      );
    } else if (
      valueAsBigInt(savedPlan[0]["baseline_revision"], "baseline revision") !==
        approved.baselineRevision ||
      valueAsString(savedPlan[0]["state"], "batch state") !== "applied" ||
      valueAsString(savedPlan[0]["recipe_json"], "import profile") !==
        recipeJson ||
      valueAsString(savedPlan[0]["conflicts_json"], "import conflicts") !==
        conflictsJson ||
      valueAsString(savedPlan[0]["decisions_json"], "import decisions") !==
        decisionsJson ||
      valueAsString(savedPlan[0]["bindings_json"], "source bindings") !==
        bindingsJson
    ) {
      throw databaseError(
        "DB_IMPORT_PLAN_HISTORY_CONFLICT",
        "This import batch revision conflicts with saved database history.",
        { planId: approved.id },
      );
    }
    const registeredNames = new Set<string>();
    const routeSchemas = new Map<string, TableSchema>();
    for (const row of registeredRows) {
      const name = valueAsString(row["table_name"], "table name");
      registeredNames.add(identifierKey(name));
      routeSchemas.set(
        identifierKey(name),
        parseStoredTableSchema(
          valueAsString(row["schema_json"], "table schema"),
          name,
        ),
      );
    }
    for (const route of profile.routes) {
      if (route.destination.kind === "new-table") {
        routeSchemas.set(
          identifierKey(route.destination.schema.name),
          route.destination.schema,
        );
      }
    }
    for (const schema of routeSchemas.values()) {
      for (const foreignKey of schema.foreignKeys ?? []) {
        if (!routeSchemas.has(identifierKey(foreignKey.referencesTable))) {
          throw databaseError(
            "DB_FOREIGN_TABLE_NOT_FOUND",
            "An import relationship refers to a table that does not exist or appear in this batch.",
            { table: schema.name, referencesTable: foreignKey.referencesTable },
          );
        }
      }
    }
    const orderedRoutes = orderRoutesByReferences(profile.routes, [
      ...routeSchemas.values(),
    ]);
    const plannedTableNames = new Set(
      profile.routes.flatMap((route) =>
        route.destination.kind === "new-table"
          ? [identifierKey(route.destination.schema.name)]
          : [],
      ),
    );
    for (const schema of sortTablesByReferences([...routeSchemas.values()])) {
      const key = identifierKey(schema.name);
      if (!plannedTableNames.has(key) || registeredNames.has(key)) continue;
      await createManagedTable(transaction, options.database.format, schema);
      registeredNames.add(key);
      tablesCreated += 1;
    }
    let rowCountersChecked = false;
    for (const route of orderedRoutes) {
      const capture = captures.find(
        (candidate) =>
          candidate.sourceKey === route.source &&
          candidate.selectionKey === route.selection,
      );
      if (capture === undefined) {
        throw databaseError(
          "DB_STALE_IMPORT_PLAN",
          "An import route no longer has captured rows.",
        );
      }
      const captureId = captureIdByBinding.get(
        routeKey(route.source, route.selection),
      );
      if (captureId === undefined) {
        throw databaseError(
          "DB_STALE_IMPORT_PLAN",
          "An import route no longer has a captured identity.",
        );
      }
      if (route.destination.kind === "new-table-infer") {
        throw databaseError(
          "DB_IMPORT_NEEDS_REVIEW",
          "An inferred destination schema must be approved before applying.",
        );
      }
      const tableName =
        route.destination.kind === "new-table"
          ? route.destination.schema.name
          : route.destination.table;
      const schema = await registeredSchema(transaction, tableName);
      if (schema === null) {
        throw databaseError(
          "DB_STALE_IMPORT_PLAN",
          `The destination table "${tableName}" no longer exists.`,
        );
      }
      const targetColumns = new Map(
        schema.columns.map((column) => [identifierKey(column.name), column]),
      );
      const columns = routeColumns(route, capture).map((column) => {
        const target = targetColumns.get(identifierKey(column.target));
        if (target === undefined) {
          throw databaseError(
            "DB_STALE_IMPORT_PLAN",
            `The destination column "${column.target}" no longer exists in table "${tableName}".`,
            { table: tableName, column: column.target },
          );
        }
        return { ...column, target: target.name };
      });
      const applicationKey = effectiveApplicationKey({
        captureId,
        tableName,
        schema,
        route,
        capture,
      });
      const application = await inspectApplicationIdentity({
        transaction,
        captureId,
        tableName,
        schema,
        route,
        capture,
      });
      if (application.state !== "not-applied") {
        if (application.state === "mapping-conflict") {
          throw databaseError(
            "DB_IMPORT_APPLICATION_CONFLICT",
            `This captured selection was already loaded into table "${tableName}" with a different column mapping. Choose another destination table to retain this interpretation.`,
            { captureId, table: tableName },
          );
        }
        importIds.push(application.importId);
        rowsReused += Number(application.rowCount);
        continue;
      }
      if (!rowCountersChecked) {
        await assertRowAllocationCounters(transaction);
        rowCountersChecked = true;
      }
      const importId = await allocate(transaction, "import", "IMP");
      const sourceFileRows = await transaction.query(
        `SELECT source_file_id FROM ${CAPTURE_TABLE} WHERE capture_id = ?`,
        [captureId],
      );
      const sourceFileId = valueAsString(
        sourceFileRows[0]?.["source_file_id"],
        "source file ID",
      );
      const recordRows = await transaction.query(
        `SELECT next_record_id FROM ${TABLE_REGISTRY_TABLE} WHERE table_name = ?`,
        [tableName],
      );
      let nextRecord = valueAsPositiveBigInt(
        recordRows[0]?.["next_record_id"],
        "Record ID counter",
      );
      await assertRecordAllocationAvailable({
        transaction,
        schema,
        nextRecord,
        count: capture.rowCount,
      });
      const importedRowRows = await transaction.query(
        `SELECT next_value FROM ${COUNTERS_TABLE} WHERE counter_name = ?`,
        ["imported_row"],
      );
      let nextImportedRow = valueAsPositiveBigInt(
        importedRowRows[0]?.["next_value"],
        "imported row counter",
      );
      let sourceRowCursor: bigint | undefined;
      let loadedRows = 0n;
      const rowChecksum = verifiedCaptureIds.has(captureId)
        ? undefined
        : createCaptureRowChecksum();
      while (true) {
        throwIfAborted(options.signal, "db.import.apply");
        const page = await readSourceRowPage({
          engine: transaction,
          table: CAPTURE_ROW_TABLE,
          captureId,
          cursor: sourceRowCursor,
          limit: CAPTURE_BATCH_ROWS,
          owner: "database",
        });
        throwIfAborted(options.signal, "db.import.apply");
        if (page.rows.length === 0) break;
        const output: Array<readonly EngineValue[]> = [];
        for (const { row, sourceRow } of page.rows) {
          const importedRowId = nextImportedRow++;
          const valuesJson = valueAsString(
            row["values_json"],
            "captured values",
          );
          rowChecksum?.update(sourceRow, valuesJson);
          const values = parseImportCellsJson(valuesJson);
          output.push([
            formatRecordId(schema.recordId, nextRecord++),
            importedRowId,
            importId,
            sourceFileId,
            capture.selectionLabel,
            sourceRow,
            ...columns.map((column) =>
              valueForColumn(
                values[column.source],
                targetColumns.get(identifierKey(column.target)) ?? {
                  name: column.target,
                  type: column.type,
                },
              ),
            ),
          ]);
        }
        await transaction.bulkInsert({
          table: tableName,
          columns: [
            "record_id",
            "_imported_row_id",
            "_import_id",
            "_source_file_id",
            "_source_selection",
            "_source_row",
            ...columns.map((column) => column.target),
          ],
          rows: output,
          signal: options.signal,
        });
        rowsImported += output.length;
        loadedRows += BigInt(output.length);
        sourceRowCursor = page.cursor;
      }
      assertCaptureRowCount({
        owner: "database",
        captureId,
        expected: capture.rowCount,
        actual: loadedRows,
      });
      if (rowChecksum !== undefined) {
        assertCaptureRowChecksum({
          owner: "database",
          captureId,
          expected: capture.rowChecksum,
          actual: rowChecksum.digest(),
        });
        verifiedCaptureIds.add(captureId);
      }
      await transaction.execute(
        `UPDATE ${COUNTERS_TABLE} SET next_value = ? WHERE counter_name = ?`,
        [nextImportedRow, "imported_row"],
      );
      await transaction.execute(
        `UPDATE ${TABLE_REGISTRY_TABLE} SET next_record_id = ? WHERE table_name = ?`,
        [nextRecord, tableName],
      );
      await insertInternalRow(
        transaction,
        DATABASE_STORAGE.tables.importApplications,
        {
          import_id: importId,
          application_key: applicationKey,
          request_id: options.requestId,
          capture_id: captureId,
          table_name: tableName,
          plan_id: approved.id,
          plan_revision: approved.planRevision,
          row_count: capture.rowCount,
        },
      );
      importIds.push(importId);
    }
    await insertInternalRow(
      transaction,
      DATABASE_STORAGE.tables.importRequests,
      {
        request_id: options.requestId,
        plan_id: approved.id,
        plan_revision: approved.planRevision,
        import_ids_json: JSON.stringify(importIds),
        capture_ids_json: JSON.stringify([...appliedCaptureIds].sort()),
        row_count: BigInt(rowsImported + rowsReused),
      },
    );
    if (delivery !== undefined) {
      batchId = await allocate(transaction, "delivery", "DEL");
      batchesRecorded = 1;
      await insertInternalRow(transaction, DATABASE_STORAGE.tables.deliveries, {
        delivery_id: batchId,
        request_id: options.requestId,
        context_json: canonicalJson(delivery),
      });
      for (const captureId of new Set(captureIdByBinding.values())) {
        await insertInternalRow(
          transaction,
          DATABASE_STORAGE.tables.deliveryMemberships,
          { delivery_id: batchId, capture_id: captureId },
        );
      }
    }
    captureIds = [...appliedCaptureIds].sort();
    await transaction.execute(
      `UPDATE ${DATABASE_METADATA_TABLE} SET revision = revision + 1`,
    );
    databaseWrite = "committed";
    throwIfAborted(options.signal, "db.import.apply");
  });
  return {
    databaseWrite,
    operation: "db.import.apply",
    artifacts: [],
    warnings: [],
    metrics: {
      rowsImported,
      rowsReused,
      tablesCreated,
      batchesRecorded,
    },
    importIds,
    captureIds,
    ...(batchId === undefined ? {} : { batchId }),
  };
}
