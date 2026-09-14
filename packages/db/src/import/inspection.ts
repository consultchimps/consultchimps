import {
  engineOf,
  parseStoredTableSchema,
  valueAsBigInt,
  valueAsString,
} from "../database.js";
import { databaseError } from "../errors.js";
import {
  CAPTURE_ROW_TABLE,
  TABLE_REGISTRY_TABLE,
  DATABASE_METADATA_TABLE,
} from "../metadata.js";
import { identifierKey } from "../schema.js";
import {
  PREPARED_ROW_TABLE,
  PREPARED_CAPTURE_TABLE,
  PREPARED_BINDING_TABLE,
  readPreparedReviewSnapshot,
} from "../prepared.js";
import { parseImportCellsJson } from "./inference.js";
import { inspectAppliedImportBatch } from "./history.js";
import { suggestImportDestination } from "./profile.js";
import { inspectApplicationIdentities } from "./application-identity.js";
import {
  preparedCapturesFromEngine,
  routeColumns,
  routeKey,
} from "./planning.js";
import type {
  ImportInspection,
  ImportExample,
  ImportBatchPage,
  PrepareImportOptions,
  ImportRouteInspection,
  ImportReviewOutcome,
  ImportBatchRef,
  ReadyImportBatchRef,
} from "./types.js";

interface ImportCursor {
  readonly source: string;
  readonly selection: string;
  readonly sourceRow: bigint;
}

function parseCursor(value: string | undefined): ImportCursor | undefined {
  if (value === undefined) return undefined;
  try {
    const fields: unknown = JSON.parse(value);
    if (
      !Array.isArray(fields) ||
      fields.length !== 3 ||
      typeof fields[0] !== "string" ||
      typeof fields[1] !== "string" ||
      typeof fields[2] !== "string" ||
      !/^\d+$/u.test(fields[2])
    ) {
      throw new Error("invalid cursor fields");
    }
    return {
      source: fields[0],
      selection: fields[1],
      sourceRow: BigInt(fields[2]),
    };
  } catch {
    throw databaseError(
      "DB_INVALID_CURSOR",
      "The import preview cursor is invalid.",
    );
  }
}

function sourceRowNumber(value: unknown): number {
  const sourceRow = valueAsBigInt(value, "source row");
  const number = Number(sourceRow);
  if (!Number.isSafeInteger(number) || number < 1) {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import batch has an invalid source row number.",
      { sourceRow: sourceRow.toString() },
    );
  }
  return number;
}

interface RouteCursor {
  readonly batchId: string;
  readonly revision: string;
  readonly fingerprint: string;
  readonly source: string;
  readonly selection: string;
}

function parseRouteCursor(value: string | undefined): RouteCursor | undefined {
  if (value === undefined) return undefined;
  try {
    const fields: unknown = JSON.parse(value);
    if (
      !Array.isArray(fields) ||
      fields.length !== 5 ||
      !fields.every((field: unknown) => typeof field === "string")
    ) {
      throw new Error("invalid route cursor fields");
    }
    const [batchId, revision, fingerprint, source, selection] = fields;
    if (
      typeof batchId !== "string" ||
      typeof revision !== "string" ||
      typeof fingerprint !== "string" ||
      typeof source !== "string" ||
      typeof selection !== "string"
    ) {
      throw new Error("invalid route cursor fields");
    }
    return { batchId, revision, fingerprint, source, selection };
  } catch {
    throw databaseError(
      "DB_INVALID_CURSOR",
      "The batch route cursor is invalid.",
    );
  }
}

export function assertImportReviewPage(
  page: ImportBatchPage,
): ImportCursor | undefined {
  if (!Number.isInteger(page.limit) || page.limit < 1 || page.limit > 100) {
    throw databaseError(
      "DB_INVALID_PAGE_SIZE",
      "Choose an import preview page size from 1 to 100.",
    );
  }
  if ((page.source === undefined) !== (page.selection === undefined)) {
    throw databaseError(
      "DB_INVALID_PREVIEW_FILTER",
      "Choose both a source and selection when filtering an import preview.",
    );
  }
  const cursor = parseCursor(page.cursor);
  if (
    page.source !== undefined &&
    cursor !== undefined &&
    (cursor.source !== page.source || cursor.selection !== page.selection)
  ) {
    throw databaseError(
      "DB_INVALID_CURSOR",
      "The import preview cursor belongs to a different source selection.",
    );
  }
  return cursor;
}

export async function inspectImport(options: {
  readonly database?: PrepareImportOptions["database"] | undefined;
  readonly prepared: PrepareImportOptions["prepared"];
  readonly page: ImportBatchPage;
  readonly routePage?:
    | {
        readonly limit?: number | undefined;
        readonly cursor?: string | undefined;
      }
    | undefined;
}): Promise<ImportInspection> {
  const routeLimit = options.routePage?.limit ?? 50;
  if (!Number.isInteger(routeLimit) || routeLimit < 1 || routeLimit > 100) {
    throw databaseError(
      "DB_INVALID_PAGE_SIZE",
      "Choose a route page size from 1 to 100.",
    );
  }
  const cursor = assertImportReviewPage(options.page);
  const routeCursor = parseRouteCursor(options.routePage?.cursor);
  if (
    options.database !== undefined &&
    options.database.id !== options.prepared.databaseId
  ) {
    throw databaseError(
      "DB_IMPORT_DATABASE_MISMATCH",
      "This import batch belongs to a different database.",
    );
  }
  const target =
    options.database === undefined ? undefined : engineOf(options.database);
  const targetRevision =
    target === undefined
      ? null
      : valueAsBigInt(
          (
            await target.query(
              `SELECT revision FROM ${DATABASE_METADATA_TABLE}`,
            )
          )[0]?.["revision"],
          "database revision",
        );
  const snapshot = await readPreparedReviewSnapshot(
    options.prepared,
    async (review, transaction) => {
      if (
        routeCursor !== undefined &&
        (routeCursor.batchId !== review.prepared.id ||
          routeCursor.revision !== review.prepared.planRevision.toString() ||
          routeCursor.fingerprint !== review.prepared.reviewFingerprint)
      ) {
        throw databaseError(
          "DB_INVALID_CURSOR",
          "The batch changed since this route cursor was created. Inspect its first page again.",
        );
      }
      if (routeCursor !== undefined) {
        const found = await preparedCapturesFromEngine(transaction, {
          source: routeCursor.source,
          selection: routeCursor.selection,
          limit: 1,
        });
        if (found.length === 0)
          throw databaseError(
            "DB_INVALID_CURSOR",
            "The route cursor does not identify a captured selection.",
          );
      }
      const captures = await preparedCapturesFromEngine(transaction, {
        after: routeCursor,
        limit: routeLimit + 1,
      });
      const totals = (
        await transaction.query(
          `SELECT (SELECT COUNT(*) FROM ${PREPARED_BINDING_TABLE}) AS route_count, COALESCE(SUM(row_count), 0) AS review_rows, COALESCE(SUM(CASE WHEN reused = 0 THEN row_count ELSE 0 END), 0) AS captured_rows FROM ${PREPARED_CAPTURE_TABLE} WHERE capture_id IN (SELECT capture_id FROM ${PREPARED_BINDING_TABLE})`,
        )
      )[0];
      const captureIds = (
        await transaction.query(
          `SELECT DISTINCT capture_id FROM ${PREPARED_BINDING_TABLE} ORDER BY capture_id`,
        )
      ).map((row) => valueAsString(row["capture_id"], "capture ID"));
      const examples: ImportExample[] = [];
      const warnings: ImportInspection["previewWarnings"][number][] = [];
      let after: { source: string; selection: string } | undefined;
      let first = true;
      for (;;) {
        const filter =
          options.page.source === undefined
            ? undefined
            : {
                source: options.page.source,
                selection: options.page.selection,
              };
        const previewCaptures =
          first && cursor !== undefined
            ? await preparedCapturesFromEngine(transaction, {
                source: cursor.source,
                selection: cursor.selection,
                limit: 1,
              })
            : await preparedCapturesFromEngine(transaction, {
                ...filter,
                after,
                limit: 100,
              });
        if (first && cursor !== undefined && previewCaptures.length === 0)
          throw databaseError(
            "DB_INVALID_CURSOR",
            "The import preview cursor does not identify a captured selection.",
          );
        const wasCursorPage = first && cursor !== undefined;
        first = false;
        if (previewCaptures.length === 0) break;
        for (const capture of previewCaptures) {
          const rowEngine = capture.reused ? target : transaction;
          if (rowEngine === undefined) {
            warnings.push({
              code: "DB_PREVIEW_DATABASE_REQUIRED",
              source: capture.sourceKey,
              selection: capture.selectionKey,
              message:
                "These captured rows are stored in the target database. Supply that database to preview them.",
            });
            continue;
          }
          if (examples.length > options.page.limit) break;
          const afterRow =
            cursor !== undefined &&
            capture.sourceKey === cursor.source &&
            capture.selectionKey === cursor.selection
              ? cursor.sourceRow
              : 0n;
          const rows = await rowEngine.query(
            `SELECT source_row, values_json FROM ${capture.reused ? CAPTURE_ROW_TABLE : PREPARED_ROW_TABLE} WHERE capture_id = ? AND source_row > ? ORDER BY source_row LIMIT ?`,
            [
              capture.captureId,
              afterRow,
              BigInt(options.page.limit + 1 - examples.length),
            ],
          );
          examples.push(
            ...rows.map((row) => ({
              source: capture.sourceKey,
              selection: capture.selectionKey,
              sourceRow: sourceRowNumber(row["source_row"]),
              values: parseImportCellsJson(
                valueAsString(row["values_json"], "captured values"),
              ),
            })),
          );
        }
        const last = previewCaptures.at(-1);
        if (
          last === undefined ||
          examples.length > options.page.limit ||
          options.page.source !== undefined ||
          (!wasCursorPage && previewCaptures.length < 100)
        )
          break;
        after = { source: last.sourceKey, selection: last.selectionKey };
      }
      return {
        review,
        captures,
        captureIds,
        examples,
        warnings,
        capturedRows: valueAsBigInt(
          totals?.["captured_rows"],
          "captured row count",
        ),
        reviewRows: valueAsBigInt(totals?.["review_rows"], "review row count"),
        routeCount: valueAsBigInt(totals?.["route_count"], "route count"),
      };
    },
  );
  const { review } = snapshot;
  const profiles = new Map(
    review.profile.routes.map((route) => [
      routeKey(route.source, route.selection),
      route,
    ]),
  );
  const tableRows =
    target === undefined
      ? []
      : await target.query(
          `SELECT table_name, schema_json FROM ${TABLE_REGISTRY_TABLE}`,
        );
  const tables = new Map(
    tableRows.map((row) => {
      const name = valueAsString(row["table_name"], "table name");
      return [
        identifierKey(name),
        {
          name,
          schema: parseStoredTableSchema(
            valueAsString(row["schema_json"], "table schema"),
            name,
          ),
        },
      ] as const;
    }),
  );
  const applied =
    options.database === undefined
      ? null
      : await inspectAppliedImportBatch({
          database: options.database,
          planId: review.prepared.id,
          planRevision: review.prepared.planRevision,
        });
  const appliedCaptureByRoute = new Map(
    (applied?.bindings ?? []).map((binding) => [
      routeKey(binding.source, binding.selection),
      binding.captureId,
    ]),
  );
  const routeContexts = snapshot.captures
    .slice(0, routeLimit)
    .map((capture) => {
      const route = profiles.get(
        routeKey(capture.sourceKey, capture.selectionKey),
      );
      const tableName =
        route?.destination.kind === "new-table"
          ? route.destination.schema.name
          : route?.destination.kind === "existing-table"
            ? route.destination.table
            : undefined;
      const registered =
        tableName === undefined
          ? undefined
          : tables.get(identifierKey(tableName));
      return { capture, route, registered };
    });
  const identities = routeContexts.flatMap(({ capture, route, registered }) =>
    route === undefined ||
    route.destination.kind === "new-table-infer" ||
    registered === undefined
      ? []
      : [
          {
            captureId:
              appliedCaptureByRoute.get(
                routeKey(capture.sourceKey, capture.selectionKey),
              ) ?? capture.captureId,
            tableName: registered.name,
            schema: registered.schema,
            route,
            capture,
          },
        ],
  );
  const states =
    target === undefined
      ? []
      : await inspectApplicationIdentities({ transaction: target, identities });
  const applications = new Map(
    identities.map((identity, index) => [
      routeKey(identity.capture.sourceKey, identity.capture.selectionKey),
      states[index],
    ]),
  );
  const routes: ImportRouteInspection[] = routeContexts.map(
    ({ capture, route, registered }) => {
      const applicationState: ImportRouteInspection["applicationState"] =
        target === undefined
          ? "not-checked"
          : (applications.get(routeKey(capture.sourceKey, capture.selectionKey))
              ?.state ??
            (route?.destination.kind === "new-table"
              ? "not-applied"
              : "unresolved"));
      return {
        source: capture.sourceKey,
        displayName: capture.displayName,
        selection: capture.selectionKey,
        label: capture.selectionLabel,
        captureId: capture.captureId,
        reused: capture.reused,
        applicationState,
        rowCount: capture.rowCount,
        destination: route?.destination ?? null,
        columns: route === undefined ? [] : routeColumns(route, capture),
        inferredColumns: capture.columns,
        destinationColumns:
          registered?.schema.columns ??
          (route?.destination.kind === "new-table"
            ? route.destination.schema.columns
            : capture.columns),
        suggestedDestination:
          route?.destination.kind === "new-table-infer"
            ? route.destination
            : route?.destination.kind === "new-table"
              ? {
                  kind: "new-table-infer",
                  name: route.destination.schema.name,
                  recordId: route.destination.schema.recordId,
                }
              : registered !== undefined
                ? {
                    kind: "new-table-infer",
                    name: registered.name,
                    recordId: registered.schema.recordId,
                  }
                : suggestImportDestination(
                    capture.selectionLabel,
                    "table-name",
                  ),
      };
    },
  );
  if (
    target !== undefined &&
    targetRevision !==
      valueAsBigInt(
        (
          await target.query(`SELECT revision FROM ${DATABASE_METADATA_TABLE}`)
        )[0]?.["revision"],
        "database revision",
      )
  ) {
    throw databaseError(
      "DB_STALE_IMPORT_PLAN",
      "The database changed while the batch was inspected. Refresh the review before applying it.",
    );
  }
  const examples = snapshot.examples.slice(0, options.page.limit);
  const lastExample = examples.at(-1);
  const lastRoute = routes.at(-1);
  return {
    prepared: review.prepared,
    application:
      options.database === undefined
        ? { state: "not-checked" }
        : applied === null
          ? { state: "pending" }
          : {
              state: "applied",
              captureIds: [
                ...new Set(
                  applied.bindings.map((binding) => binding.captureId),
                ),
              ].sort(),
            },
    targetRevision,
    conflicts: review.conflicts,
    capturedRows: snapshot.capturedRows,
    reviewRows: snapshot.reviewRows,
    routeCount: snapshot.routeCount,
    captureIds: snapshot.captureIds,
    routes,
    examples,
    previewWarnings: snapshot.warnings,
    ...(snapshot.captures.length > routeLimit && lastRoute !== undefined
      ? {
          nextRouteCursor: JSON.stringify([
            review.prepared.id,
            review.prepared.planRevision.toString(),
            review.prepared.reviewFingerprint,
            lastRoute.source,
            lastRoute.selection,
          ]),
        }
      : {}),
    ...(snapshot.examples.length > options.page.limit &&
    lastExample !== undefined
      ? {
          nextCursor: JSON.stringify([
            lastExample.source,
            lastExample.selection,
            String(lastExample.sourceRow),
          ]),
        }
      : {}),
  };
}

export async function inspectUpdatedImport(options: {
  readonly database: PrepareImportOptions["database"];
  readonly prepared: PrepareImportOptions["prepared"];
  readonly expected: ImportBatchRef | ReadyImportBatchRef;
  readonly page: ImportBatchPage;
}): Promise<ImportReviewOutcome> {
  let inspection: ImportInspection;
  try {
    inspection = await inspectImport(options);
  } catch (cause) {
    throw databaseError(
      "DB_BATCH_REVIEW_REFRESH_REQUIRED",
      "The batch was updated, but its review could not be refreshed. Inspect the saved batch again before applying it.",
      { batchUpdated: true, batchId: options.expected.id },
      cause,
    );
  }
  if (
    inspection.prepared.reviewFingerprint !== options.expected.reviewFingerprint
  ) {
    throw databaseError(
      "DB_STALE_IMPORT_PLAN",
      "The batch changed again while its updated review was being read. Inspect its latest revision before applying it.",
      { batchUpdated: true, batchId: options.expected.id },
    );
  }
  return { prepared: inspection.prepared, inspection };
}
