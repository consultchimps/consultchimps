import {
  engineOf,
  parseStoredTableSchema,
  valueAsBigInt,
  valueAsString,
} from "../database.js";
import { databaseError } from "../errors.js";
import { CAPTURE_ROW_TABLE, TABLE_REGISTRY_TABLE } from "../metadata.js";
import { identifierKey } from "../schema.js";
import {
  PREPARED_CAPTURE_TABLE,
  PREPARED_ROW_TABLE,
  preparedEngineOf,
  readPreparedReview,
} from "../prepared.js";
import { parseImportCellsJson } from "./inference.js";
import { inspectApplicationIdentity } from "./application-identity.js";
import { preparedCaptures, routeColumns, routeKey } from "./planning.js";
import type {
  ImportInspection,
  ImportExample,
  PreparedImportPage,
  PrepareImportOptions,
  ImportRouteInspection,
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
      "The import plan has an invalid source row number.",
      { sourceRow: sourceRow.toString() },
    );
  }
  return number;
}

export async function inspectImport(options: {
  readonly database?: PrepareImportOptions["database"] | undefined;
  readonly prepared: PrepareImportOptions["prepared"];
  readonly page: PreparedImportPage;
}): Promise<ImportInspection> {
  if (
    !Number.isInteger(options.page.limit) ||
    options.page.limit < 1 ||
    options.page.limit > 100
  ) {
    throw databaseError(
      "DB_INVALID_PAGE_SIZE",
      "Choose an import preview page size from 1 to 100.",
    );
  }
  if (
    (options.page.source === undefined) !==
    (options.page.selection === undefined)
  ) {
    throw databaseError(
      "DB_INVALID_PREVIEW_FILTER",
      "Choose both a source and selection when filtering an import preview.",
    );
  }
  const cursor = parseCursor(options.page.cursor);
  if (
    options.page.source !== undefined &&
    cursor !== undefined &&
    (cursor.source !== options.page.source ||
      cursor.selection !== options.page.selection)
  ) {
    throw databaseError(
      "DB_INVALID_CURSOR",
      "The import preview cursor belongs to a different source selection.",
    );
  }
  const engine = preparedEngineOf(options.prepared);
  if (
    options.database !== undefined &&
    options.database.id !== options.prepared.databaseId
  ) {
    throw databaseError(
      "DB_IMPORT_DATABASE_MISMATCH",
      "This import plan belongs to a different database.",
    );
  }
  const review = await readPreparedReview(options.prepared);
  const { recipe, conflicts } = review;
  const preparedCaptureList = await preparedCaptures(options.prepared);
  const selected = preparedCaptureList.filter(
    (capture) =>
      options.page.source === undefined ||
      (capture.sourceKey === options.page.source &&
        capture.selectionKey === options.page.selection),
  );
  const cursorIndex =
    cursor === undefined
      ? 0
      : selected.findIndex(
          (capture) =>
            capture.sourceKey === cursor.source &&
            capture.selectionKey === cursor.selection,
        );
  if (cursorIndex < 0) {
    throw databaseError(
      "DB_INVALID_CURSOR",
      "The import preview cursor does not identify a captured selection.",
    );
  }
  const previewWarnings: ImportInspection["previewWarnings"][number][] = [];
  const pageExamples: ImportExample[] = [];
  for (const capture of selected.slice(cursorIndex)) {
    const rowEngine = capture.reused
      ? options.database === undefined
        ? undefined
        : engineOf(options.database)
      : engine;
    if (rowEngine === undefined) {
      previewWarnings.push({
        code: "DB_PREVIEW_DATABASE_REQUIRED",
        source: capture.sourceKey,
        selection: capture.selectionKey,
        message:
          "These captured rows are stored in the target database. Supply that database to preview them.",
      });
      continue;
    }
    if (pageExamples.length > options.page.limit) continue;
    const rowTable = capture.reused ? CAPTURE_ROW_TABLE : PREPARED_ROW_TABLE;
    const afterRow =
      cursor !== undefined &&
      capture.sourceKey === cursor.source &&
      capture.selectionKey === cursor.selection
        ? cursor.sourceRow
        : 0n;
    const rows = await rowEngine.query(
      `SELECT source_row, values_json FROM ${rowTable} WHERE capture_id = ? AND source_row > ? ORDER BY source_row LIMIT ?`,
      [
        capture.captureId,
        afterRow,
        BigInt(options.page.limit + 1 - pageExamples.length),
      ],
    );
    pageExamples.push(
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
  const recipes = new Map(
    recipe.routes.map((route) => [
      routeKey(route.source, route.selection),
      route,
    ]),
  );
  const captures = await engine.query(
    `SELECT sum(row_count) AS count FROM ${PREPARED_CAPTURE_TABLE} WHERE reused = 0`,
  );
  const examples = pageExamples.slice(0, options.page.limit);
  const last = examples.at(-1);
  const registeredTables =
    options.database === undefined
      ? []
      : await engineOf(options.database).query(
          `SELECT table_name, schema_json FROM ${TABLE_REGISTRY_TABLE}`,
        );
  const tables = new Map(
    registeredTables.map((row) => {
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
  const routes: ImportRouteInspection[] = await Promise.all(
    preparedCaptureList.map(async (capture) => {
      const route = recipes.get(
        routeKey(capture.sourceKey, capture.selectionKey),
      );
      let applicationState: ImportRouteInspection["applicationState"] =
        options.database === undefined ? "not-checked" : "unresolved";
      if (
        options.database !== undefined &&
        route !== undefined &&
        route.destination.kind !== "new-table-infer"
      ) {
        const requestedTable =
          route.destination.kind === "new-table"
            ? route.destination.schema.name
            : route.destination.table;
        const registered = tables.get(identifierKey(requestedTable));
        if (registered === undefined) {
          applicationState =
            route.destination.kind === "new-table"
              ? "not-applied"
              : "unresolved";
        } else {
          applicationState = (
            await inspectApplicationIdentity({
              transaction: engineOf(options.database),
              captureId: capture.captureId,
              tableName: registered.name,
              schema: registered.schema,
              route,
              capture,
            })
          ).state;
        }
      }
      return {
        source: capture.sourceKey,
        selection: capture.selectionKey,
        label: capture.selectionLabel,
        captureId: capture.captureId,
        reused: capture.reused,
        applicationState,
        rowCount: capture.rowCount,
        destination: route?.destination ?? null,
        columns: route === undefined ? [] : routeColumns(route, capture),
        inferredColumns: capture.columns,
      };
    }),
  );
  return {
    prepared: review.prepared,
    conflicts,
    capturedRows: valueAsBigInt(captures[0]?.["count"] ?? 0n, "captured rows"),
    routes,
    examples,
    previewWarnings,
    ...(pageExamples.length > options.page.limit && last !== undefined
      ? {
          nextCursor: JSON.stringify([
            last.source,
            last.selection,
            String(last.sourceRow),
          ]),
        }
      : {}),
  };
}
