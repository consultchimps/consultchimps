import { valueAsBigInt, valueAsString } from "../database.js";
import { databaseError } from "../errors.js";
import {
  PREPARED_BINDING_TABLE,
  PREPARED_CAPTURE_TABLE,
  PREPARED_ROW_TABLE,
  preparedEngineOf,
  preparedRef,
  readPreparedRecipe,
} from "../prepared.js";
import { parseImportCellsJson } from "./inference.js";
import { preparedCaptures, routeColumns, routeKey } from "./planning.js";
import type {
  ImportInspection,
  PreparedImportPage,
  PrepareImportOptions,
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

function previewQuery(
  page: PreparedImportPage,
  cursor: ImportCursor | undefined,
) {
  const filtered = page.source !== undefined;
  const filterSql = filtered
    ? ` WHERE b.source_key = ? AND b.selection_key = ?${cursor === undefined ? "" : " AND r.source_row > ?"}`
    : cursor === undefined
      ? ""
      : " WHERE b.source_key > ? OR (b.source_key = ? AND (b.selection_key > ? OR (b.selection_key = ? AND r.source_row > ?)))";
  const values = filtered
    ? [
        page.source ?? "",
        page.selection ?? "",
        ...(cursor === undefined ? [] : [cursor.sourceRow]),
        BigInt(page.limit + 1),
      ]
    : cursor === undefined
      ? [BigInt(page.limit + 1)]
      : [
          cursor.source,
          cursor.source,
          cursor.selection,
          cursor.selection,
          cursor.sourceRow,
          BigInt(page.limit + 1),
        ];
  return { filterSql, values };
}

export async function inspectImport(options: {
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
  const query = previewQuery(options.page, cursor);
  const rows = await engine.query(
    `SELECT b.source_key, b.selection_key, r.source_row, r.values_json FROM ${PREPARED_ROW_TABLE} r JOIN ${PREPARED_BINDING_TABLE} b ON b.capture_id = r.capture_id${query.filterSql} ORDER BY b.source_key, b.selection_key, r.source_row LIMIT ?`,
    query.values,
  );
  const { recipe, conflicts } = await readPreparedRecipe(options.prepared);
  const preparedCaptureList = await preparedCaptures(options.prepared);
  const recipes = new Map(
    recipe.routes.map((route) => [
      routeKey(route.source, route.selection),
      route,
    ]),
  );
  const captures = await engine.query(
    `SELECT sum(row_count) AS count FROM ${PREPARED_CAPTURE_TABLE} WHERE reused = 0`,
  );
  const examples = rows.slice(0, options.page.limit).map((row) => ({
    source: valueAsString(row["source_key"], "source key"),
    selection: valueAsString(row["selection_key"], "selection key"),
    sourceRow: sourceRowNumber(row["source_row"]),
    values: parseImportCellsJson(
      valueAsString(row["values_json"], "captured values"),
    ),
  }));
  const last = examples.at(-1);
  return {
    prepared: await preparedRef(options.prepared),
    conflicts,
    capturedRows: valueAsBigInt(captures[0]?.["count"] ?? 0n, "captured rows"),
    routes: preparedCaptureList.map((capture) => {
      const route = recipes.get(
        routeKey(capture.sourceKey, capture.selectionKey),
      );
      return {
        source: capture.sourceKey,
        selection: capture.selectionKey,
        label: capture.selectionLabel,
        captureId: capture.captureId,
        reused: capture.reused,
        rowCount: capture.rowCount,
        destination: route?.destination ?? null,
        columns: route === undefined ? [] : routeColumns(route, capture),
        inferredColumns: capture.columns,
      };
    }),
    examples,
    ...(rows.length > options.page.limit && last !== undefined
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
