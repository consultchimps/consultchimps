import { isConsultChimpsError } from "@consultchimps/core";
import { sha256 } from "@noble/hashes/sha2.js";
import { bytesToHex } from "@noble/hashes/utils.js";

import { valueAsString } from "../database.js";
import { databaseError } from "../errors.js";
import type { EngineTransaction } from "../internal/engine.js";
import { canonicalJson } from "../internal/json.js";
import { PLAN_TABLE } from "../metadata.js";
import { identifierKey, type TableSchema } from "../schema.js";
import { parseImportRecipe } from "../validators.js";
import { routeColumns, routeKey, type PreparedCapture } from "./planning.js";
import type { ImportRecipe } from "./types.js";

const APPLICATION_KEY_PREFIX = "mapping-v1:";

function corruptHistory(message: string, cause?: unknown): never {
  throw databaseError("DB_CORRUPT_DATABASE", message, undefined, cause);
}

function destinationTable(
  route: ImportRecipe["routes"][number],
): string | undefined {
  if (route.destination.kind === "existing-table") {
    return route.destination.table;
  }
  if (route.destination.kind === "new-table") {
    return route.destination.schema.name;
  }
  return undefined;
}

export function effectiveApplicationKey(options: {
  readonly captureId: string;
  readonly tableName: string;
  readonly schema: TableSchema;
  readonly route: ImportRecipe["routes"][number];
  readonly capture: PreparedCapture;
}): string {
  const targets = new Map(
    options.schema.columns.map((column) => [
      identifierKey(column.name),
      column.name,
    ]),
  );
  const mappings = routeColumns(options.route, options.capture)
    .map((column) => {
      const target = targets.get(identifierKey(column.target));
      if (target === undefined) {
        throw databaseError(
          "DB_STALE_IMPORT_PLAN",
          `The destination column "${column.target}" no longer exists in table "${options.tableName}".`,
          { table: options.tableName, column: column.target },
        );
      }
      return { source: column.source, target: identifierKey(target) };
    })
    .sort(
      (left, right) =>
        (left.target < right.target
          ? -1
          : left.target > right.target
            ? 1
            : 0) ||
        (left.source < right.source ? -1 : left.source > right.source ? 1 : 0),
    );
  return `${APPLICATION_KEY_PREFIX}${bytesToHex(
    sha256(
      new TextEncoder().encode(
        canonicalJson([
          options.captureId,
          identifierKey(options.tableName),
          mappings,
        ]),
      ),
    ),
  )}`;
}

function legacyApplicationKey(options: {
  readonly captureId: string;
  readonly tableName: string;
  readonly route: ImportRecipe["routes"][number];
  readonly capture: PreparedCapture;
}): string {
  return bytesToHex(
    sha256(
      new TextEncoder().encode(
        canonicalJson([
          options.captureId,
          options.tableName,
          options.route.destination,
          routeColumns(options.route, options.capture),
        ]),
      ),
    ),
  );
}

function parseHistoryJson(value: unknown, field: string): unknown {
  try {
    return JSON.parse(valueAsString(value, field));
  } catch (cause) {
    return corruptHistory(
      "The saved import application has invalid plan history. Restore a verified database copy before retrying.",
      cause,
    );
  }
}

function historicalBindingKeys(value: unknown, captureId: string): Set<string> {
  if (!Array.isArray(value)) {
    return corruptHistory(
      "The saved import application has invalid source bindings. Restore a verified database copy before retrying.",
    );
  }
  const keys = new Set<string>();
  for (const binding of value) {
    if (
      typeof binding !== "object" ||
      binding === null ||
      Array.isArray(binding)
    ) {
      return corruptHistory(
        "The saved import application has invalid source bindings. Restore a verified database copy before retrying.",
      );
    }
    const fields = binding as Record<string, unknown>;
    if (
      typeof fields["captureId"] !== "string" ||
      typeof fields["source"] !== "string" ||
      typeof fields["selection"] !== "string"
    ) {
      return corruptHistory(
        "The saved import application has invalid source bindings. Restore a verified database copy before retrying.",
      );
    }
    if (fields["captureId"] === captureId) {
      keys.add(routeKey(fields["source"], fields["selection"]));
    }
  }
  if (keys.size === 0) {
    return corruptHistory(
      "The saved import application is missing its source binding. Restore a verified database copy before retrying.",
    );
  }
  return keys;
}

export async function historicalEffectiveApplicationKey(options: {
  readonly transaction: EngineTransaction;
  readonly storedKey: string;
  readonly planId: string;
  readonly planRevision: bigint;
  readonly captureId: string;
  readonly tableName: string;
  readonly schema: TableSchema;
  readonly capture: PreparedCapture;
}): Promise<string> {
  if (/^mapping-v1:[0-9a-f]{64}$/u.test(options.storedKey)) {
    return options.storedKey;
  }
  if (!/^[0-9a-f]{64}$/u.test(options.storedKey)) {
    return corruptHistory(
      "The saved import application has an invalid identity. Restore a verified database copy before retrying.",
    );
  }
  const rows = await options.transaction.query(
    `SELECT state, recipe_json, bindings_json FROM ${PLAN_TABLE} WHERE plan_id = ? AND plan_revision = ?`,
    [options.planId, options.planRevision],
  );
  if (
    rows.length !== 1 ||
    valueAsString(rows[0]?.["state"], "plan state") !== "applied"
  ) {
    return corruptHistory(
      "The saved import application is missing its applied plan history. Restore a verified database copy before retrying.",
    );
  }
  let recipe: ImportRecipe;
  try {
    recipe = parseImportRecipe(
      parseHistoryJson(rows[0]?.["recipe_json"], "import recipe"),
    );
  } catch (cause) {
    return corruptHistory(
      "The saved import application has invalid plan history. Restore a verified database copy before retrying.",
      cause,
    );
  }
  const bindings = historicalBindingKeys(
    parseHistoryJson(rows[0]?.["bindings_json"], "source bindings"),
    options.captureId,
  );
  const routes = recipe.routes.filter(
    (route) =>
      bindings.has(routeKey(route.source, route.selection)) &&
      destinationTable(route) !== undefined &&
      identifierKey(destinationTable(route)!) ===
        identifierKey(options.tableName),
  );
  const legacyMatches = routes.filter(
    (route) =>
      legacyApplicationKey({
        captureId: options.captureId,
        tableName: options.tableName,
        route,
        capture: options.capture,
      }) === options.storedKey,
  );
  let effectiveKeys: Set<string>;
  try {
    effectiveKeys = new Set(
      legacyMatches.map((route) =>
        effectiveApplicationKey({
          captureId: options.captureId,
          tableName: options.tableName,
          schema: options.schema,
          route,
          capture: options.capture,
        }),
      ),
    );
  } catch (cause) {
    if (!isConsultChimpsError(cause)) throw cause;
    return corruptHistory(
      "The saved import application has invalid column mapping history. Restore a verified database copy before retrying.",
      cause,
    );
  }
  if (effectiveKeys.size !== 1) {
    return corruptHistory(
      "The saved import application does not identify one column mapping. Restore a verified database copy before retrying.",
    );
  }
  return [...effectiveKeys][0]!;
}
