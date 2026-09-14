import { isConsultChimpsError } from "@consultchimps/core";
import { sha256 } from "@noble/hashes/sha2.js";
import { bytesToHex } from "@noble/hashes/utils.js";

import {
  valueAsBigInt,
  valueAsNonNegativeBigInt,
  valueAsString,
} from "../database.js";
import { databaseError } from "../errors.js";
import type { EngineRow, EngineTransaction } from "../internal/engine.js";
import { canonicalJson } from "../internal/json.js";
import { APPLICATION_TABLE, PLAN_TABLE } from "../metadata.js";
import { identifierKey, type TableSchema } from "../schema.js";
import { parseImportProfile } from "../validators.js";
import { effectiveMappingKey } from "./application-key.js";
import { routeColumns, routeKey, type PreparedCapture } from "./planning.js";
import type { ImportProfile } from "./types.js";

function corruptHistory(message: string, cause?: unknown): never {
  throw databaseError("DB_CORRUPT_DATABASE", message, undefined, cause);
}

function destinationTable(
  route: ImportProfile["routes"][number],
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
  readonly route: ImportProfile["routes"][number];
  readonly capture: PreparedCapture;
}): string {
  return effectiveMappingKey({
    captureId: options.captureId,
    tableName: options.tableName,
    schema: options.schema,
    columns: routeColumns(options.route, options.capture),
  });
}

export type ImportApplicationState =
  | { readonly state: "not-applied" }
  | {
      readonly state: "already-applied";
      readonly importId: string;
      readonly rowCount: bigint;
    }
  | { readonly state: "mapping-conflict" };

interface ApplicationIdentityInput {
  readonly captureId: string;
  readonly tableName: string;
  readonly schema: TableSchema;
  readonly route: ImportProfile["routes"][number];
  readonly capture: PreparedCapture;
}

export async function inspectApplicationIdentity(
  options: ApplicationIdentityInput & {
    readonly transaction: EngineTransaction;
  },
): Promise<ImportApplicationState> {
  const applications = await options.transaction.query(
    `SELECT import_id, application_key, plan_id, plan_revision, row_count FROM ${APPLICATION_TABLE} WHERE capture_id = ? AND table_name = ? ORDER BY import_id`,
    [options.captureId, options.tableName],
  );
  return evaluateApplicationIdentity(options, applications);
}

export async function inspectApplicationIdentities(options: {
  readonly transaction: EngineTransaction;
  readonly identities: readonly ApplicationIdentityInput[];
}): Promise<readonly ImportApplicationState[]> {
  if (options.identities.length === 0) return [];
  const applications = await options.transaction.query(
    `SELECT capture_id, table_name, import_id, application_key, plan_id, plan_revision, row_count FROM ${APPLICATION_TABLE} WHERE ${options.identities.map(() => "(capture_id = ? AND table_name = ?)").join(" OR ")} ORDER BY import_id`,
    options.identities.flatMap((identity) => [
      identity.captureId,
      identity.tableName,
    ]),
  );
  const grouped = new Map<string, EngineRow[]>();
  for (const application of applications) {
    const key = JSON.stringify([
      application["capture_id"],
      application["table_name"],
    ]);
    const rows = grouped.get(key) ?? [];
    rows.push(application);
    grouped.set(key, rows);
  }
  return Promise.all(
    options.identities.map((identity) =>
      evaluateApplicationIdentity(
        { ...identity, transaction: options.transaction },
        grouped.get(JSON.stringify([identity.captureId, identity.tableName])) ??
          [],
      ),
    ),
  );
}

async function evaluateApplicationIdentity(
  options: ApplicationIdentityInput & {
    readonly transaction: EngineTransaction;
  },
  applications: readonly EngineRow[],
): Promise<ImportApplicationState> {
  if (applications.length > 1) {
    throw databaseError(
      "DB_CORRUPT_DATABASE",
      "The database has duplicate import applications for one capture and table. Restore a verified database copy before retrying.",
      { captureId: options.captureId, table: options.tableName },
    );
  }
  const application = applications[0];
  if (application === undefined) return { state: "not-applied" };
  const rowCount = valueAsNonNegativeBigInt(
    application["row_count"],
    "import application row count",
  );
  if (rowCount !== options.capture.rowCount) {
    throw databaseError(
      "DB_CORRUPT_DATABASE",
      "The saved import application row count does not match its captured selection. Restore a verified database copy before retrying.",
      { captureId: options.captureId, table: options.tableName },
    );
  }
  let currentKey: string;
  try {
    currentKey = effectiveApplicationKey(options);
  } catch (cause) {
    if (!isConsultChimpsError(cause) || cause.code !== "DB_STALE_IMPORT_PLAN") {
      throw cause;
    }
    return { state: "mapping-conflict" };
  }
  const storedKey = await historicalEffectiveApplicationKey({
    transaction: options.transaction,
    storedKey: valueAsString(application["application_key"], "application key"),
    planId: valueAsString(application["plan_id"], "batch ID"),
    planRevision: valueAsBigInt(application["plan_revision"], "batch revision"),
    captureId: options.captureId,
    tableName: options.tableName,
    schema: options.schema,
    capture: options.capture,
  });
  return storedKey === currentKey
    ? {
        state: "already-applied",
        importId: valueAsString(application["import_id"], "import ID"),
        rowCount,
      }
    : { state: "mapping-conflict" };
}

function legacyApplicationKey(options: {
  readonly captureId: string;
  readonly tableName: string;
  readonly route: ImportProfile["routes"][number];
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
      "The saved import application has invalid batch history. Restore a verified database copy before retrying.",
      cause,
    );
  }
}

export function historicalCaptureBindings(
  value: unknown,
): ReadonlyMap<string, string> {
  const parsed = parseHistoryJson(value, "source bindings");
  if (!Array.isArray(parsed)) {
    return corruptHistory(
      "The saved import application has invalid source bindings. Restore a verified database copy before retrying.",
    );
  }
  const bindings = new Map<string, string>();
  for (const binding of parsed) {
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
    const key = routeKey(fields["source"], fields["selection"]);
    const existing = bindings.get(key);
    if (existing !== undefined && existing !== fields["captureId"]) {
      return corruptHistory(
        "The saved import application has conflicting source bindings. Restore a verified database copy before retrying.",
      );
    }
    bindings.set(key, fields["captureId"]);
  }
  return bindings;
}

function historicalBindingKeys(value: unknown, captureId: string): Set<string> {
  const keys = new Set<string>();
  for (const [key, boundCaptureId] of historicalCaptureBindings(value)) {
    if (boundCaptureId === captureId) keys.add(key);
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
    valueAsString(rows[0]?.["state"], "batch state") !== "applied"
  ) {
    return corruptHistory(
      "The saved import application is missing its applied batch history. Restore a verified database copy before retrying.",
    );
  }
  let profile: ImportProfile;
  try {
    profile = parseImportProfile(
      parseHistoryJson(rows[0]?.["recipe_json"], "import profile"),
    );
  } catch (cause) {
    return corruptHistory(
      "The saved import application has invalid batch history. Restore a verified database copy before retrying.",
      cause,
    );
  }
  const bindings = historicalBindingKeys(
    rows[0]?.["bindings_json"],
    options.captureId,
  );
  const routes = profile.routes.filter(
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
