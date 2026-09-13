import { valueAsNonNegativeBigInt, valueAsString } from "../database.js";
import { databaseError } from "../errors.js";
import type { EngineTransaction } from "../internal/engine.js";
import { APPLICATION_TABLE, CAPTURE_TABLE, PLAN_TABLE } from "../metadata.js";
import { identifierKey } from "../schema.js";
import { historicalCaptureBindings } from "./application-identity.js";
import { routeKey } from "./planning.js";
import type { ImportRecipe } from "./types.js";

const RECEIPT_APPLICATION_QUERY_VALUES = 400;

export function receiptIds(value: unknown): string[] {
  let parsed: unknown;
  try {
    parsed = JSON.parse(valueAsString(value, "receipt identifiers"));
  } catch {
    parsed = undefined;
  }
  if (
    !Array.isArray(parsed) ||
    !parsed.every((item: unknown) => typeof item === "string")
  ) {
    throw databaseError(
      "DB_CORRUPT_METADATA",
      "The saved import receipt contains invalid identifiers. Restore a verified database copy before retrying.",
    );
  }
  return parsed;
}

function corruptReceiptRowCount(): never {
  throw databaseError(
    "DB_CORRUPT_DATABASE",
    "The saved import receipt row count does not match its captured applications. Restore a verified database copy before retrying.",
  );
}

function corruptReceiptApplications(): never {
  throw databaseError(
    "DB_CORRUPT_DATABASE",
    "The saved import receipt does not match its approved import routes. Restore a verified database copy before retrying.",
  );
}

function corruptReceiptCaptures(): never {
  throw databaseError(
    "DB_CORRUPT_DATABASE",
    "The saved import receipt capture identifiers do not match its source bindings. Restore a verified database copy before retrying.",
  );
}

function receiptApplicationIdentity(
  captureId: string,
  tableName: string,
): string {
  return JSON.stringify([captureId, identifierKey(tableName)]);
}

function incrementIdentityCount(
  counts: Map<string, number>,
  key: string,
): void {
  counts.set(key, (counts.get(key) ?? 0) + 1);
}

export async function validatedReceiptRowCount(options: {
  readonly transaction: EngineTransaction;
  readonly importIds: readonly string[];
  readonly captureIds: readonly string[];
  readonly storedRowCount: unknown;
  readonly recipe: ImportRecipe;
  readonly planId: string;
  readonly planRevision: bigint;
}): Promise<bigint> {
  const receiptRowCount = valueAsNonNegativeBigInt(
    options.storedRowCount,
    "receipt row count",
  );
  const planRows = await options.transaction.query(
    `SELECT state, bindings_json FROM ${PLAN_TABLE} WHERE plan_id = ? AND plan_revision = ?`,
    [options.planId, options.planRevision],
  );
  if (
    planRows.length !== 1 ||
    valueAsString(planRows[0]?.["state"], "plan state") !== "applied"
  ) {
    corruptReceiptApplications();
  }
  const bindings = historicalCaptureBindings(planRows[0]?.["bindings_json"]);
  const expectedCaptureIds = new Set(bindings.values());
  const actualCaptureIds = new Set(options.captureIds);
  if (
    actualCaptureIds.size !== options.captureIds.length ||
    actualCaptureIds.size !== expectedCaptureIds.size ||
    [...expectedCaptureIds].some(
      (captureId) => !actualCaptureIds.has(captureId),
    )
  ) {
    corruptReceiptCaptures();
  }
  const expectedIdentities = new Map<string, number>();
  for (const route of options.recipe.routes) {
    const captureId = bindings.get(routeKey(route.source, route.selection));
    if (captureId === undefined) corruptReceiptApplications();
    const tableName =
      route.destination.kind === "new-table"
        ? route.destination.schema.name
        : route.destination.kind === "existing-table"
          ? route.destination.table
          : undefined;
    if (tableName === undefined) corruptReceiptApplications();
    incrementIdentityCount(
      expectedIdentities,
      receiptApplicationIdentity(captureId, tableName),
    );
  }
  const applications = new Map<
    string,
    { readonly rowCount: bigint; readonly identity: string }
  >();
  const uniqueImportIds = [...new Set(options.importIds)];
  for (
    let offset = 0;
    offset < uniqueImportIds.length;
    offset += RECEIPT_APPLICATION_QUERY_VALUES
  ) {
    const ids = uniqueImportIds.slice(
      offset,
      offset + RECEIPT_APPLICATION_QUERY_VALUES,
    );
    const rows = await options.transaction.query(
      `SELECT application.import_id, application.capture_id, application.table_name, application.row_count AS application_row_count, capture.row_count AS capture_row_count FROM ${APPLICATION_TABLE} AS application LEFT JOIN ${CAPTURE_TABLE} AS capture ON capture.capture_id = application.capture_id WHERE application.import_id IN (${ids.map(() => "?").join(", ")})`,
      ids,
    );
    for (const row of rows) {
      const importId = valueAsString(row["import_id"], "import ID");
      const applicationRowCount = valueAsNonNegativeBigInt(
        row["application_row_count"],
        "import application row count",
      );
      const captureRowCount = valueAsNonNegativeBigInt(
        row["capture_row_count"],
        "capture row count",
      );
      if (applicationRowCount !== captureRowCount) corruptReceiptRowCount();
      applications.set(importId, {
        rowCount: applicationRowCount,
        identity: receiptApplicationIdentity(
          valueAsString(row["capture_id"], "capture ID"),
          valueAsString(row["table_name"], "table name"),
        ),
      });
    }
  }
  let applicationRowCount = 0n;
  const actualIdentities = new Map<string, number>();
  for (const importId of options.importIds) {
    const application = applications.get(importId);
    if (application === undefined) corruptReceiptRowCount();
    applicationRowCount += application.rowCount;
    incrementIdentityCount(actualIdentities, application.identity);
  }
  if (receiptRowCount !== applicationRowCount) corruptReceiptRowCount();
  if (
    actualIdentities.size !== expectedIdentities.size ||
    [...expectedIdentities].some(
      ([identity, count]) => actualIdentities.get(identity) !== count,
    )
  ) {
    corruptReceiptApplications();
  }
  return receiptRowCount;
}
