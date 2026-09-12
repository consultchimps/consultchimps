import { engineOf, inspectDatabase } from "../database.js";
import { databaseError } from "../errors.js";
import { APPLICATION_TABLE } from "../metadata.js";
import { readPreparedRecipe, updatePreparedPlan } from "../prepared.js";
import type { preparedRef } from "../prepared.js";
import { readSchemaFingerprint } from "../records.js";
import { evaluateConflicts, preparedCaptures, routeKey } from "./planning.js";
import type {
  ImportDecision,
  ImportRecipe,
  PrepareImportOptions,
} from "./types.js";

export async function resolveImport(options: {
  readonly database: PrepareImportOptions["database"];
  readonly prepared: PrepareImportOptions["prepared"];
  readonly decisions: readonly ImportDecision[];
  readonly rebase?: boolean | undefined;
}): Promise<
  ReturnType<typeof preparedRef> extends Promise<infer T> ? T : never
> {
  if (options.prepared.databaseId !== options.database.id) {
    throw databaseError(
      "DB_IMPORT_DATABASE_MISMATCH",
      "This import plan belongs to a different database.",
    );
  }
  const baseline = options.rebase
    ? {
        inspection: await inspectDatabase({ database: options.database }),
        schemaFingerprint: await readSchemaFingerprint(
          engineOf(options.database),
          options.database.format,
        ),
      }
    : undefined;
  const current = await readPreparedRecipe(options.prepared);
  const captures = await preparedCaptures(options.prepared);
  const routes = new Map(
    current.recipe.routes.map((route) => [
      routeKey(route.source, route.selection),
      route,
    ]),
  );
  const reviewDecisions = new Map(
    current.decisions.map((decision) => [
      routeKey(decision.source, decision.selection),
      decision,
    ]),
  );
  for (const conflict of current.conflicts) {
    if (conflict.kind !== "inferred-schema") continue;
    const key = routeKey(conflict.source, conflict.selection);
    const route = routes.get(key);
    if (route?.destination.kind === "new-table-infer") {
      const capture = captures.find(
        (candidate) =>
          candidate.sourceKey === conflict.source &&
          candidate.selectionKey === conflict.selection,
      );
      const existingApplication =
        capture === undefined
          ? []
          : await engineOf(options.database).query(
              `SELECT import_id FROM ${APPLICATION_TABLE} WHERE capture_id = ? AND table_name = ? LIMIT 1`,
              [capture.captureId, conflict.schema.name],
            );
      const approvedRoute = {
        ...route,
        destination:
          existingApplication.length > 0
            ? { kind: "existing-table", table: conflict.schema.name }
            : { kind: "new-table", schema: conflict.schema },
        columns: conflict.schema.columns.map((column) => ({
          source: column.name,
          target: column.name,
          type: column.type,
        })),
      } satisfies ImportRecipe["routes"][number];
      routes.set(key, approvedRoute);
      reviewDecisions.set(key, {
        kind: "route",
        source: approvedRoute.source,
        selection: approvedRoute.selection,
        destination: approvedRoute.destination,
        columns: approvedRoute.columns,
      });
    }
  }
  for (const decision of options.decisions) {
    const key = routeKey(decision.source, decision.selection);
    if (decision.kind === "exclude") {
      routes.delete(key);
    } else {
      routes.set(key, {
        source: decision.source,
        selection: decision.selection,
        destination: decision.destination,
        columns: decision.columns,
      });
    }
    reviewDecisions.set(key, decision);
  }
  const recipe: ImportRecipe = { version: 1, routes: [...routes.values()] };
  const included = captures.filter((capture) =>
    routes.has(routeKey(capture.sourceKey, capture.selectionKey)),
  );
  const conflicts = await evaluateConflicts(
    options.database,
    options.prepared,
    included,
    recipe,
  );
  return updatePreparedPlan({
    prepared: options.prepared,
    recipe,
    conflicts,
    ready: conflicts.length === 0,
    decisions: [...reviewDecisions.values()],
    ...(baseline === undefined
      ? {}
      : {
          baselineRevision: baseline.inspection.revision,
          baselineSchemaFingerprint: baseline.schemaFingerprint,
        }),
  });
}
