import { engineOf, inspectDatabase } from "../database.js";
import { databaseError } from "../errors.js";
import { APPLICATION_TABLE } from "../metadata.js";
import { readPreparedReviewSnapshot, updatePreparedPlan } from "../prepared.js";
import type { PreparedReview, preparedRef } from "../prepared.js";
import { readSchemaFingerprint } from "../records.js";
import { validateImportProfile } from "../validators.js";
import {
  evaluateConflicts,
  preparedCapturesFromEngine,
  routeColumns,
  routeKey,
  type PreparedCapture,
} from "./planning.js";
import type {
  ImportDecision,
  ImportBatchPage,
  ImportBatchRef,
  ReadyImportBatchRef,
  ImportReviewOutcome,
  ImportProfile,
  PrepareImportOptions,
} from "./types.js";

import { inspectUpdatedImport, assertImportReviewPage } from "./inspection.js";

interface ReviewSnapshot {
  readonly review: PreparedReview;
  readonly captures: readonly PreparedCapture[];
}

async function preparedReviewSnapshot(
  prepared: PrepareImportOptions["prepared"],
): Promise<ReviewSnapshot> {
  return readPreparedReviewSnapshot(prepared, async (review, transaction) => ({
    review,
    captures: await preparedCapturesFromEngine(transaction),
  }));
}

async function resolveReviewedImport(options: {
  readonly database: PrepareImportOptions["database"];
  readonly prepared: PrepareImportOptions["prepared"];
  readonly decisions: readonly ImportDecision[];
  readonly rebase?: boolean | undefined;
  readonly snapshot: ReviewSnapshot;
}): Promise<
  ReturnType<typeof preparedRef> extends Promise<infer T> ? T : never
> {
  if (options.prepared.databaseId !== options.database.id) {
    throw databaseError(
      "DB_IMPORT_DATABASE_MISMATCH",
      "This import batch belongs to a different database.",
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
  const current = options.snapshot.review;
  const captures = options.snapshot.captures;
  const routes = new Map(
    current.profile.routes.map((route) => [
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
  const knownSelections = new Set([
    ...routes.keys(),
    ...captures.map((capture) =>
      routeKey(capture.sourceKey, capture.selectionKey),
    ),
  ]);
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
        columns:
          capture === undefined ? route.columns : routeColumns(route, capture),
      } satisfies ImportProfile["routes"][number];
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
      if (!knownSelections.has(key)) {
        throw databaseError(
          "DB_IMPORT_DECISION_NOT_FOUND",
          `The exclusion for source "${decision.source}" selection "${decision.selection}" does not match a captured selection or an existing profile route. Check the source alias and selection key.`,
          { source: decision.source, selection: decision.selection },
        );
      }
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
  const profile: ImportProfile = { version: 1, routes: [...routes.values()] };
  validateImportProfile(profile);
  const excluded = new Set(
    [...reviewDecisions.values()].flatMap((decision) =>
      decision.kind === "exclude"
        ? [routeKey(decision.source, decision.selection)]
        : [],
    ),
  );
  const included = captures.filter(
    (capture) =>
      !excluded.has(routeKey(capture.sourceKey, capture.selectionKey)),
  );
  const conflicts = await evaluateConflicts(
    options.database,
    options.prepared,
    included,
    profile,
  );
  return updatePreparedPlan({
    prepared: options.prepared,
    profile,
    conflicts,
    ready: conflicts.length === 0,
    decisions: [...reviewDecisions.values()],
    expectedReviewFingerprint: current.prepared.reviewFingerprint,
    ...(baseline === undefined
      ? {}
      : {
          baselineRevision: baseline.inspection.revision,
          baselineSchemaFingerprint: baseline.schemaFingerprint,
        }),
  });
}

interface ResolveImportOptions {
  readonly database: PrepareImportOptions["database"];
  readonly prepared: PrepareImportOptions["prepared"];
  readonly decisions: readonly ImportDecision[];
  readonly rebase?: boolean | undefined;
  readonly reviewPage?: ImportBatchPage | undefined;
}

export function resolveImport(
  options: ResolveImportOptions & { readonly reviewPage: ImportBatchPage },
): Promise<ImportReviewOutcome>;
export function resolveImport(
  options: ResolveImportOptions,
): Promise<ImportBatchRef | ReadyImportBatchRef>;
export async function resolveImport(
  options: ResolveImportOptions,
): Promise<ImportReviewOutcome | ImportBatchRef | ReadyImportBatchRef> {
  if (options.reviewPage !== undefined)
    assertImportReviewPage(options.reviewPage);
  const prepared = await resolveReviewedImport({
    ...options,
    snapshot: await preparedReviewSnapshot(options.prepared),
  });
  if (options.reviewPage === undefined) return prepared;
  return inspectUpdatedImport({
    database: options.database,
    prepared: options.prepared,
    expected: prepared,
    page: options.reviewPage,
  });
}

export async function replaceImportProfile(options: {
  readonly database: PrepareImportOptions["database"];
  readonly prepared: PrepareImportOptions["prepared"];
  readonly profile: ImportProfile;
  readonly rebase?: boolean | undefined;
}): ReturnType<typeof resolveImport> {
  validateImportProfile(options.profile);
  const snapshot = await preparedReviewSnapshot(options.prepared);
  const current = snapshot.review;
  const captures = snapshot.captures;
  const replacementKeys = new Set(
    options.profile.routes.map((route) =>
      routeKey(route.source, route.selection),
    ),
  );
  const existing = new Map(
    [
      ...current.profile.routes,
      ...captures.map((capture) => ({
        source: capture.sourceKey,
        selection: capture.selectionKey,
      })),
    ].map((route) => [
      routeKey(route.source, route.selection),
      { source: route.source, selection: route.selection },
    ]),
  );
  return resolveReviewedImport({
    database: options.database,
    prepared: options.prepared,
    decisions: [
      ...options.profile.routes.map((route): ImportDecision => ({
        kind: "route",
        ...route,
      })),
      ...[...existing.values()]
        .filter(
          ({ source, selection }) =>
            !replacementKeys.has(routeKey(source, selection)),
        )
        .map(({ source, selection }): ImportDecision => ({
          kind: "exclude",
          source,
          selection,
          reason: "Omitted from the replacement profile.",
        })),
    ],
    rebase: options.rebase,
    snapshot,
  });
}
