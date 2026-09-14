import {
  applyImport,
  applySchema,
  checkpointDatabaseWrite,
  createWorkbookImportSource,
  draftImportProfile,
  identifierKey,
  inspectDatabase,
  inspectImport,
  listBatches,
  parseDatabaseSchema,
  planSchema,
  prepareImport,
  recordBatch,
  resolveImport,
  type ColumnDefinition,
  type Database,
  type BatchContext,
  type BatchRecord,
  type ImportCell,
  type ImportConflict,
  type ImportDecision,
  type ImportInspection,
  type ImportBatch,
  type ImportRouteInspection,
  type SchemaPlan,
} from "@consultchimps/db";
import {
  configureBrowserDatabaseRuntime,
  type BrowserDatabaseRuntime,
} from "@consultchimps/db/browser";
import {
  ConsultChimpsError,
  isConsultChimpsError,
  throwIfAborted,
  type OperationProgress,
} from "@consultchimps/core";

import { basePath } from "@/lib/shared";
import {
  BrowserBlobSource,
  BrowserOpfsFile,
  browserScratchFactory,
  createBrowserExportName,
  removeOpfsFile,
  withBrowserExportLease,
} from "@/lib/workspace-files";
import {
  workspaceSourceDescription,
  workspaceSourceKey,
} from "@/lib/workspace-source";
import {
  closeTrackedResources,
  replaceActiveWorkspace,
} from "@/lib/workspace-replacement";
import { workspaceWorkingCopyName } from "@/lib/workspace-naming";
import type {
  WorkspaceCommand,
  WorkspaceDeliveryContext,
  WorkspaceDeliveryPage,
  WorkspaceDeliverySummary,
  WorkspaceEvent,
  WorkspaceImportColumn,
  WorkspaceImportRegion,
  WorkspaceImportResult,
  WorkspacePreparedImport,
  WorkspacePreviewPage,
  WorkspaceProgress,
  WorkspaceRouteDecision,
  WorkspaceSchemaApplication,
  WorkspaceSchemaPlan,
  WorkspaceSummary,
} from "@/lib/workspace-protocol";
import {
  closeAndRetainFailures,
  type CleanupFailure,
  importCleanupError,
  RetryableCleanupOwners,
  savedPlanCleanupError,
  stagedPrivatePlanCleanup,
} from "./workspace-import-cleanup";

const scope = self as unknown as {
  addEventListener(
    type: "message",
    listener: (event: MessageEvent<WorkspaceCommand>) => void,
  ): void;
  postMessage(message: WorkspaceEvent): void;
};

interface OpenWorkspace {
  readonly database: Database;
  readonly workingCopyName: string;
}

interface HeldImport {
  readonly prepared: ImportBatch;
}

let browserRuntime: Promise<BrowserDatabaseRuntime> | null = null;
let workspace: OpenWorkspace | null = null;
const schemaPlans = new Map<string, SchemaPlan>();
const imports = new Map<string, HeldImport>();
const controllers = new Map<number, AbortController>();
const sourceCleanupOwners = new RetryableCleanupOwners();
const privatePlanCleanupOwners = new RetryableCleanupOwners();
const savedPlanCleanupOwners = new RetryableCleanupOwners();
const IMPORT_ROUTE_PAGE_SIZE = 50;

function runtime(): Promise<BrowserDatabaseRuntime> {
  const configured =
    browserRuntime ??
    configureBrowserDatabaseRuntime({
      sqlite: {
        wasmUrl: `${basePath}/database-wasm/sqlite3.wasm`,
        directory: "/consultchimps-sqlite",
        initialCapacity: 16,
      },
      duckdb: {
        wasmUrl: `${basePath}/database-wasm/duckdb-eh.wasm`,
        workerUrl: `${basePath}/database-wasm/duckdb-browser-eh.worker.js`,
      },
      opfsDirectory: "consultchimps-databases",
    });
  browserRuntime = configured;
  return configured;
}

function current(): OpenWorkspace {
  if (workspace === null) {
    throw new Error("Create or open a database before using this operation");
  }
  return workspace;
}

function boundedNumber(value: bigint, label: string): number {
  const result = Number(value);
  if (!Number.isSafeInteger(result)) {
    throw new Error(`The ${label} is too large to display safely`);
  }
  return result;
}

async function summaryOf(open: OpenWorkspace): Promise<WorkspaceSummary> {
  const inspection = await inspectDatabase({ database: open.database });
  return {
    databaseId: inspection.id,
    format: inspection.format,
    formatVersion: inspection.formatVersion,
    workingCopyName: open.workingCopyName,
    tables: inspection.tables.map((table) => ({
      id: identifierKey(table.name),
      name: table.name,
      rowCount: boundedNumber(table.rowCount, `row count for ${table.name}`),
      columns: table.schema.columns.map((column) => ({
        name: column.name,
        type: column.type,
        nullable: column.nullable !== false,
      })),
    })),
    importCount: boundedNumber(inspection.completedImports, "import count"),
    deliveryCount: boundedNumber(
      inspection.recordedBatches,
      "recorded batch count",
    ),
  };
}

function progressOf(progress: OperationProgress): WorkspaceProgress {
  return {
    phase: progress.stage,
    completed: progress.completed,
    total: progress.total,
    message: progress.detail ?? progress.stage,
  };
}

function onProgress(id: number): (progress: OperationProgress) => void {
  return (progress) => {
    scope.postMessage({ type: "progress", id, progress: progressOf(progress) });
  };
}

function postError(id: number, error: unknown): void {
  const aborted =
    typeof error === "object" &&
    error !== null &&
    "name" in error &&
    error.name === "AbortError";
  scope.postMessage({
    type: "error",
    id,
    message: aborted
      ? "The operation was cancelled"
      : error instanceof Error
        ? error.message
        : "An unexpected problem occurred",
    ...(aborted
      ? { code: "OPERATION_ABORTED" }
      : isConsultChimpsError(error)
        ? { code: error.code }
        : {}),
  });
}

async function closeImports(): Promise<void> {
  await closeTrackedResources({
    resources: imports,
    close: ({ prepared }) => prepared.close(),
  });
}

async function cleanupFailureOf(
  cleanup: () => Promise<void>,
): Promise<CleanupFailure | undefined> {
  try {
    await cleanup();
    return undefined;
  } catch (error) {
    return { error };
  }
}

async function retryImportPreparationCleanup(): Promise<void> {
  const [sources, plans] = await Promise.all([
    cleanupFailureOf(() => sourceCleanupOwners.close()),
    cleanupFailureOf(() => privatePlanCleanupOwners.close()),
  ]);
  if (sources !== undefined || plans !== undefined) {
    throw importCleanupError({
      ...(sources === undefined ? {} : { sourceCleanupFailure: sources }),
      ...(plans === undefined ? {} : { planCleanupFailure: plans }),
      preparationCompleted: false,
    });
  }
}

async function replaceWorkspace(
  next: OpenWorkspace,
  signal: AbortSignal,
): Promise<WorkspaceSummary> {
  const previous = workspace;
  return replaceActiveWorkspace({
    previous,
    next,
    async prepare(open) {
      signal.throwIfAborted();
      const summary = await summaryOf(open);
      signal.throwIfAborted();
      return summary;
    },
    closeDependencies: closeImports,
    close: (open) => open.database.close(),
    activate() {
      schemaPlans.clear();
      workspace = next;
    },
  });
}

function schemaPlanDto(id: string, plan: SchemaPlan): WorkspaceSchemaPlan {
  const columnDto = (
    value: SchemaPlan["creates"][number]["columns"][number],
  ) => ({
    name: value.name,
    type: value.type,
    nullable: value.nullable !== false,
    ...(value.precision === undefined ? {} : { precision: value.precision }),
    ...(value.scale === undefined ? {} : { scale: value.scale }),
  });
  return {
    id,
    changes: [
      ...plan.creates.map((table) => ({
        kind: "create-table" as const,
        table: {
          name: table.name,
          recordId: {
            prefix: table.recordId.prefix,
            separator: table.recordId.separator ?? "-",
            padding: table.recordId.padding,
          },
          columns: table.columns.map(columnDto),
          foreignKeys: (table.foreignKeys ?? []).map((foreignKey) => ({
            column: foreignKey.column,
            referencesTable: foreignKey.referencesTable,
          })),
        },
      })),
      ...plan.adds.flatMap((addition) =>
        addition.columns.map((column) => ({
          kind: "add-column" as const,
          table: addition.table,
          column: columnDto(column),
        })),
      ),
    ],
    conflicts: plan.conflicts.map((conflict) => ({
      table: conflict.table,
      column: "column" in conflict ? conflict.column : null,
      message:
        conflict.kind === "column-type"
          ? `Type conflict: column ${conflict.column} is ${conflict.existing.type}, but the proposed schema uses ${conflict.proposed.type}`
          : conflict.kind === "table-definition"
            ? conflict.message
            : `Column ${conflict.column} is required and cannot be added to existing rows without a value`,
    })),
    ready: plan.state === "ready",
  };
}

function regionId(source: string, selection: string): string {
  return JSON.stringify([source, selection]);
}

function conflictText(conflict: ImportConflict): string {
  switch (conflict.kind) {
    case "missing-destination":
      return "Choose a destination table";
    case "source-selection-not-found":
      return `Source ${workspaceSourceDescription(conflict.source)} selection ${conflict.selection} was not captured; choose an available source and selection`;
    case "missing-column":
      return `Destination column ${conflict.column} does not exist`;
    case "source-column-not-found":
      return `Source column ${conflict.column} does not exist in the captured rows`;
    case "required-column-unmapped":
      return `Required destination column ${conflict.target} needs a source column`;
    case "required-value":
      return `Row ${String(conflict.sourceRow)} needs a value for required column ${conflict.target}`;
    case "invalid-value":
      return `Row ${String(conflict.sourceRow)} column ${conflict.column} cannot be stored as ${conflict.expected} in ${conflict.target}`;
    case "foreign-key-value-not-found":
      return `Row ${String(conflict.sourceRow)} column ${conflict.column} does not match a record in ${conflict.referencesTable}`;
    case "incompatible-column":
      return `Column ${conflict.target} needs the ${conflict.expected} type`;
    case "decimal-capacity":
      return `Column ${conflict.target} needs decimal(${conflict.requiredPrecision}, ${conflict.requiredScale}), but the destination allows decimal(${conflict.targetPrecision}, ${conflict.targetScale})`;
    case "conflicting-application-mapping":
      return `The same captured data has a different column mapping for ${conflict.table}; use one consistent mapping or choose another table`;
    case "inferred-schema":
      return "Review the inferred columns and approve the destination";
    case "table-exists":
      return `Table ${conflict.table} already exists`;
    case "table-not-found":
      return `Table ${conflict.table} does not exist`;
  }
}

function routeConflicts(
  route: ImportRouteInspection,
  conflicts: readonly ImportConflict[],
): readonly ImportConflict[] {
  return conflicts.filter(
    (conflict) =>
      (!("source" in conflict) || conflict.source === route.source) &&
      (!("selection" in conflict) || conflict.selection === route.selection),
  );
}

function columnsFor(
  route: ImportRouteInspection,
  conflicts: readonly ImportConflict[],
): WorkspaceImportColumn[] {
  return route.inferredColumns.map((column) => {
    const mapped = route.columns.find((entry) => entry.source === column.name);
    const destination = mapped?.target ?? column.name;
    const target = route.destinationColumns.find(
      (candidate) =>
        identifierKey(candidate.name) === identifierKey(destination),
    );
    const conflict = conflicts.find(
      (candidate) =>
        "source" in candidate &&
        candidate.source === route.source &&
        candidate.selection === route.selection &&
        "column" in candidate &&
        candidate.column === column.name,
    );
    return {
      source: column.name,
      destination,
      inferredType: column.type,
      destinationType: mapped?.type ?? target?.type ?? column.type,
      compatible: conflict === undefined,
      message: conflict === undefined ? null : conflictText(conflict),
    };
  });
}

function importDto(
  inspection: ImportInspection,
  routeCursor: string | null,
): WorkspacePreparedImport {
  const regions = inspection.routes.map((route): WorkspaceImportRegion => {
    const conflicts = routeConflicts(route, inspection.conflicts);
    return {
      id: regionId(route.source, route.selection),
      sourceId: route.source,
      fileName: route.displayName,
      label: route.label,
      rowCount: boundedNumber(route.rowCount, "captured row count"),
      columns: columnsFor(route, conflicts),
      route:
        route.destination?.kind === "existing-table"
          ? { kind: "append", table: route.destination.table }
          : route.destination?.kind === "new-table"
            ? { kind: "create", table: route.destination.schema.name }
            : {
                kind: "unresolved",
                suggestedTable: route.suggestedDestination.name,
              },
      conflicts: conflicts.map(conflictText),
    };
  });
  const includedRoutes = inspection.routes.filter(
    (route) => route.destination !== null,
  );
  const routeCount = boundedNumber(inspection.routeCount, "batch route count");
  const application =
    inspection.application.state === "applied" ? "applied" : "pending";
  return {
    id: inspection.prepared.id,
    reviewFingerprint: inspection.prepared.reviewFingerprint,
    state: inspection.prepared.state,
    application,
    duplicateOf:
      routeCount <= inspection.routes.length &&
      includedRoutes.length > 0 &&
      includedRoutes.every(
        (route) => route.applicationState === "already-applied",
      )
        ? "existing-application"
        : null,
    captureIds:
      inspection.application.state === "applied"
        ? inspection.application.captureIds
        : inspection.captureIds,
    regions,
    routeCount,
    routeCursor,
    nextRouteCursor: inspection.nextRouteCursor ?? null,
    totalRows: boundedNumber(inspection.reviewRows, "review row count"),
    warningCount: inspection.conflicts.length,
  };
}

function requireReviewFingerprint(
  inspection: ImportInspection,
  expected: string,
): void {
  if (inspection.prepared.reviewFingerprint !== expected) {
    throw new ConsultChimpsError(
      "DB_STALE_IMPORT_PLAN",
      "The saved batch review changed after it was displayed. Review the latest batch before continuing.",
      { details: { batchId: inspection.prepared.id } },
    );
  }
}

function deliveryContext(input: WorkspaceDeliveryContext): BatchContext {
  const scopeValue: BatchContext["scope"] =
    input.coverage === "full"
      ? { kind: "full" }
      : input.coverage === "partial"
        ? {
            kind: "partial",
            description: input.note.trim() || "Partial batch",
          }
        : { kind: "unknown" };
  return {
    label:
      [input.vendor, input.entity, input.phase]
        .map((value) => value.trim())
        .filter(Boolean)
        .join(" | ") || input.requestId,
    scope: scopeValue,
    ...(input.effectiveDate === null
      ? {}
      : { effectiveDate: input.effectiveDate }),
    ...(input.receivedDate === null
      ? {}
      : { receivedDate: input.receivedDate }),
    attributes: {
      vendor: input.vendor,
      entity: input.entity,
      phase: input.phase,
      note: input.note,
    },
  };
}

function deliveryDto(delivery: BatchRecord): WorkspaceDeliverySummary {
  const attributes = delivery.context.attributes ?? {};
  return {
    id: delivery.id,
    requestId: delivery.requestId,
    label: delivery.context.label,
    vendor:
      typeof attributes["vendor"] === "string" ? attributes["vendor"] : "",
    entity:
      typeof attributes["entity"] === "string" ? attributes["entity"] : "",
    phase: typeof attributes["phase"] === "string" ? attributes["phase"] : "",
    scope: delivery.context.scope,
    effectiveDate: delivery.context.effectiveDate ?? null,
    receivedDate: delivery.context.receivedDate ?? null,
    captureIds: delivery.captureIds,
    reusedCapture: delivery.reusedCaptureIds.length > 0,
  };
}

function cellValue(
  cell: ImportCell | undefined,
): boolean | null | number | string {
  if (cell === undefined || cell.kind === "blank") return null;
  if (cell.kind === "boolean") return cell.value;
  if (cell.kind === "string") return cell.value;
  if (cell.kind === "number") return cell.raw;
  if (cell.kind === "date") return cell.iso;
  if (cell.kind === "error") return cell.error;
  return cell.cached.kind === "missing"
    ? "Formula has no cached value"
    : cellValue(cell.cached);
}

function decisionFor(
  decision: WorkspaceRouteDecision,
  route: ImportRouteInspection,
): ImportDecision {
  const columns = decision.columns
    .filter(
      (column): column is typeof column & { readonly destination: string } =>
        column.destination !== null,
    )
    .map((column) => ({
      source: column.source,
      target: column.destination,
      type: column.type as ColumnDefinition["type"],
    }));
  if (decision.route.kind === "append") {
    return {
      kind: "route",
      source: route.source,
      selection: route.selection,
      destination: { kind: "existing-table", table: decision.route.table },
      columns,
    };
  }
  const schemaColumns = columns.map((column): ColumnDefinition => {
    const inferred = route.inferredColumns.find(
      (candidate) => candidate.name === column.source,
    );
    const destination = route.destinationColumns.find(
      (candidate) =>
        identifierKey(candidate.name) === identifierKey(column.target),
    );
    if (destination?.type === column.type) {
      return { ...destination, name: column.target };
    }
    if (column.type === "decimal") {
      return {
        name: column.target,
        type: "decimal",
        nullable: inferred?.nullable,
        precision: inferred?.type === "decimal" ? inferred.precision : 38,
        scale: inferred?.type === "decimal" ? inferred.scale : 10,
      };
    }
    return {
      name: column.target,
      type: column.type,
      ...(inferred?.nullable === undefined
        ? {}
        : { nullable: inferred.nullable }),
    };
  });
  return {
    kind: "route",
    source: route.source,
    selection: route.selection,
    destination: {
      kind: "new-table",
      schema: {
        ...(route.destination?.kind === "new-table"
          ? route.destination.schema
          : {
              name: route.suggestedDestination.name,
              recordId: route.suggestedDestination.recordId,
              columns: route.destinationColumns,
            }),
        name: decision.route.table,
        columns: schemaColumns,
      },
    },
    columns,
  };
}

async function handleCreate(
  id: number,
  command: Extract<WorkspaceCommand, { readonly type: "create" }>,
  signal: AbortSignal,
): Promise<void> {
  const created = await (
    await runtime()
  ).createDatabase({
    name: command.name,
    format: command.format,
    ...(command.schema === undefined
      ? {}
      : { schema: parseDatabaseSchema(command.schema) }),
    ...(command.overwrite === undefined
      ? {}
      : { overwrite: command.overwrite }),
    signal,
  });
  const next = {
    database: created.database,
    workingCopyName: command.name,
  };
  const summary = await replaceWorkspace(next, signal);
  scope.postMessage({ type: "ready", id, summary });
}

async function handleOpen(
  id: number,
  command: Extract<WorkspaceCommand, { readonly type: "open" }>,
  signal: AbortSignal,
): Promise<void> {
  const name =
    command.name === undefined
      ? workspaceWorkingCopyName(command.file.name)
      : command.name;
  const database = await (
    await runtime()
  ).importDatabase({
    name,
    source: new BrowserBlobSource(command.file.name, command.file),
    ...(command.overwrite === undefined
      ? {}
      : { overwrite: command.overwrite }),
    signal,
    onProgress: onProgress(id),
  });
  const next = { database, workingCopyName: name };
  const summary = await replaceWorkspace(next, signal);
  scope.postMessage({ type: "ready", id, summary });
}

async function handleReopen(
  id: number,
  command: Extract<WorkspaceCommand, { readonly type: "reopen" }>,
  signal: AbortSignal,
): Promise<void> {
  signal.throwIfAborted();
  const database = await (
    await runtime()
  ).openDatabase({
    name: command.name,
    ...(command.readonly === undefined ? {} : { readonly: command.readonly }),
  });
  const next = { database, workingCopyName: command.name };
  const summary = await replaceWorkspace(next, signal);
  scope.postMessage({ type: "ready", id, summary });
}

async function handleSchema(
  id: number,
  command: Extract<
    WorkspaceCommand,
    { readonly type: "applySchema" | "planSchema" }
  >,
  signal: AbortSignal,
): Promise<void> {
  if (command.type === "planSchema") {
    const plan = await planSchema({
      database: current().database,
      schema: parseDatabaseSchema(command.schema),
    });
    throwIfAborted(signal, "db.schema.plan");
    const planId = globalThis.crypto.randomUUID();
    schemaPlans.clear();
    schemaPlans.set(planId, plan);
    scope.postMessage({
      type: "schemaPlanned",
      id,
      plan: schemaPlanDto(planId, plan),
    });
    return;
  }
  const plan = schemaPlans.get(command.planId);
  if (plan === undefined) {
    throw new Error("Review the schema again before applying it");
  }
  const result = await applySchema({
    database: current().database,
    plan,
    signal,
  });
  schemaPlans.delete(command.planId);
  const checkpointed = await checkpointDatabaseWrite({
    database: current().database,
    result,
  });
  const summary = await refreshSummary();
  const checkpoint: WorkspaceSchemaApplication["checkpoint"] =
    checkpointed.checkpoint.state === "checkpoint-completed"
      ? { state: "saved" }
      : {
          state: "failed",
          code: "DB_BROWSER_SCHEMA_PERSISTENCE_REQUIRED",
          message:
            "The schema changes were applied, but the browser could not finish saving the database. Keep this tab open and export the database again to retry saving it. Do not apply the schema changes again",
        };
  scope.postMessage({
    type: "schemaApplied",
    id,
    result: { summary, checkpoint },
  });
}

async function listSavedImports(id: number): Promise<void> {
  await closeImports();
  const retainedCleanupFailure = await cleanupFailureOf(() =>
    savedPlanCleanupOwners.close(),
  );
  if (retainedCleanupFailure !== undefined) {
    throw savedPlanCleanupError({
      operationFailures: [],
      cleanupFailures: [retainedCleanupFailure.error],
    });
  }
  const listing = await (
    await runtime()
  ).listImportBatches({ database: current().database });
  const ignoredPlans = [...listing.ignored];
  const openedImports: {
    readonly held: HeldImport;
    readonly plan: WorkspacePreparedImport;
  }[] = [];
  const operationFailures: unknown[] = [];
  const cleanupFailures: unknown[] = [];
  for (const entry of listing.imports) {
    if (entry.databaseId !== current().database.id) continue;
    let prepared: ImportBatch | undefined;
    try {
      prepared = await (await runtime()).openImportBatch({ name: entry.name });
      const inspection = await inspectImport({
        database: current().database,
        prepared,
        page: { limit: 1 },
      });
      const held = { prepared } satisfies HeldImport;
      openedImports.push({ held, plan: importDto(inspection, null) });
      prepared = undefined;
    } catch (error) {
      const ignoredPlan = isConsultChimpsError(error)
        ? { name: entry.name, code: error.code, message: error.message }
        : {
            name: entry.name,
            code: "DB_BROWSER_PREPARED_IMPORT_UNREADABLE",
            message:
              "This saved batch could not be reopened. Reload the database tool, then prepare the original sources again or restore a verified batch copy if it remains unavailable.",
          };
      if (prepared === undefined) {
        ignoredPlans.push(ignoredPlan);
        continue;
      }
      const opened = prepared;
      const cleanupOwner = { close: () => opened.close() };
      const cleanupFailure = await cleanupFailureOf(() => cleanupOwner.close());
      if (cleanupFailure === undefined) {
        ignoredPlans.push(ignoredPlan);
        continue;
      }
      savedPlanCleanupOwners.retain(cleanupOwner);
      operationFailures.push(error);
      cleanupFailures.push(cleanupFailure.error);
    }
  }
  if (cleanupFailures.length > 0) {
    const openedCleanupFailure = await cleanupFailureOf(() =>
      closeAndRetainFailures(
        openedImports.map((entry) => entry.held.prepared),
        savedPlanCleanupOwners,
      ),
    );
    if (openedCleanupFailure !== undefined) {
      cleanupFailures.push(openedCleanupFailure.error);
    }
    throw savedPlanCleanupError({ operationFailures, cleanupFailures });
  }
  for (const entry of openedImports) imports.set(entry.plan.id, entry.held);
  scope.postMessage({
    type: "importsListed",
    id,
    plans: openedImports.map((entry) => entry.plan),
    ignoredPlans,
  });
}

async function prepareSources(
  id: number,
  command: Extract<WorkspaceCommand, { readonly type: "prepareImport" }>,
  signal: AbortSignal,
): Promise<void> {
  await retryImportPreparationCleanup();
  const workbookSources: Awaited<
    ReturnType<typeof createWorkbookImportSource>
  >[] = [];
  const sourceInputs = command.sources.map((source, ordinal) => ({
    source,
    key: workspaceSourceKey(source, ordinal),
  }));
  let privatePlan:
    { readonly name: string; readonly prepared: ImportBatch } | undefined;
  let completed:
    | { readonly held: HeldImport; readonly plan: WorkspacePreparedImport }
    | undefined;
  let operationFailure: CleanupFailure | undefined;
  try {
    for (const { key, source } of sourceInputs) {
      workbookSources.push(
        await createWorkbookImportSource({
          key,
          bytes: new BrowserBlobSource(source.file.name, source.file),
          scratch: browserScratchFactory,
          signal,
          onProgress: onProgress(id),
        }),
      );
    }
    const sources = workbookSources.map((source) => source.source);
    const profile = await draftImportProfile({
      sources,
      naming: { kind: "selection-label" },
    });
    const databaseInspection = await inspectDatabase({
      database: current().database,
    });
    const privatePlanName = `.consultchimps-import-${globalThis.crypto.randomUUID()}.sqlite`;
    const prepared = await (
      await runtime()
    ).createImportBatch({
      name: privatePlanName,
      database: current().database,
      profile,
      baselineRevision: databaseInspection.revision,
      signal,
    });
    privatePlan = { name: privatePlanName, prepared };
    const outcome = await prepareImport({
      database: current().database,
      prepared,
      sources,
      profile,
      reviewPage: { limit: 1 },
      signal,
      onProgress: onProgress(id),
    });
    let inspection = outcome.inspection;
    const automatic = inspection.conflicts.flatMap(
      (conflict): ImportDecision[] => {
        if (conflict.kind !== "inferred-schema") return [];
        const existing = databaseInspection.tables.find(
          (table) =>
            identifierKey(table.name) === identifierKey(conflict.schema.name),
        );
        if (existing === undefined) return [];
        const matches = conflict.schema.columns.every((column) =>
          existing.schema.columns.some(
            (target) =>
              identifierKey(target.name) === identifierKey(column.name) &&
              target.type === column.type,
          ),
        );
        if (!matches) return [];
        return [
          {
            kind: "route",
            source: conflict.source,
            selection: conflict.selection,
            destination: { kind: "existing-table", table: existing.name },
            columns: conflict.schema.columns.map((column) => ({
              source: column.name,
              target: column.name,
              type: column.type,
            })),
          },
        ];
      },
    );
    if (automatic.length > 0) {
      inspection = (
        await resolveImport({
          database: current().database,
          prepared,
          decisions: automatic,
          reviewPage: { limit: 1 },
        })
      ).inspection;
    }
    const held = { prepared } satisfies HeldImport;
    completed = { held, plan: importDto(inspection, null) };
  } catch (error) {
    operationFailure = { error };
  }

  const sourceCleanupFailure = await cleanupFailureOf(() =>
    closeAndRetainFailures(workbookSources, sourceCleanupOwners),
  );
  if (operationFailure !== undefined || sourceCleanupFailure !== undefined) {
    let planCleanupFailure: CleanupFailure | undefined;
    if (privatePlan !== undefined) {
      const plan = privatePlan;
      const cleanupOwner = stagedPrivatePlanCleanup({
        close: () => plan.prepared.close(),
        discard: async () =>
          (await runtime()).discardImportBatch({ name: plan.name }),
      });
      planCleanupFailure = await cleanupFailureOf(() => cleanupOwner.close());
      if (planCleanupFailure !== undefined) {
        privatePlanCleanupOwners.retain(cleanupOwner);
      }
    }
    if (
      sourceCleanupFailure !== undefined ||
      planCleanupFailure !== undefined
    ) {
      throw importCleanupError({
        ...(operationFailure === undefined ? {} : { operationFailure }),
        ...(sourceCleanupFailure === undefined ? {} : { sourceCleanupFailure }),
        ...(planCleanupFailure === undefined ? {} : { planCleanupFailure }),
        preparationCompleted: operationFailure === undefined,
      });
    }
    throw operationFailure!.error;
  }

  if (completed === undefined) {
    throw new Error("The prepared batch result is unavailable");
  }
  imports.set(completed.plan.id, completed.held);
  privatePlan = undefined;
  scope.postMessage({ type: "importPrepared", id, plan: completed.plan });
}

function parseRegionId(id: string): readonly [string, string] {
  let value: unknown;
  try {
    value = JSON.parse(id);
  } catch {
    throw new Error("The selected batch region is no longer available");
  }
  if (
    !Array.isArray(value) ||
    value.length !== 2 ||
    typeof value[0] !== "string" ||
    typeof value[1] !== "string"
  ) {
    throw new Error("The selected batch region is no longer available");
  }
  return [value[0], value[1]];
}

async function inspectImportPage(
  id: number,
  command: Extract<WorkspaceCommand, { readonly type: "inspectImport" }>,
): Promise<void> {
  const held = imports.get(command.planId);
  if (held === undefined) {
    throw new Error("Prepare the batch again before reviewing it");
  }
  const inspection = await inspectImport({
    database: current().database,
    prepared: held.prepared,
    page: { limit: 1 },
    routePage: {
      limit: IMPORT_ROUTE_PAGE_SIZE,
      ...(command.routeCursor === null ? {} : { cursor: command.routeCursor }),
    },
  });
  scope.postMessage({
    type: "importInspected",
    id,
    plan: importDto(inspection, command.routeCursor),
  });
}

async function previewImport(
  id: number,
  command: Extract<WorkspaceCommand, { readonly type: "previewImport" }>,
): Promise<void> {
  const held = imports.get(command.planId);
  if (held === undefined) {
    throw new Error("Prepare the batch again before previewing it");
  }
  const [source, selection] = parseRegionId(command.regionId);
  const inspection = await inspectImport({
    database: current().database,
    prepared: held.prepared,
    page: {
      limit: command.limit,
      ...(command.cursor === null ? {} : { cursor: command.cursor }),
      source,
      selection,
    },
    routePage: {
      limit: IMPORT_ROUTE_PAGE_SIZE,
      ...(command.routeCursor === null ? {} : { cursor: command.routeCursor }),
    },
  });
  const route = inspection.routes.find(
    (candidate) =>
      candidate.source === source && candidate.selection === selection,
  );
  if (route === undefined) {
    throw new Error("The selected batch region is no longer available");
  }
  const columns = route.inferredColumns.map((column) => column.name);
  const page: WorkspacePreviewPage = {
    planId: command.planId,
    regionId: command.regionId,
    columns,
    rows: inspection.examples.map((example) =>
      columns.map((column) => cellValue(example.values[column])),
    ),
    cursor: command.cursor,
    nextCursor: inspection.nextCursor ?? null,
  };
  scope.postMessage({ type: "importPreview", id, page });
}

async function updateImport(
  id: number,
  command: Extract<WorkspaceCommand, { readonly type: "resolveImport" }>,
): Promise<void> {
  const held = imports.get(command.planId);
  if (held === undefined) {
    throw new Error("Prepare the batch again before updating it");
  }
  const inspection = await inspectImport({
    database: current().database,
    prepared: held.prepared,
    page: { limit: 1 },
    routePage: {
      limit: IMPORT_ROUTE_PAGE_SIZE,
      ...(command.routeCursor === null ? {} : { cursor: command.routeCursor }),
    },
  });
  requireReviewFingerprint(inspection, command.reviewFingerprint);
  const decisions = command.decisions.map((decision) => {
    const route = inspection.routes.find(
      (candidate) =>
        regionId(candidate.source, candidate.selection) === decision.regionId,
    );
    if (route === undefined) {
      throw new Error("A batch decision refers to an unavailable region");
    }
    return decisionFor(decision, route);
  });
  const resolved = await resolveImport({
    database: current().database,
    prepared: held.prepared,
    decisions,
    reviewPage: { limit: 1 },
  });
  scope.postMessage({
    type: "importResolved",
    id,
    plan: importDto(resolved.inspection, null),
  });
}

function checkpointDto(
  checkpoint: Awaited<ReturnType<typeof checkpointDatabaseWrite>>["checkpoint"],
): WorkspaceImportResult["checkpoint"] {
  return checkpoint.state === "checkpoint-completed"
    ? { state: "saved" }
    : {
        state: "failed",
        code: checkpoint.code,
        message: checkpoint.message,
      };
}

async function refreshSummary(): Promise<WorkspaceImportResult["summary"]> {
  try {
    return { state: "updated", value: await summaryOf(current()) };
  } catch {
    return {
      state: "refresh-required",
      code: "DB_BROWSER_SUMMARY_REFRESH_REQUIRED",
      message:
        "The database operation completed, but the browser could not refresh its summary. Keep this tab open and reopen the database view before continuing",
    };
  }
}

async function applyImportBatch(
  id: number,
  command: Extract<WorkspaceCommand, { readonly type: "applyImport" }>,
  signal: AbortSignal,
): Promise<void> {
  const held = imports.get(command.planId);
  if (held === undefined) {
    throw new Error("Update the batch review before applying it");
  }
  const inspection = await inspectImport({
    database: current().database,
    prepared: held.prepared,
    page: { limit: 1 },
  });
  requireReviewFingerprint(inspection, command.reviewFingerprint);
  if (inspection.prepared.state !== "ready") {
    throw new Error("Update the batch review before applying it");
  }
  const result = await applyImport({
    database: current().database,
    prepared: held.prepared,
    approved: inspection.prepared,
    requestId: command.delivery.requestId,
    batchContext: deliveryContext(command.delivery),
    signal,
    onProgress: onProgress(id),
  });
  const checkpointed = await checkpointDatabaseWrite({
    database: current().database,
    result,
  });
  const summary = await refreshSummary();
  scope.postMessage({
    type: "importApplied",
    id,
    result: {
      importId: result.importIds[0] ?? command.delivery.requestId,
      receiptId: result.batchId ?? command.delivery.requestId,
      outcome:
        result.metrics.rowsImported === 0 && result.metrics.rowsReused > 0
          ? "duplicate"
          : "applied",
      appendedRows: result.metrics.rowsImported,
      skippedRows: result.metrics.rowsReused,
      unresolvedRows: 0,
      schemaChanges: result.metrics.tablesCreated,
      deliveriesRecorded: result.metrics.batchesRecorded,
      captureIds: result.captureIds,
      summary,
      checkpoint: checkpointDto(checkpointed.checkpoint),
    },
  });
}

async function recordPreparedDelivery(
  id: number,
  command: Extract<WorkspaceCommand, { readonly type: "recordDelivery" }>,
  signal: AbortSignal,
): Promise<void> {
  const held = imports.get(command.planId);
  if (held === undefined) {
    throw new Error("Update the batch review before recording it again");
  }
  signal.throwIfAborted();
  const inspection = await inspectImport({
    database: current().database,
    prepared: held.prepared,
    page: { limit: 1 },
  });
  requireReviewFingerprint(inspection, command.reviewFingerprint);
  if (inspection.prepared.state !== "ready") {
    throw new Error("Update the batch review before recording it again");
  }
  const captureIds =
    inspection.application.state === "applied"
      ? inspection.application.captureIds
      : inspection.captureIds;
  signal.throwIfAborted();
  const record = await recordBatch({
    database: current().database,
    captureIds: [...new Set(captureIds)],
    context: deliveryContext(command.delivery),
    requestId: command.delivery.requestId,
  });
  const checkpointed = await checkpointDatabaseWrite({
    database: current().database,
    result: record,
  });
  const summary = await refreshSummary();
  scope.postMessage({
    type: "importApplied",
    id,
    result: {
      importId: record.batch.id,
      receiptId: record.batch.id,
      outcome: "duplicate",
      appendedRows: 0,
      skippedRows: boundedNumber(inspection.reviewRows, "review row count"),
      unresolvedRows: 0,
      schemaChanges: 0,
      deliveriesRecorded: record.metrics.batchesRecorded,
      captureIds: record.batch.captureIds,
      summary,
      checkpoint: checkpointDto(checkpointed.checkpoint),
    },
  });
}

async function deliveryHistory(
  id: number,
  command: Extract<WorkspaceCommand, { readonly type: "listDeliveries" }>,
): Promise<void> {
  const page = await listBatches({
    database: current().database,
    limit: command.limit,
    ...(command.cursor === null ? {} : { cursor: command.cursor }),
  });
  const dto: WorkspaceDeliveryPage = {
    deliveries: page.batches.map(deliveryDto),
    nextCursor: page.nextCursor ?? null,
  };
  scope.postMessage({ type: "deliveries", id, page: dto });
}

async function exportWorkspace(
  id: number,
  command: Extract<WorkspaceCommand, { readonly type: "export" }>,
  signal: AbortSignal,
): Promise<void> {
  const extension = command.format === "duckdb" ? "duckdb" : "sqlite";
  const name = `${current().workingCopyName.replace(/\.[^.]+$/u, "")}.${extension}`;
  const destinationName = createBrowserExportName(extension);
  await withBrowserExportLease(destinationName, async () => {
    let destination: BrowserOpfsFile;
    try {
      destination = await BrowserOpfsFile.open(destinationName, true, false);
    } catch (error) {
      try {
        await removeOpfsFile(destinationName);
      } catch (cleanupError) {
        if (!(
          cleanupError instanceof DOMException &&
          cleanupError.name === "NotFoundError"
        )) {
          throw new AggregateError(
            [error, cleanupError],
            "The browser export could not finish cleaning its private storage",
          );
        }
      }
      throw error;
    }
    let completed = false;
    let destinationClosed = false;
    let operationError: unknown;
    try {
      await current().database.checkpoint();
      await (
        await runtime()
      ).exportDatabase({
        database: current().database,
        name,
        destination,
        format: command.format,
        overwrite: true,
        signal,
        onProgress: onProgress(id),
      });
      const file = await destination.file();
      await destination.close();
      destinationClosed = true;
      scope.postMessage({
        type: "exported",
        id,
        file,
        name,
        format: command.format,
      });
      completed = true;
    } catch (error) {
      operationError = error;
      throw error;
    } finally {
      const cleanupFailures: unknown[] = [];
      if (!destinationClosed) {
        try {
          await destination.close();
        } catch (error) {
          cleanupFailures.push(error);
        }
      }
      if (!completed) {
        try {
          await removeOpfsFile(destinationName);
        } catch (error) {
          if (!(
            error instanceof DOMException && error.name === "NotFoundError"
          )) {
            cleanupFailures.push(error);
          }
        }
      }
      if (cleanupFailures.length > 0) {
        throw new AggregateError(
          operationError === undefined
            ? cleanupFailures
            : [operationError, ...cleanupFailures],
          "The browser export could not finish cleaning its private storage",
        );
      }
    }
  });
}

async function handle(id: number, command: WorkspaceCommand): Promise<void> {
  const controller = new AbortController();
  controllers.set(id, controller);
  try {
    switch (command.type) {
      case "cancel":
        controllers.get(command.targetId)?.abort();
        scope.postMessage({ type: "closed", id });
        return;
      case "create":
        await handleCreate(id, command, controller.signal);
        return;
      case "open":
        await handleOpen(id, command, controller.signal);
        return;
      case "reopen":
        await handleReopen(id, command, controller.signal);
        return;
      case "planSchema":
      case "applySchema":
        await handleSchema(id, command, controller.signal);
        return;
      case "prepareImport":
        await prepareSources(id, command, controller.signal);
        return;
      case "listImports":
        await listSavedImports(id);
        return;
      case "inspectImport":
        await inspectImportPage(id, command);
        return;
      case "previewImport":
        await previewImport(id, command);
        return;
      case "resolveImport":
        await updateImport(id, command);
        return;
      case "applyImport":
        await applyImportBatch(id, command, controller.signal);
        return;
      case "recordDelivery":
        await recordPreparedDelivery(id, command, controller.signal);
        return;
      case "listDeliveries":
        await deliveryHistory(id, command);
        return;
      case "export":
        await exportWorkspace(id, command, controller.signal);
        return;
      case "close": {
        const activeWorkspace = workspace;
        const [
          importCleanup,
          retainedCleanup,
          savedPlanCleanup,
          databaseCleanup,
        ] = await Promise.allSettled([
          closeImports(),
          retryImportPreparationCleanup(),
          savedPlanCleanupOwners.close(),
          activeWorkspace?.database.close() ?? Promise.resolve(),
        ]);
        schemaPlans.clear();
        if (databaseCleanup.status === "fulfilled") workspace = null;
        const failures = [
          importCleanup,
          retainedCleanup,
          savedPlanCleanup,
          databaseCleanup,
        ].flatMap((result) =>
          result.status === "rejected" ? [result.reason] : [],
        );
        if (failures.length === 1) throw failures[0];
        if (failures.length > 1) {
          throw new AggregateError(
            failures,
            "The workspace could not finish closing its resources",
          );
        }
        scope.postMessage({ type: "closed", id });
        return;
      }
    }
  } finally {
    controllers.delete(id);
  }
}

scope.addEventListener("message", (event) => {
  if (event.data.type === "cancel") {
    controllers.get(event.data.targetId)?.abort();
    scope.postMessage({ type: "closed", id: event.data.id });
    return;
  }
  void handle(event.data.id, event.data).catch((error: unknown) => {
    postError(event.data.id, error);
  });
});
