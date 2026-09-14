import {
  APPLICATION_TABLE,
  CAPTURE_ROW_TABLE,
  CAPTURE_TABLE,
  COUNTERS_TABLE,
  DATABASE_METADATA_TABLE,
  DELIVERY_MEMBERSHIP_TABLE,
  DELIVERY_TABLE,
  IMPORT_REQUEST_TABLE,
  PLAN_TABLE,
  SOURCE_CONTENT_TABLE,
  SOURCE_FILE_TABLE,
  SOURCE_NAME_TABLE,
  TABLE_REGISTRY_TABLE,
} from "../metadata.js";
import {
  defineInternalStorageSchema,
  type InternalStorageSchema,
  type InternalStorageTable,
} from "./storage-schema.js";

function text<const Name extends string>(
  name: Name,
  nullable = false,
): { readonly name: Name; readonly storage: "text"; readonly nullable?: true } {
  return nullable
    ? { name, storage: "text", nullable: true }
    : { name, storage: "text" };
}

function integer<const Name extends string>(name: Name) {
  return { name, storage: "integer" } as const;
}

export const PREPARED_METADATA_TABLE = "_consultchimps_prepared";
export const PREPARED_CAPTURE_TABLE = "_consultchimps_prepared_captures";
export const PREPARED_BINDING_TABLE = "_consultchimps_prepared_bindings";
export const PREPARED_ROW_TABLE = "_consultchimps_prepared_rows";

export interface DatabaseStorageTables extends Readonly<
  Record<string, InternalStorageTable>
> {
  readonly database: InternalStorageTable;
  readonly tableRegistry: InternalStorageTable;
  readonly counters: InternalStorageTable;
  readonly sourceContents: InternalStorageTable;
  readonly sourceFiles: InternalStorageTable;
  readonly sourceNames: InternalStorageTable;
  readonly captures: InternalStorageTable;
  readonly importPlans: InternalStorageTable;
  readonly captureRows: InternalStorageTable;
  readonly importApplications: InternalStorageTable;
  readonly importRequests: InternalStorageTable;
  readonly deliveries: InternalStorageTable;
  readonly deliveryMemberships: InternalStorageTable;
}

export interface PreparedStorageTables extends Readonly<
  Record<string, InternalStorageTable>
> {
  readonly metadata: InternalStorageTable;
  readonly captures: InternalStorageTable;
  readonly bindings: InternalStorageTable;
  readonly rows: InternalStorageTable;
}

export const DATABASE_STORAGE: InternalStorageSchema<DatabaseStorageTables> =
  defineInternalStorageSchema({
    tables: {
      database: {
        name: DATABASE_METADATA_TABLE,
        columns: [
          text("database_id"),
          text("format"),
          integer("format_version"),
          integer("revision"),
        ],
        primaryKey: ["database_id"],
      },
      tableRegistry: {
        name: TABLE_REGISTRY_TABLE,
        columns: [
          text("table_name"),
          text("schema_json"),
          integer("schema_version"),
          integer("next_record_id"),
        ],
        primaryKey: ["table_name"],
      },
      counters: {
        name: COUNTERS_TABLE,
        columns: [text("counter_name"), integer("next_value")],
        primaryKey: ["counter_name"],
        copy: { key: ["counter_name"], order: 1 },
      },
      sourceContents: {
        name: SOURCE_CONTENT_TABLE,
        columns: [text("content_hash"), integer("byte_count")],
        primaryKey: ["content_hash"],
        copy: { key: ["content_hash"], order: 2 },
      },
      sourceFiles: {
        name: SOURCE_FILE_TABLE,
        columns: [
          text("source_file_id"),
          text("content_hash"),
          text("display_name"),
        ],
        primaryKey: ["source_file_id"],
        uniqueKeys: [["content_hash"]],
        copy: { key: ["source_file_id"], order: 3 },
      },
      sourceNames: {
        name: SOURCE_NAME_TABLE,
        columns: [text("source_file_id"), text("display_name")],
        primaryKey: ["source_file_id", "display_name"],
        copy: { key: ["source_file_id", "display_name"], order: 4 },
      },
      captures: {
        name: CAPTURE_TABLE,
        columns: [
          text("capture_id"),
          text("source_file_id"),
          text("source_key"),
          text("selection_key"),
          text("selection_label"),
          text("reader_version"),
          text("state"),
          integer("row_count"),
          text("columns_json"),
        ],
        primaryKey: ["capture_id"],
        uniqueKeys: [["source_file_id", "selection_key", "reader_version"]],
        copy: { key: ["capture_id"], order: 5 },
      },
      importPlans: {
        name: PLAN_TABLE,
        columns: [
          text("plan_id"),
          integer("plan_revision"),
          integer("baseline_revision"),
          text("state"),
          text("recipe_json"),
          text("conflicts_json"),
          text("decisions_json"),
          text("bindings_json"),
        ],
        primaryKey: ["plan_id", "plan_revision"],
        copy: { key: ["plan_id", "plan_revision"], order: 7 },
      },
      captureRows: {
        name: CAPTURE_ROW_TABLE,
        columns: [
          text("capture_id"),
          integer("source_row"),
          text("values_json"),
        ],
        primaryKey: ["capture_id", "source_row"],
        duckdb: { omitPrimaryKey: true },
        copy: { key: ["capture_id", "source_row"], order: 6 },
      },
      importApplications: {
        name: APPLICATION_TABLE,
        columns: [
          text("import_id"),
          text("application_key"),
          text("request_id"),
          text("capture_id"),
          text("table_name"),
          text("plan_id"),
          integer("plan_revision"),
          integer("row_count"),
        ],
        primaryKey: ["import_id"],
        uniqueKeys: [
          ["application_key"],
          ["request_id", "capture_id", "table_name"],
        ],
        copy: { key: ["import_id"], order: 8 },
      },
      importRequests: {
        name: IMPORT_REQUEST_TABLE,
        columns: [
          text("request_id"),
          text("plan_id"),
          integer("plan_revision"),
          text("import_ids_json"),
          text("capture_ids_json"),
          integer("row_count"),
        ],
        primaryKey: ["request_id"],
        copy: { key: ["request_id"], order: 9 },
      },
      deliveries: {
        name: DELIVERY_TABLE,
        columns: [
          text("delivery_id"),
          text("request_id"),
          text("context_json"),
        ],
        primaryKey: ["delivery_id"],
        uniqueKeys: [["request_id"]],
        copy: { key: ["delivery_id"], order: 10 },
      },
      deliveryMemberships: {
        name: DELIVERY_MEMBERSHIP_TABLE,
        columns: [text("delivery_id"), text("capture_id")],
        primaryKey: ["delivery_id", "capture_id"],
        copy: { key: ["delivery_id", "capture_id"], order: 11 },
      },
    },
  });

export const PREPARED_STORAGE: InternalStorageSchema<PreparedStorageTables> =
  defineInternalStorageSchema({
    tables: {
      metadata: {
        name: PREPARED_METADATA_TABLE,
        columns: [
          integer("format_version"),
          text("plan_id"),
          text("database_id"),
          integer("baseline_revision"),
          text("schema_fingerprint"),
          integer("plan_revision"),
          text("state"),
          text("recipe_json"),
          text("conflicts_json"),
          text("decisions_json"),
          text("review_fingerprint"),
        ],
        primaryKey: ["plan_id"],
      },
      captures: {
        name: PREPARED_CAPTURE_TABLE,
        columns: [
          text("capture_id"),
          text("source_file_id", true),
          text("source_key"),
          text("display_name"),
          text("selection_key"),
          text("selection_label"),
          text("reader_version"),
          text("content_hash"),
          integer("byte_count"),
          integer("reused"),
          integer("row_count"),
          text("columns_json"),
          text("row_checksum"),
        ],
        primaryKey: ["capture_id"],
      },
      bindings: {
        name: PREPARED_BINDING_TABLE,
        columns: [
          text("source_key"),
          text("selection_key"),
          text("capture_id"),
          text("display_name"),
        ],
        primaryKey: ["source_key", "selection_key"],
      },
      rows: {
        name: PREPARED_ROW_TABLE,
        columns: [
          text("capture_id"),
          integer("source_row"),
          text("values_json"),
        ],
        primaryKey: ["capture_id", "source_row"],
      },
    },
  });
