import type { DatabaseFormat } from "../schema.js";
import type { EngineRow, EngineTransaction } from "./engine.js";

export interface InternalTableStorage {
  readonly table: string;
  readonly columns: readonly string[];
  readonly integerColumns?: readonly string[];
  readonly nullableColumns?: readonly string[];
  readonly primaryKey?: readonly string[];
  readonly uniqueKeys?: readonly (readonly string[])[];
  readonly duckdbWithoutPrimaryKey?: boolean;
}

function enabled(value: unknown): boolean {
  return value === true || value === 1 || value === 1n;
}

export async function assertInternalTableStorage(options: {
  readonly query: EngineTransaction["query"];
  readonly format: DatabaseFormat;
  readonly layout: InternalTableStorage;
  readonly columns: readonly EngineRow[];
  readonly invalid: (details: Record<string, unknown>) => Error;
}): Promise<void> {
  const { query, format, layout, columns, invalid } = options;
  const primaryKey =
    format === "duckdb" && layout.duckdbWithoutPrimaryKey === true
      ? []
      : (layout.primaryKey ?? []);
  for (const column of columns) {
    const name = column["name"];
    const type = column["type"];
    const expectedType =
      typeof name === "string" && layout.integerColumns?.includes(name)
        ? "BIGINT"
        : "VARCHAR";
    const nullable =
      typeof name === "string" &&
      (layout.nullableColumns?.includes(name) === true ||
        (format === "sqlite" &&
          primaryKey.length === 1 &&
          primaryKey[0] === name));
    if (
      typeof name !== "string" ||
      typeof type !== "string" ||
      type.toUpperCase() !== expectedType ||
      enabled(column["notnull"]) === nullable ||
      column["dflt_value"] !== null ||
      (format === "sqlite" &&
        (String(column["hidden"]) !== "0" ||
          String(column["pk"]) !== String(primaryKey.indexOf(name) + 1)))
    ) {
      throw invalid({
        table: layout.table,
        column: name,
        storageMismatch: true,
      });
    }
  }

  const expectedKeys = [
    ...(primaryKey.length === 0
      ? []
      : [JSON.stringify(["PRIMARY KEY", primaryKey.join(",")])]),
    ...(layout.uniqueKeys ?? []).map((key) =>
      JSON.stringify(["UNIQUE", key.join(",")]),
    ),
  ].sort();
  let actualKeys: string[];
  if (format === "duckdb") {
    const constraints = await query(
      "SELECT constraint_type, array_to_string(constraint_column_names, ',') AS column_names FROM duckdb_constraints() WHERE database_name = current_database() AND schema_name = 'main' AND table_name = ? AND constraint_type <> 'NOT NULL'",
      [layout.table],
    );
    actualKeys = constraints.map((constraint) =>
      JSON.stringify([
        constraint["constraint_type"],
        constraint["column_names"],
      ]),
    );
    const indexes = await query(
      "SELECT index_name FROM duckdb_indexes() WHERE database_name = current_database() AND schema_name = 'main' AND table_name = ? AND is_unique",
      [layout.table],
    );
    if (indexes.length > 0) actualKeys.push("unsupported unique index");
  } else {
    const definitions = await query(
      "SELECT sql FROM main.sqlite_schema WHERE type = 'table' AND name = ?",
      [layout.table],
    );
    const sql = definitions[0]?.["sql"];
    const keywords =
      typeof sql === "string"
        ? sql.replace(
            /'(?:''|[^'])*'|"(?:""|[^"])*"|`(?:``|[^`])*`|\[[^\]]*\]|--[^\n]*|\/\*[\s\S]*?\*\//gu,
            " ",
          )
        : "";
    const keys = await query(
      "SELECT i.name AS index_name, i.origin, i.partial, x.name, x.cid, x.coll, x.desc FROM pragma_index_list(?, 'main') i JOIN pragma_index_xinfo(i.name, 'main') x WHERE i.\"unique\" = 1 AND x.key = 1 ORDER BY i.name, x.seqno",
      [layout.table],
    );
    const grouped = new Map<string, { type: string; columns: string[] }>();
    for (const key of keys) {
      const name = key["index_name"];
      if (
        typeof name !== "string" ||
        typeof key["name"] !== "string" ||
        (key["origin"] !== "pk" && key["origin"] !== "u") ||
        String(key["partial"]) !== "0" ||
        String(key["desc"]) !== "0" ||
        key["coll"] !== "BINARY"
      ) {
        throw invalid({ table: layout.table, constraintMismatch: true });
      }
      let group = grouped.get(name);
      if (group === undefined) {
        group = {
          type: key["origin"] === "pk" ? "PRIMARY KEY" : "UNIQUE",
          columns: [],
        };
        grouped.set(name, group);
      }
      group.columns.push(key["name"]);
    }
    actualKeys = [...grouped.values()].map((group) =>
      JSON.stringify([group.type, group.columns.join(",")]),
    );
    if (
      typeof sql !== "string" ||
      /\b(?:CHECK|COLLATE|GENERATED|DEFERRABLE|STRICT|FOREIGN)\b|\bWITHOUT\s+ROWID\b|\bON\s+CONFLICT\b/iu.test(
        keywords,
      )
    )
      actualKeys.push("unsupported table constraint");
  }
  if (JSON.stringify(actualKeys.sort()) !== JSON.stringify(expectedKeys)) {
    throw invalid({ table: layout.table, constraintMismatch: true });
  }
}
