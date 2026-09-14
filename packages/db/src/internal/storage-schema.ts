import type { DatabaseFormat } from "../schema.js";
import { quoteIdentifier } from "../schema.js";
import type { EngineRow, EngineTransaction, EngineValue } from "./engine.js";

export type InternalStorageKind = "integer" | "text";

export interface InternalStorageColumn<Name extends string = string> {
  readonly name: Name;
  readonly storage: InternalStorageKind;
  readonly nullable?: true;
}

export interface InternalStorageTable<Name extends string = string> {
  readonly name: string;
  readonly columns: readonly InternalStorageColumn<Name>[];
  readonly primaryKey?: readonly Name[];
  readonly uniqueKeys?: readonly (readonly Name[])[];
  readonly duckdb?: { readonly omitPrimaryKey: true };
  readonly copy?: {
    readonly key: readonly Name[];
    readonly order: number;
  };
}

export interface InternalStorageSchema<
  Tables extends Readonly<Record<string, InternalStorageTable>> = Readonly<
    Record<string, InternalStorageTable>
  >,
> {
  readonly tables: Tables;
}

export type InternalStorageRow = Readonly<Record<string, EngineValue>>;

export function defineInternalStorageSchema<
  const Tables extends Readonly<Record<string, InternalStorageTable>>,
>(schema: InternalStorageSchema<Tables>): InternalStorageSchema<Tables> {
  return schema;
}

function effectivePrimaryKey(
  table: InternalStorageTable,
  format: DatabaseFormat,
): readonly string[] {
  return format === "duckdb" && table.duckdb?.omitPrimaryKey === true
    ? []
    : (table.primaryKey ?? []);
}

function storageType(column: InternalStorageColumn): string {
  return column.storage === "integer" ? "BIGINT" : "VARCHAR";
}

function enabled(value: unknown): boolean {
  return value === true || value === 1 || value === 1n;
}

async function assertTableMatchesDescriptor(options: {
  readonly query: EngineTransaction["query"];
  readonly format: DatabaseFormat;
  readonly table: InternalStorageTable;
  readonly columns: readonly EngineRow[];
  readonly invalid: (details: Record<string, unknown>) => Error;
}): Promise<void> {
  const { query, format, table, columns, invalid } = options;
  const primaryKey = effectivePrimaryKey(table, format);
  const expectedColumns = new Map(
    table.columns.map((column) => [column.name, column]),
  );
  for (const column of columns) {
    const name = column["name"];
    const type = column["type"];
    const declared =
      typeof name === "string" ? expectedColumns.get(name) : undefined;
    const nullable =
      declared?.nullable === true ||
      (format === "sqlite" &&
        primaryKey.length === 1 &&
        primaryKey[0] === name);
    if (
      declared === undefined ||
      typeof type !== "string" ||
      type.toUpperCase() !== storageType(declared) ||
      enabled(column["notnull"]) === nullable ||
      column["dflt_value"] !== null ||
      (format === "sqlite" &&
        (String(column["hidden"]) !== "0" ||
          String(column["pk"]) !==
            String(primaryKey.indexOf(declared.name) + 1)))
    ) {
      throw invalid({ table: table.name, column: name, storageMismatch: true });
    }
  }

  const expectedKeys = [
    ...(primaryKey.length === 0
      ? []
      : [JSON.stringify(["PRIMARY KEY", primaryKey.join(",")])]),
    ...(table.uniqueKeys ?? []).map((key) =>
      JSON.stringify(["UNIQUE", key.join(",")]),
    ),
  ].sort();
  let actualKeys: string[];
  if (format === "duckdb") {
    const constraints = await query(
      "SELECT constraint_type, array_to_string(constraint_column_names, ',') AS column_names FROM duckdb_constraints() WHERE database_name = current_database() AND schema_name = 'main' AND table_name = ? AND constraint_type <> 'NOT NULL'",
      [table.name],
    );
    actualKeys = constraints.map((constraint) =>
      JSON.stringify([
        constraint["constraint_type"],
        constraint["column_names"],
      ]),
    );
    const indexes = await query(
      "SELECT index_name FROM duckdb_indexes() WHERE database_name = current_database() AND schema_name = 'main' AND table_name = ? AND is_unique",
      [table.name],
    );
    if (indexes.length > 0) actualKeys.push("unsupported unique index");
  } else {
    const definitions = await query(
      "SELECT sql FROM main.sqlite_schema WHERE type = 'table' AND name = ?",
      [table.name],
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
      [table.name],
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
        throw invalid({ table: table.name, constraintMismatch: true });
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
    throw invalid({ table: table.name, constraintMismatch: true });
  }
}

export function internalColumnNames(
  table: InternalStorageTable,
): readonly string[] {
  return table.columns.map(({ name }) => name);
}

export function internalCreateTableSql(
  table: InternalStorageTable,
  format: DatabaseFormat,
): string {
  const primaryKey = effectivePrimaryKey(table, format);
  const columns = table.columns.map((column) => {
    const isSinglePrimaryKey =
      primaryKey.length === 1 && primaryKey[0] === column.name;
    return [
      quoteIdentifier(column.name),
      " ",
      storageType(column),
      isSinglePrimaryKey
        ? " PRIMARY KEY"
        : column.nullable === true
          ? ""
          : " NOT NULL",
    ].join("");
  });
  const constraints = [
    ...(primaryKey.length > 1
      ? [`PRIMARY KEY (${primaryKey.map(quoteIdentifier).join(", ")})`]
      : []),
    ...(table.uniqueKeys ?? []).map(
      (key) => `UNIQUE (${key.map(quoteIdentifier).join(", ")})`,
    ),
  ];
  return `CREATE TABLE ${quoteIdentifier(table.name)} (${[...columns, ...constraints].join(", ")})`;
}

export async function createInternalTables(
  transaction: Pick<EngineTransaction, "execute">,
  schema: InternalStorageSchema,
  format: DatabaseFormat,
): Promise<void> {
  for (const table of Object.values(schema.tables)) {
    await transaction.execute(internalCreateTableSql(table, format));
  }
}

export async function validateInternalTables(options: {
  readonly query: EngineTransaction["query"];
  readonly format: DatabaseFormat;
  readonly schema: InternalStorageSchema;
  readonly invalid: (details: Record<string, unknown>) => Error;
}): Promise<void> {
  const tables = Object.values(options.schema.tables);
  const names = tables.map(({ name }) => name);
  const placeholders = names.map(() => "?").join(", ");
  const tableRows = await options.query(
    options.format === "sqlite"
      ? `SELECT name AS table_name FROM sqlite_master WHERE type = 'table' AND name IN (${placeholders})`
      : `SELECT table_name FROM information_schema.tables WHERE table_schema = 'main' AND table_type = 'BASE TABLE' AND table_name IN (${placeholders})`,
    names,
  );
  const existing = new Set(
    tableRows.flatMap((row) =>
      typeof row["table_name"] === "string" ? [row["table_name"]] : [],
    ),
  );
  const missingTables = names.filter((name) => !existing.has(name));
  if (missingTables.length > 0) throw options.invalid({ missingTables });

  for (const table of tables) {
    const columns = await options.query(
      options.format === "sqlite"
        ? "SELECT * FROM pragma_table_xinfo(?, 'main') ORDER BY cid"
        : "SELECT * FROM pragma_table_info(?) ORDER BY cid",
      [table.name],
    );
    const actualColumns = columns.map((row) =>
      typeof row["name"] === "string" ? row["name"] : null,
    );
    const expectedColumns = internalColumnNames(table);
    if (
      actualColumns.length !== expectedColumns.length ||
      actualColumns.some((column, index) => column !== expectedColumns[index])
    ) {
      throw options.invalid({
        table: table.name,
        expectedColumns: [...expectedColumns],
        actualColumns,
      });
    }
    await assertTableMatchesDescriptor({
      query: options.query,
      format: options.format,
      table,
      columns,
      invalid: options.invalid,
    });
  }
}

export async function insertInternalRow(
  transaction: Pick<EngineTransaction, "execute">,
  table: InternalStorageTable,
  row: InternalStorageRow,
  options: { readonly onConflict?: "do-nothing" } = {},
): Promise<void> {
  const columns = internalColumnNames(table);
  const expected = new Set(columns);
  const entries = Object.entries(row);
  const actual = new Set(entries.map(([column]) => column));
  const missingColumns = columns.filter((column) => !actual.has(column));
  const unexpectedColumns = entries
    .map(([column]) => column)
    .filter((column) => !expected.has(column));
  if (missingColumns.length > 0 || unexpectedColumns.length > 0) {
    throw new Error(
      `Internal row for ${quoteIdentifier(table.name)} does not match its storage layout. Missing: ${missingColumns.join(", ") || "none"}. Unexpected: ${unexpectedColumns.join(", ") || "none"}.`,
    );
  }
  const values = new Map(entries);
  const suffix =
    options.onConflict === "do-nothing" ? " ON CONFLICT DO NOTHING" : "";
  await transaction.execute(
    `INSERT INTO ${quoteIdentifier(table.name)} (${columns.map(quoteIdentifier).join(", ")}) VALUES (${columns.map(() => "?").join(", ")})${suffix}`,
    columns.map((column) => {
      const value = values.get(column);
      if (value === undefined) {
        throw new Error(
          `Internal row for ${quoteIdentifier(table.name)} has an undefined value for ${quoteIdentifier(column)}. Use null for a nullable column.`,
        );
      }
      return value;
    }),
  );
}

export interface InternalCopyTable {
  readonly name: string;
  readonly columns: readonly InternalStorageColumn[];
  readonly key: readonly string[];
}

export function internalCopyTables(
  schema: InternalStorageSchema,
): readonly InternalCopyTable[] {
  return Object.values(schema.tables)
    .flatMap((table) =>
      table.copy === undefined
        ? []
        : [
            {
              name: table.name,
              columns: table.columns,
              key: table.copy.key,
              order: table.copy.order,
            },
          ],
    )
    .sort((left, right) => left.order - right.order)
    .map(({ name, columns, key }) => ({ name, columns, key }));
}
