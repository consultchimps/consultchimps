import { isConsultChimpsError } from "@consultchimps/core";
import { openReadOnlySqlite } from "@consultchimps/db/sqlite-read";
import type {
  ReadOnlySqlite,
  SqliteReadRuntimeConfig,
  SqliteReadValue,
} from "@consultchimps/db/sqlite-read";
import { ConsultChimpsError } from "@consultchimps/core";
import {
  cleanupRequired,
  runtimeUnavailable,
  unreadableModel,
} from "./errors.js";

/**
 * The storage catalog inside `metadata.sqlitedb`, read through
 * `@consultchimps/db/sqlite-read`. That reader refuses every PRAGMA except the
 * query form of `data_version`, so optional columns are discovered from an
 * empty projection rather than `PRAGMA table_info`, and the result is paged
 * with an explicit row limit rather than trusting the reader's default.
 */

/** One page of the catalog query. Small enough that no page nears maxResultBytes. */
const PAGE_ROWS = 2048;
/** A model past this many column-partitions is refused before any paging. */
const MAX_CATALOG_ROWS = 1_048_576;

export interface CatalogColumn {
  readonly id: number;
  readonly name: string;
  readonly storagePosition: number;
  readonly dataType: number;
  readonly hidden: boolean;
  readonly calculated: boolean;
  readonly dax: string | undefined;
  readonly dictionary: string | null;
  readonly hierarchyIndex: string | null;
  /** Partition members in ascending StoragePosition then ID. */
  readonly idfs: readonly string[];
  readonly baseId: bigint;
  readonly magnitude: number;
  /** ColumnStorage.Statistics_RowCount, the catalog's declared row count. */
  readonly rowCount: number;
}

export interface CatalogTable {
  readonly id: number;
  readonly name: string;
  readonly hidden: boolean;
  readonly calculated: boolean;
  readonly dax: string | undefined;
  /** The declared row count, the largest its columns agree on. */
  rowCount: number;
  readonly columns: readonly CatalogColumn[];
}

export interface Catalog {
  readonly tables: readonly CatalogTable[];
  /** The reader's linear memory, for the pipeline's peak accounting. */
  readonly wasmMemoryBytes: number;
}

/** Map the database reader's controlled codes into the Power BI refusal contract. */
function mapReadError(error: unknown): never {
  if (isConsultChimpsError(error)) {
    if (error.code === "DB_SQLITE_READ_RUNTIME_UNAVAILABLE")
      throw runtimeUnavailable("sqlite", "load");
    if (error.code === "DB_SQLITE_READ_LIMIT_EXCEEDED") {
      const details = error.details as { option?: string; limit?: number };
      throw new ConsultChimpsError(
        "PBI_EXPORT_LIMIT_EXCEEDED",
        "Reading the model catalog would exceed a configured SQLite limit. Use a smaller model, or explicitly raise the named limit if your environment can support it.",
        {
          details: {
            stage: "catalog",
            option: details.option ?? "maxResultBytes",
            limit: details.limit ?? 0,
          },
        },
      );
    }
    if (error.code === "DB_SQLITE_READ_CLEANUP_REQUIRED")
      throw cleanupRequired();
  }
  throw unreadableModel("catalog");
}

function asNumber(value: SqliteReadValue | undefined): number {
  if (typeof value === "bigint") return Number(value);
  if (typeof value === "number") return value;
  throw unreadableModel("catalog");
}

function asBigInt(value: SqliteReadValue | undefined): bigint {
  if (typeof value === "bigint") return value;
  if (typeof value === "number" && Number.isInteger(value))
    return BigInt(value);
  return 0n;
}

function asText(value: SqliteReadValue | undefined): string | null {
  return typeof value === "string" ? value : null;
}

/**
 * The catalog's declared row count for a column's storage. It is the ceiling
 * every per-column parser measures a declared record count against, so a
 * missing or nonsensical value has to be zero rather than a licence to trust
 * whatever the member's own header says.
 */
function declaredRows(value: SqliteReadValue | undefined): number {
  const rows =
    typeof value === "bigint"
      ? Number(value)
      : typeof value === "number"
        ? value
        : 0;
  return Number.isSafeInteger(rows) && rows > 0 ? rows : 0;
}

/** DAX is capped so a pathological model cannot make the manifest unbounded. */
const DAX_LIMIT = 32_768;

export function capDax(value: string | null): string | undefined {
  if (value === null || value.length === 0) return undefined;
  if (value.length <= DAX_LIMIT) return value;
  // Never split a surrogate pair when shortening.
  const cut = value.slice(0, DAX_LIMIT);
  const last = cut.charCodeAt(cut.length - 1);
  return last >= 0xd800 && last <= 0xdbff ? cut.slice(0, -1) : cut;
}

/** Column presence from an empty projection, which needs no PRAGMA. */
function columnsOf(reader: ReadOnlySqlite, table: string): Set<string> {
  try {
    return new Set(reader.query(`SELECT * FROM "${table}" LIMIT 0`).columns);
  } catch {
    return new Set<string>();
  }
}

export interface CatalogReadResult extends Catalog {
  close(): void;
}

/**
 * Open the catalog and read every exportable column, ordered exactly as ADR
 * Decision 5 requires: tables by numeric ID, columns by StoragePosition then
 * ID, partitions by StoragePosition then ID.
 */
export async function readCatalog(
  bytes: Uint8Array,
  runtime: SqliteReadRuntimeConfig | undefined,
): Promise<CatalogReadResult> {
  let reader: ReadOnlySqlite;
  try {
    reader = await openReadOnlySqlite(bytes, {
      ...(runtime === undefined ? {} : { runtime }),
      maxDatabaseBytes: Math.max(bytes.byteLength, 1),
      maxRows: PAGE_ROWS,
      maxResultBytes: 8 * 1024 * 1024,
      maxSqlBytes: 8 * 1024,
    });
  } catch (error) {
    mapReadError(error);
  }
  try {
    const columnFields = columnsOf(reader, "Column");
    if (columnFields.size === 0) throw unreadableModel("catalog");
    const typeColumn = columnFields.has("Type") ? "c.Type" : "c.BindingType";
    const dataTypeColumn = columnFields.has("InferredDataType")
      ? "CASE WHEN c.ExplicitDataType = 1 THEN c.InferredDataType ELSE c.ExplicitDataType END"
      : "c.ExplicitDataType";
    const expressionColumn = columnFields.has("Expression")
      ? "c.Expression"
      : "NULL";
    const tableFields = columnsOf(reader, "Table");
    const hiddenColumn = tableFields.has("IsHidden") ? "t.IsHidden" : "0";

    const from = `
      FROM Column c
      JOIN [Table] t ON c.TableID = t.ID
      JOIN ColumnStorage cs ON c.ColumnStorageID = cs.ID
      JOIN AttributeHierarchy ah ON ah.ColumnID = c.ID
      JOIN AttributeHierarchyStorage ahs ON ah.AttributeHierarchyStorageID = ahs.ID
      LEFT JOIN StorageFile sfh ON sfh.ID = ahs.StorageFileID
      LEFT JOIN DictionaryStorage ds ON ds.ID = cs.DictionaryStorageID
      LEFT JOIN StorageFile sfd ON sfd.ID = ds.StorageFileID
      JOIN ColumnPartitionStorage cps ON cps.ColumnStorageID = cs.ID
      JOIN StorageFile sfi ON sfi.ID = cps.StorageFileID
      JOIN PartitionStorage ps ON ps.ID = cps.PartitionStorageID
      WHERE ${typeColumn} IN (1,2,4)`;

    const total = asNumber(
      reader.query(`SELECT COUNT(*) ${from}`).rows[0]?.[0],
    );
    if (total > MAX_CATALOG_ROWS)
      throw new ConsultChimpsError(
        "PBI_EXPORT_LIMIT_EXCEEDED",
        "This model has more column partitions than the reader will enumerate. Use a smaller model.",
        {
          details: {
            stage: "catalog",
            option: "catalogRows",
            limit: MAX_CATALOG_ROWS,
            required: total,
          },
        },
      );

    // Calculated tables: a partition of type 2 carries the table's DAX.
    const calculated = new Map<number, string | undefined>();
    for (const row of reader.query(
      "SELECT TableID, QueryDefinition FROM Partition WHERE Type = 2 ORDER BY TableID, ID LIMIT 2048",
    ).rows) {
      const id = asNumber(row[0]);
      if (!calculated.has(id)) calculated.set(id, capDax(asText(row[1])));
    }

    const select = `
      SELECT t.ID, t.Name, ${hiddenColumn},
             c.ID, COALESCE(c.ExplicitName, c.InferredName), c.IsHidden,
             ${typeColumn}, ${expressionColumn}, ${dataTypeColumn},
             sfd.FileName, sfh.FileName, sfi.FileName,
             ds.BaseId, ds.Magnitude,
             cs.StoragePosition, ps.StoragePosition, ps.ID,
             cs.Statistics_RowCount
      ${from}
      ORDER BY t.ID, cs.StoragePosition, c.ID, ps.StoragePosition, ps.ID
      LIMIT ? OFFSET ?`;

    const tables: CatalogTable[] = [];
    const columnsById = new Map<
      string,
      { column: CatalogColumn; idfs: string[] }
    >();
    let currentTable:
      { table: CatalogTable; columns: CatalogColumn[] } | undefined;

    for (let offset = 0; offset < total; offset += PAGE_ROWS) {
      const page = reader.query(select, [PAGE_ROWS, offset]);
      for (const row of page.rows) {
        const tableId = asNumber(row[0]);
        const tableName = asText(row[1]) ?? "";
        const tableHidden = asNumber(row[2]) !== 0;
        if (currentTable === undefined || currentTable.table.id !== tableId) {
          const columns: CatalogColumn[] = [];
          const table: CatalogTable = {
            id: tableId,
            name: tableName,
            hidden: tableHidden,
            calculated: calculated.has(tableId),
            dax: calculated.get(tableId),
            rowCount: 0,
            columns,
          };
          currentTable = { table, columns };
          tables.push(table);
        }
        const columnId = asNumber(row[3]);
        const key = `${tableId}:${columnId}`;
        const seen = columnsById.get(key);
        const idf = asText(row[11]);
        if (seen !== undefined) {
          // A second partition for the same column. The query's ORDER BY has
          // already put them in storage order, so appending preserves it.
          if (idf !== null) seen.idfs.push(idf);
          continue;
        }
        const idfs: string[] = idf === null ? [] : [idf];
        const columnType = asNumber(row[6]);
        const column: CatalogColumn = {
          id: columnId,
          name: asText(row[4]) ?? "",
          storagePosition: asNumber(row[14]),
          dataType: asNumber(row[8]),
          hidden: asNumber(row[5]) !== 0,
          calculated: columnType === 2,
          dax: columnType === 2 ? capDax(asText(row[7])) : undefined,
          dictionary: asText(row[9]),
          hierarchyIndex: asText(row[10]),
          idfs,
          baseId: asBigInt(row[12]),
          magnitude:
            typeof row[13] === "number" ? row[13] : Number(row[13] ?? 1),
          rowCount: declaredRows(row[17]),
        };
        currentTable.table.rowCount = Math.max(
          currentTable.table.rowCount,
          column.rowCount,
        );
        columnsById.set(key, { column, idfs });
        currentTable.columns.push(column);
      }
      if (page.rows.length < PAGE_ROWS) break;
    }
    return {
      tables,
      wasmMemoryBytes: reader.wasmMemoryBytes,
      close(): void {
        try {
          reader.close();
        } catch (error) {
          mapReadError(error);
        }
      },
    };
  } catch (error) {
    try {
      reader.close();
    } catch {
      // The read failure below is the actionable one; whatever could not be
      // released belongs to this reader alone and is unreachable afterwards.
    }
    if (isConsultChimpsError(error) && error.code.startsWith("PBI_"))
      throw error;
    mapReadError(error);
  }
}
