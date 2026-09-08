/**
 * Turning a chosen file into the tables a workspace import can create.
 *
 * This is the only place in the app that knows a `.xlsx` holds worksheets and a
 * `.csv` holds one table. Everything after it is format-neutral: the reader
 * hands `@consultchimps/db` a `Table` and the library decides column types,
 * names, and Record IDs, so a future `db import` command reaches the same
 * tables from the same file without any of this code.
 *
 * It lives beside the workspace worker rather than inside it because the
 * workbook reader pulls in a spreadsheet engine, and the worker imports this
 * module on demand so merely opening a workspace never fetches it.
 */
import { ConsultChimpsError } from "@consultchimps/core";
import {
  parseCsvTable,
  suggestRecordIdPrefix,
  suggestTableName,
  type ImportTableRequest,
} from "@consultchimps/db";
import type { Table } from "@consultchimps/tabular";

import type {
  ImportSourceDescription,
  ImportTableChoice,
} from "./workspace-protocol";

const CSV_NAME = /\.csv$/iu;
const WORKBOOK_NAME = /\.(?:xlsx|xlsm)$/iu;

/** One table a file offers, under the name the file gives it. */
interface ImportSource {
  readonly name: string;
  /** Absent when the worksheet held no header row with rows of values under it. */
  readonly table: Table | undefined;
  /**
   * Cells holding a formula the workbook carries no calculated value for. A
   * workbook Excel has calculated and saved has none.
   */
  readonly uncachedFormulaCells: number;
}

/** The file name without the extension the import recognized it by. */
function withoutExtension(fileName: string): string {
  return fileName.replace(CSV_NAME, "").replace(WORKBOOK_NAME, "");
}

function decodeUtf8(bytes: Uint8Array, fileName: string): string {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch (error) {
    // Silently replacing the bytes it cannot read would import damaged text and
    // say nothing, so a file in another encoding is refused with the fix.
    throw new ConsultChimpsError(
      "WORKSPACE_IMPORT_NOT_UTF8",
      `"${fileName}" is not UTF-8 text, so its characters cannot be read reliably. Save it again as CSV UTF-8 and import it once more.`,
      { cause: error, details: { file: fileName } },
    );
  }
}

/**
 * Refuse a worksheet whose formulas have never been calculated.
 *
 * A formula cell with no cached result reads as empty, because empty is all the
 * file says, so importing it would create a table with holes where the numbers
 * should be and report a clean success. There is no way to work the values out
 * here: the calculation belongs to Excel. So the import stops and says what to
 * do about it, which is the one outcome that neither loses data nor pretends.
 */
function assertCalculated(source: ImportSource, fileName: string): void {
  if (source.uncachedFormulaCells === 0) {
    return;
  }
  const cells =
    source.uncachedFormulaCells === 1
      ? "1 cell"
      : `${source.uncachedFormulaCells} cells`;
  throw new ConsultChimpsError(
    "WORKSPACE_IMPORT_UNCACHED_FORMULAS",
    `${cells} in "${source.name}" hold a formula that "${fileName}" carries no calculated value for, so importing it would leave those values empty. Open the workbook in Excel, let it calculate, save it, and import it again.`,
    {
      details: {
        cells: source.uncachedFormulaCells,
        file: fileName,
        source: source.name,
      },
    },
  );
}

/** Read every table a file offers, in the order the file lists them. */
async function readSources(
  fileName: string,
  bytes: Uint8Array,
): Promise<ImportSource[]> {
  if (CSV_NAME.test(fileName)) {
    return [
      {
        name: fileName,
        table: parseCsvTable(decodeUtf8(bytes, fileName), { file: fileName }),
        // Delimited text holds values, never formulas.
        uncachedFormulaCells: 0,
      },
    ];
  }
  if (!WORKBOOK_NAME.test(fileName)) {
    throw new ConsultChimpsError(
      "WORKSPACE_IMPORT_UNSUPPORTED_FILE",
      `"${fileName}" is not a file this import reads. Choose an Excel .xlsx or .xlsm workbook, or a .csv file.`,
      { details: { file: fileName } },
    );
  }

  // Loaded on demand: the workbook readers carry a spreadsheet engine and an
  // OOXML parser, and a visitor who only opens and saves a workspace should
  // never download either.
  const { describeWorkbookBytes, readWorkbookTablesBytes } =
    await import("@consultchimps/xlsx/bytes");
  const input = { name: fileName, bytes };
  // Two readers, because they answer two questions and only one of them can
  // answer the second. The table reader gives the data; the description gives
  // every worksheet, including the ones that yield no table, and the count of
  // formula cells the workbook carries no calculated value for. That count
  // cannot come from the table reader: its spreadsheet engine drops a numeric
  // formula cell with no cached value while parsing, so by the time a `Table`
  // exists the condition is gone.
  const [{ description }, tables] = await Promise.all([
    describeWorkbookBytes(input),
    readWorkbookTablesBytes(input),
  ]);
  const tableBySheet = new Map(
    tables.map((table) => [table.source?.sheet ?? "", table]),
  );
  const sources = description.sheets
    .map((sheet) => ({
      name: sheet.name,
      table: tableBySheet.get(sheet.name),
      uncachedFormulaCells: sheet.uncachedFormulaCells,
    }))
    // A worksheet with neither data nor uncalculated formulas has nothing to
    // say, so it is left out. One that cannot be imported because of its
    // formulas is kept, so the visitor is told why rather than left to wonder
    // where their worksheet went.
    .filter(
      (source) => source.table !== undefined || source.uncachedFormulaCells > 0,
    );

  if (sources.length === 0) {
    throw new ConsultChimpsError(
      "WORKSPACE_IMPORT_NO_WORKSHEETS",
      `No visible worksheet in "${fileName}" has a header row with rows under it, so there is nothing to import.`,
      { details: { file: fileName } },
    );
  }
  return sources;
}

/**
 * List what a file could contribute, with a suggested table name and Record ID
 * prefix for each. Nothing is created; the page shows the suggestions in a form
 * the visitor edits before any import runs.
 */
export async function describeImportSources(
  fileName: string,
  bytes: Uint8Array,
): Promise<ImportSourceDescription[]> {
  const sources = await readSources(fileName, bytes);
  const single = sources.length === 1;
  return sources.map((source) => {
    // A one-table file is named after the file, not after a worksheet the
    // visitor never sees, which is what makes "customers.csv" suggest
    // "customers" rather than the file name with its extension attached.
    const label = single ? withoutExtension(source.name) : source.name;
    const suggestedTableName = suggestTableName(label);
    return {
      name: source.name,
      rowCount: source.table?.rows.length ?? 0,
      columnCount: source.table?.columns.length ?? 0,
      uncachedFormulaCells: source.uncachedFormulaCells,
      suggestedTableName,
      suggestedRecordIdPrefix: suggestRecordIdPrefix(suggestedTableName),
    };
  });
}

/**
 * Match the visitor's choices back to the file's tables. A choice naming a
 * source the file does not hold is refused rather than skipped, because
 * importing fewer tables than were asked for is the kind of quiet difference
 * nobody notices until the data is missing.
 */
export async function resolveImportRequests(
  fileName: string,
  bytes: Uint8Array,
  choices: readonly ImportTableChoice[],
): Promise<ImportTableRequest[]> {
  const sources = new Map(
    (await readSources(fileName, bytes)).map((source) => [source.name, source]),
  );
  return choices.map((choice) => {
    const source = sources.get(choice.source);
    if (source === undefined) {
      throw new ConsultChimpsError(
        "WORKSPACE_IMPORT_SOURCE_MISSING",
        `"${fileName}" no longer holds "${choice.source}", so it cannot be imported. Choose the file again.`,
        { details: { file: fileName, source: choice.source } },
      );
    }
    // Checked here as well as in the page, because the page can only hold a
    // button back and this is what decides whether anything is written.
    assertCalculated(source, fileName);
    if (source.table === undefined) {
      throw new ConsultChimpsError(
        "WORKSPACE_IMPORT_SOURCE_EMPTY",
        `"${choice.source}" holds no header row with rows under it, so there is nothing to import from it.`,
        { details: { file: fileName, source: choice.source } },
      );
    }
    return {
      name: choice.tableName,
      recordId: {
        prefix: choice.recordIdPrefix,
        padding: choice.recordIdPadding,
      },
      table: source.table,
    };
  });
}
