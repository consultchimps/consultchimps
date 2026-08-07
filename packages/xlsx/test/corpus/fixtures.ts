/**
 * Conformance corpus fixture generators (ARCHITECTURE.md, Phase 0).
 *
 * Every fixture is built as a complete OOXML package by hand so the corpus can
 * express structures SheetJS cannot round-trip: shared and array formulas,
 * conditional formatting, data validation, cell comments, sheet-scoped defined
 * names, very-hidden sheets, pivot tables with their caches, and a macro
 * project. Building the package directly also keeps the byte layout stable, so
 * a behaviour pin fails when the library changes rather than when SheetJS does.
 *
 * The corpus is authored in PAIRS: `shape: "table"` and `shape: "range"` place
 * the same logical rows in the same cells, differing only in whether an Excel
 * Table part claims the region. Operation tests run both and pin any asymmetry.
 */
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import JSZip from "jszip";

/** Which binding a fixture exercises: an Excel Table or a plain worksheet range. */
export type CorpusShape = "table" | "range";

/** How the per-row `Doubled` column and the aggregate cells are expressed. */
export type CorpusFormulaFlavor = "none" | "a1" | "structured";

export interface CorpusWorkbookOptions {
  /** Excel Table or plain worksheet range. Required; fixtures come in pairs. */
  shape: CorpusShape;
  /** Per-row and aggregate formula style. `structured` requires the table shape. */
  formulas?: CorpusFormulaFlavor | undefined;
  /** Add an array formula at Data!G4 (outside the region). */
  arrayFormula?: boolean | undefined;
  /** Express the `Ratio` column as a shared formula instead of literals. */
  sharedFormula?: boolean | undefined;
  /** Add a formula with no cached value at Summary!B4. */
  uncachedFormula?: boolean | undefined;
  /** Table shape only: append a totals row below the data rows. */
  totalsRow?: boolean | undefined;
  /** Merged cells, conditional formatting, data validation, and a hyperlink. */
  dependents?: boolean | undefined;
  /** A cell comment (with its legacy VML drawing) anchored on Data!B6. */
  comments?: boolean | undefined;
  /** Workbook-scoped `CorpusRange` and sheet-scoped `LocalNote`. */
  definedNames?: boolean | undefined;
  /** A `Summary` worksheet carrying cross-sheet cached aggregates. */
  summarySheet?: boolean | undefined;
  /** A second data block below the region, at Data!A12. */
  footerBlock?: boolean | undefined;
  /** A `hidden` worksheet that also carries the split column, and a `veryHidden` one. */
  hiddenSheets?: boolean | undefined;
  /** Emit `xl/calcChain.xml`. Defaults to true whenever formulas are present. */
  calcChain?: boolean | undefined;
  /** A pivot table on `Summary` with its cache definition and cache records. */
  pivot?: boolean | undefined;
  /** Macro-enabled package: a stub `xl/vbaProject.bin` and the .xlsm content types. */
  macro?: boolean | undefined;
}

export interface CorpusDataRow {
  /** Value of the `Record` column. */
  record: number;
  /** Value of the `Client` column. */
  client: string;
  /** Value of the `Group` column, which every split in the corpus splits by. */
  group: string;
  /** Value of the `Amount` column. */
  amount: number;
  /** Worksheet row number holding this record in the source fixture. */
  row: number;
}

/** The six logical records every paired fixture carries in rows 4 to 9. */
export const CORPUS_ROWS: readonly CorpusDataRow[] = [
  { record: 1, client: "Client A", group: "Alpha", amount: 10, row: 4 },
  { record: 2, client: "Client B", group: "Beta", amount: 20, row: 5 },
  { record: 3, client: "Client C", group: "Alpha", amount: 30, row: 6 },
  { record: 4, client: "Client D", group: "Beta", amount: 40, row: 7 },
  { record: 5, client: "Client E", group: "Gamma", amount: 50, row: 8 },
  { record: 6, client: "Client F", group: "Alpha", amount: 60, row: 9 },
];

/** Header text of the column every corpus split groups by. */
export const CORPUS_SPLIT_COLUMN = "Group";
/** Worksheet row holding the region headers. */
export const CORPUS_HEADER_ROW = 3;
/** First worksheet row holding a data record. */
export const CORPUS_FIRST_DATA_ROW = 4;
/** Last worksheet row holding a data record. */
export const CORPUS_LAST_DATA_ROW = 9;
/** Worksheet row holding the Excel Table totals row, when one is present. */
export const CORPUS_TOTALS_ROW = 10;
/** Worksheet row holding the footer / second data block, when one is present. */
export const CORPUS_FOOTER_ROW = 12;
/** Sum of every `Amount` in the corpus, cached by the aggregate formulas. */
export const CORPUS_TOTAL_AMOUNT = 210;
/** Worksheet holding the split region. */
export const CORPUS_SHEET = "Data";
/** Worksheet holding cross-sheet aggregates and the pivot table. */
export const CORPUS_SUMMARY_SHEET = "Summary";
/** Hidden worksheet that also carries the split column. */
export const CORPUS_HIDDEN_SHEET = "Hidden";
/** Very-hidden worksheet without the split column. */
export const CORPUS_VERY_HIDDEN_SHEET = "VeryHidden";
/** Display name of the Excel Table in the `table` shape. */
export const CORPUS_TABLE_NAME = "DataTable";
/** Workbook-scoped defined name covering the region. */
export const CORPUS_RANGE_NAME = "CorpusRange";
/** Sheet-scoped defined name pointing at the footer block. */
export const CORPUS_LOCAL_NAME = "LocalNote";
/** Text held in Data!H6, outside the region, on an Alpha row. */
export const CORPUS_SIDE_NOTE = "Alpha side note";
/** Text of the cell comment anchored on Data!B6. */
export const CORPUS_COMMENT_TEXT = "Corpus comment on an Alpha row";
/** Stub macro project bytes carried by macro-enabled fixtures. */
export const CORPUS_VBA_BYTES: readonly number[] = [
  0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
];

/** Part paths a corpus fixture may contain. */
export const CORPUS_PARTS = {
  calcChain: "xl/calcChain.xml",
  comments: "xl/comments1.xml",
  contentTypes: "[Content_Types].xml",
  dataSheet: "xl/worksheets/sheet1.xml",
  hiddenSheet: "xl/worksheets/sheet3.xml",
  pivotCacheDefinition: "xl/pivotCache/pivotCacheDefinition1.xml",
  pivotCacheRecords: "xl/pivotCache/pivotCacheRecords1.xml",
  pivotTable: "xl/pivotTables/pivotTable1.xml",
  sharedStrings: "xl/sharedStrings.xml",
  summarySheet: "xl/worksheets/sheet2.xml",
  table: "xl/tables/table1.xml",
  vbaProject: "xl/vbaProject.bin",
  veryHiddenSheet: "xl/worksheets/sheet4.xml",
  workbook: "xl/workbook.xml",
} as const;

const NAMESPACE_MAIN =
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NAMESPACE_RELATIONSHIPS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const VBA_RELATIONSHIP_TYPE =
  "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
// Identical options must produce identical bytes, so entries carry a fixed date.
const FIXED_ENTRY_DATE = new Date(Date.UTC(1980, 0, 1));

interface Relationship {
  id: string;
  target: string;
  type: string;
  external?: boolean | undefined;
}

interface CellSpec {
  ref: string;
  style?: number | undefined;
  type?: "b" | "n" | "s" | "str" | undefined;
  value?: string | undefined;
  formula?: string | undefined;
  formulaAttributes?: string | undefined;
}

interface ResolvedOptions {
  arrayFormula: boolean;
  calcChain: boolean;
  comments: boolean;
  definedNames: boolean;
  dependents: boolean;
  footerBlock: boolean;
  formulas: CorpusFormulaFlavor;
  hiddenSheets: boolean;
  macro: boolean;
  pivot: boolean;
  shape: CorpusShape;
  sharedFormula: boolean;
  summarySheet: boolean;
  totalsRow: boolean;
  uncachedFormula: boolean;
}

function escapeXml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

class StringTable {
  private readonly indexByText = new Map<string, number>();
  private readonly texts: string[] = [];
  private references = 0;

  public index(text: string): number {
    this.references += 1;
    const existing = this.indexByText.get(text);
    if (existing !== undefined) {
      return existing;
    }
    const index = this.texts.length;
    this.indexByText.set(text, index);
    this.texts.push(text);
    return index;
  }

  public xml(): string {
    const items = this.texts
      .map((text) => `<si><t>${escapeXml(text)}</t></si>`)
      .join("");
    return `${XML_DECLARATION}<sst xmlns="${NAMESPACE_MAIN}" count="${this.references}" uniqueCount="${this.texts.length}">${items}</sst>`;
  }
}

class RelationshipTable {
  private readonly relationships: Relationship[] = [];

  public add(type: string, target: string, external = false): string {
    const id = `rId${this.relationships.length + 1}`;
    this.relationships.push({ id, target, type, external });
    return id;
  }

  public get size(): number {
    return this.relationships.length;
  }

  public xml(): string {
    const entries = this.relationships
      .map(
        (relationship) =>
          `<Relationship Id="${relationship.id}" Type="${relationship.type}" Target="${escapeXml(
            relationship.target,
          )}"${relationship.external ? ' TargetMode="External"' : ""}/>`,
      )
      .join("");
    return `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;
  }
}

function cellXml(spec: CellSpec): string {
  const attributes = [`r="${spec.ref}"`];
  if (spec.style !== undefined) {
    attributes.push(`s="${spec.style}"`);
  }
  if (spec.type !== undefined && spec.type !== "n") {
    attributes.push(`t="${spec.type}"`);
  }
  const children: string[] = [];
  if (spec.formula !== undefined) {
    const openingAttributes = spec.formulaAttributes ?? "";
    children.push(
      spec.formula === ""
        ? `<f${openingAttributes}/>`
        : `<f${openingAttributes}>${escapeXml(spec.formula)}</f>`,
    );
  }
  if (spec.value !== undefined) {
    children.push(`<v>${escapeXml(spec.value)}</v>`);
  }
  return children.length === 0
    ? `<c ${attributes.join(" ")}/>`
    : `<c ${attributes.join(" ")}>${children.join("")}</c>`;
}

function rowXml(rowNumber: number, cells: CellSpec[]): string {
  return `<row r="${rowNumber}">${cells.map(cellXml).join("")}</row>`;
}

function resolveOptions(options: CorpusWorkbookOptions): ResolvedOptions {
  const formulas = options.formulas ?? "a1";
  if (formulas === "structured" && options.shape !== "table") {
    throw new Error(
      'Structured-reference formulas require the "table" shape; the range shape has no table to reference.',
    );
  }
  const summarySheet = options.summarySheet ?? true;
  const uncachedFormula = options.uncachedFormula ?? false;
  if (uncachedFormula && !summarySheet) {
    throw new Error(
      "The uncachedFormula fixture lives on the Summary worksheet; enable summarySheet.",
    );
  }
  const pivot = options.pivot ?? false;
  if (pivot && !summarySheet) {
    throw new Error(
      "The pivot table is anchored on the Summary worksheet; enable summarySheet.",
    );
  }
  return {
    arrayFormula: options.arrayFormula ?? false,
    calcChain: options.calcChain ?? formulas !== "none",
    comments: options.comments ?? true,
    definedNames: options.definedNames ?? true,
    dependents: options.dependents ?? true,
    footerBlock: options.footerBlock ?? true,
    formulas,
    hiddenSheets: options.hiddenSheets ?? true,
    macro: options.macro ?? false,
    pivot,
    shape: options.shape,
    sharedFormula: options.sharedFormula ?? false,
    summarySheet,
    totalsRow: options.totalsRow ?? options.shape === "table",
    uncachedFormula,
  };
}

/** The Excel Table reference, or the equivalent region reference for a range. */
export function corpusRegionReference(options: CorpusWorkbookOptions): string {
  const resolved = resolveOptions(options);
  const lastRow = resolved.totalsRow ? CORPUS_TOTALS_ROW : CORPUS_LAST_DATA_ROW;
  return `A${CORPUS_HEADER_ROW}:F${lastRow}`;
}

function doubledCell(
  resolved: ResolvedOptions,
  entry: CorpusDataRow,
): CellSpec {
  const doubled = String(entry.amount * 2);
  if (resolved.formulas === "none") {
    return { ref: `E${entry.row}`, value: doubled };
  }
  const formula =
    resolved.formulas === "structured"
      ? `${CORPUS_TABLE_NAME}[[#This Row],[Amount]]*2`
      : `D${entry.row}*2`;
  return { ref: `E${entry.row}`, formula, value: doubled };
}

function ratioCell(resolved: ResolvedOptions, entry: CorpusDataRow): CellSpec {
  const ratio = String(entry.amount / 2);
  if (!resolved.sharedFormula) {
    return { ref: `F${entry.row}`, value: ratio };
  }
  const isMaster = entry.row === CORPUS_FIRST_DATA_ROW;
  return {
    ref: `F${entry.row}`,
    formula: isMaster ? `D${entry.row}/2` : "",
    formulaAttributes: isMaster
      ? ` t="shared" ref="F${CORPUS_FIRST_DATA_ROW}:F${CORPUS_LAST_DATA_ROW}" si="0"`
      : ' t="shared" si="0"',
    value: ratio,
  };
}

function aggregateFormula(resolved: ResolvedOptions, column: string): string {
  return resolved.formulas === "structured"
    ? `SUBTOTAL(109,${CORPUS_TABLE_NAME}[Amount])`
    : `SUM(${column}${CORPUS_FIRST_DATA_ROW}:${column}${CORPUS_LAST_DATA_ROW})`;
}

function buildDataSheet(
  resolved: ResolvedOptions,
  strings: StringTable,
  relationships: RelationshipTable,
): string {
  const rows: string[] = [];

  rows.push(
    rowXml(1, [
      {
        ref: "A1",
        style: 1,
        type: "s",
        value: String(strings.index("Corpus allocation report")),
      },
    ]),
  );

  rows.push(
    rowXml(CORPUS_HEADER_ROW, [
      {
        ref: "A3",
        style: 1,
        type: "s",
        value: String(strings.index("Record")),
      },
      {
        ref: "B3",
        style: 1,
        type: "s",
        value: String(strings.index("Client")),
      },
      {
        ref: "C3",
        style: 1,
        type: "s",
        value: String(strings.index(CORPUS_SPLIT_COLUMN)),
      },
      {
        ref: "D3",
        style: 1,
        type: "s",
        value: String(strings.index("Amount")),
      },
      {
        ref: "E3",
        style: 1,
        type: "s",
        value: String(strings.index("Doubled")),
      },
      { ref: "F3", style: 1, type: "s", value: String(strings.index("Ratio")) },
    ]),
  );

  for (const entry of CORPUS_ROWS) {
    const cells: CellSpec[] = [
      { ref: `A${entry.row}`, value: String(entry.record) },
      {
        ref: `B${entry.row}`,
        type: "s",
        value: String(strings.index(entry.client)),
      },
      {
        ref: `C${entry.row}`,
        type: "s",
        value: String(strings.index(entry.group)),
      },
      { ref: `D${entry.row}`, value: String(entry.amount) },
      doubledCell(resolved, entry),
      ratioCell(resolved, entry),
    ];
    if (resolved.arrayFormula && entry.row === CORPUS_FIRST_DATA_ROW) {
      cells.push({
        ref: `G${entry.row}`,
        formula: `SUM(D${CORPUS_FIRST_DATA_ROW}:D${CORPUS_LAST_DATA_ROW})`,
        formulaAttributes: ` t="array" ref="G${entry.row}:G${entry.row}"`,
        value: String(CORPUS_TOTAL_AMOUNT),
      });
    }
    // Cells outside the region on a row that a split must remove whole.
    if (entry.row === 6) {
      cells.push({
        ref: "H6",
        type: "s",
        value: String(strings.index(CORPUS_SIDE_NOTE)),
      });
      cells.push({ ref: "I6" });
    }
    rows.push(rowXml(entry.row, cells));
  }

  if (resolved.totalsRow) {
    rows.push(
      rowXml(CORPUS_TOTALS_ROW, [
        {
          ref: `A${CORPUS_TOTALS_ROW}`,
          type: "s",
          value: String(strings.index("Total")),
        },
        {
          ref: `D${CORPUS_TOTALS_ROW}`,
          formula:
            resolved.formulas === "none"
              ? undefined
              : aggregateFormula(resolved, "D"),
          value: String(CORPUS_TOTAL_AMOUNT),
        },
      ]),
    );
  }

  if (resolved.footerBlock) {
    rows.push(
      rowXml(CORPUS_FOOTER_ROW, [
        {
          ref: `A${CORPUS_FOOTER_ROW}`,
          type: "s",
          value: String(strings.index("Footer note")),
        },
        {
          ref: `B${CORPUS_FOOTER_ROW}`,
          formula:
            resolved.formulas === "none"
              ? undefined
              : aggregateFormula(resolved, "D"),
          value: String(CORPUS_TOTAL_AMOUNT),
        },
      ]),
    );
  }

  const lastRow = resolved.footerBlock
    ? CORPUS_FOOTER_ROW
    : resolved.totalsRow
      ? CORPUS_TOTALS_ROW
      : CORPUS_LAST_DATA_ROW;
  const sections: string[] = [
    `<dimension ref="A1:I${lastRow}"/>`,
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews>',
    '<sheetFormatPr defaultRowHeight="15"/>',
    '<cols><col min="1" max="6" width="16" customWidth="1"/></cols>',
    `<sheetData>${rows.join("")}</sheetData>`,
  ];

  if (resolved.dependents) {
    // Merged ranges: one anchored above the data, one on a data row that a
    // filtered split moves upward.
    sections.push(
      '<mergeCells count="2"><mergeCell ref="A1:F1"/><mergeCell ref="H6:I6"/></mergeCells>',
    );
    sections.push(
      `<conditionalFormatting sqref="D${CORPUS_FIRST_DATA_ROW}:D${CORPUS_LAST_DATA_ROW}"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>25</formula></cfRule></conditionalFormatting>`,
    );
    sections.push(
      `<dataValidations count="1"><dataValidation type="whole" operator="between" allowBlank="1" sqref="A${CORPUS_FIRST_DATA_ROW}:A${CORPUS_LAST_DATA_ROW}"><formula1>1</formula1><formula2>999</formula2></dataValidation></dataValidations>`,
    );
    const hyperlinkId = relationships.add(
      `${RELATIONSHIP_TYPE}/hyperlink`,
      "https://example.com/record/3",
      true,
    );
    sections.push(
      `<hyperlinks><hyperlink ref="A6" r:id="${hyperlinkId}"/></hyperlinks>`,
    );
  }

  sections.push(
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
  );

  if (resolved.comments) {
    const drawingId = relationships.add(
      `${RELATIONSHIP_TYPE}/vmlDrawing`,
      "../drawings/vmlDrawing1.vml",
    );
    relationships.add(`${RELATIONSHIP_TYPE}/comments`, "../comments1.xml");
    sections.push(`<legacyDrawing r:id="${drawingId}"/>`);
  }

  if (resolved.shape === "table") {
    const tableId = relationships.add(
      `${RELATIONSHIP_TYPE}/table`,
      "../tables/table1.xml",
    );
    sections.push(
      `<tableParts count="1"><tablePart r:id="${tableId}"/></tableParts>`,
    );
  }

  return `${XML_DECLARATION}<worksheet xmlns="${NAMESPACE_MAIN}" xmlns:r="${NAMESPACE_RELATIONSHIPS}">${sections.join("")}</worksheet>`;
}

function buildTablePart(resolved: ResolvedOptions): string {
  const lastRow = resolved.totalsRow ? CORPUS_TOTALS_ROW : CORPUS_LAST_DATA_ROW;
  const totalsAttributes = resolved.totalsRow
    ? ' totalsRowCount="1"'
    : ' totalsRowShown="0"';
  const columns = [
    "Record",
    "Client",
    CORPUS_SPLIT_COLUMN,
    "Amount",
    "Doubled",
    "Ratio",
  ]
    .map((name, index) => {
      if (name === "Amount" && resolved.totalsRow) {
        return `<tableColumn id="${index + 1}" name="${name}" totalsRowFunction="sum"/>`;
      }
      if (name === "Doubled" && resolved.formulas === "structured") {
        return `<tableColumn id="${index + 1}" name="${name}"><calculatedColumnFormula>${escapeXml(
          `${CORPUS_TABLE_NAME}[[#This Row],[Amount]]*2`,
        )}</calculatedColumnFormula></tableColumn>`;
      }
      return `<tableColumn id="${index + 1}" name="${name}"/>`;
    })
    .join("");
  return `${XML_DECLARATION}<table xmlns="${NAMESPACE_MAIN}" id="1" name="${CORPUS_TABLE_NAME}" displayName="${CORPUS_TABLE_NAME}" ref="A${CORPUS_HEADER_ROW}:F${lastRow}" headerRowCount="1"${totalsAttributes}><autoFilter ref="A${CORPUS_HEADER_ROW}:F${CORPUS_LAST_DATA_ROW}"/><tableColumns count="6">${columns}</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
}

function buildSummarySheet(
  resolved: ResolvedOptions,
  strings: StringTable,
  relationships: RelationshipTable,
): string {
  const rows: string[] = [];
  rows.push(
    rowXml(1, [
      {
        ref: "A1",
        style: 1,
        type: "s",
        value: String(strings.index("Portfolio summary")),
      },
    ]),
  );
  rows.push(
    rowXml(2, [
      {
        ref: "A2",
        type: "s",
        value: String(strings.index("Total across every record")),
      },
      {
        ref: "B2",
        formula: `SUM(${CORPUS_SHEET}!D${CORPUS_FIRST_DATA_ROW}:D${CORPUS_LAST_DATA_ROW})`,
        value: String(CORPUS_TOTAL_AMOUNT),
      },
    ]),
  );
  rows.push(
    rowXml(3, [
      { ref: "A3", type: "s", value: String(strings.index("First amount")) },
      {
        ref: "B3",
        formula: `${CORPUS_SHEET}!D${CORPUS_FIRST_DATA_ROW}`,
        value: String(CORPUS_ROWS[0]?.amount ?? 0),
      },
    ]),
  );
  if (resolved.uncachedFormula) {
    rows.push(
      rowXml(4, [
        {
          ref: "A4",
          type: "s",
          value: String(strings.index("Never recalculated")),
        },
        {
          ref: "B4",
          formula: `SUM(${CORPUS_SHEET}!F${CORPUS_FIRST_DATA_ROW}:F${CORPUS_LAST_DATA_ROW})`,
        },
      ]),
    );
  }

  const sections: string[] = [];
  if (resolved.pivot) {
    rows.push(
      rowXml(6, [
        {
          ref: "A6",
          style: 1,
          type: "s",
          value: String(strings.index("Row Labels")),
        },
        {
          ref: "B6",
          style: 1,
          type: "s",
          value: String(strings.index("Sum of Amount")),
        },
      ]),
    );
    const totals = new Map<string, number>();
    for (const entry of CORPUS_ROWS) {
      totals.set(entry.group, (totals.get(entry.group) ?? 0) + entry.amount);
    }
    let pivotRow = 7;
    for (const [group, total] of totals) {
      rows.push(
        rowXml(pivotRow, [
          {
            ref: `A${pivotRow}`,
            type: "s",
            value: String(strings.index(group)),
          },
          { ref: `B${pivotRow}`, value: String(total) },
        ]),
      );
      pivotRow += 1;
    }
    rows.push(
      rowXml(pivotRow, [
        {
          ref: `A${pivotRow}`,
          type: "s",
          value: String(strings.index("Grand Total")),
        },
        { ref: `B${pivotRow}`, value: String(CORPUS_TOTAL_AMOUNT) },
      ]),
    );
  }

  const lastRow = resolved.pivot ? 11 : resolved.uncachedFormula ? 4 : 3;
  sections.push(`<dimension ref="A1:B${lastRow}"/>`);
  sections.push('<sheetViews><sheetView workbookViewId="0"/></sheetViews>');
  sections.push('<sheetFormatPr defaultRowHeight="15"/>');
  sections.push(`<sheetData>${rows.join("")}</sheetData>`);
  sections.push(
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
  );

  if (resolved.pivot) {
    relationships.add(
      `${RELATIONSHIP_TYPE}/pivotTable`,
      "../pivotTables/pivotTable1.xml",
    );
  }

  return `${XML_DECLARATION}<worksheet xmlns="${NAMESPACE_MAIN}" xmlns:r="${NAMESPACE_RELATIONSHIPS}">${sections.join("")}</worksheet>`;
}

function buildHiddenSheet(strings: StringTable): string {
  const rows = [
    rowXml(1, [
      {
        ref: "A1",
        style: 1,
        type: "s",
        value: String(strings.index("Record")),
      },
      {
        ref: "B1",
        style: 1,
        type: "s",
        value: String(strings.index(CORPUS_SPLIT_COLUMN)),
      },
    ]),
    rowXml(2, [
      { ref: "A2", value: "7" },
      { ref: "B2", type: "s", value: String(strings.index("Alpha")) },
    ]),
    rowXml(3, [
      { ref: "A3", value: "8" },
      { ref: "B3", type: "s", value: String(strings.index("Beta")) },
    ]),
    rowXml(4, [
      { ref: "A4", value: "9" },
      { ref: "B4", type: "s", value: String(strings.index("Alpha")) },
    ]),
  ];
  return `${XML_DECLARATION}<worksheet xmlns="${NAMESPACE_MAIN}" xmlns:r="${NAMESPACE_RELATIONSHIPS}"><dimension ref="A1:B4"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>${rows.join(
    "",
  )}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
}

function buildVeryHiddenSheet(strings: StringTable): string {
  const rows = [
    rowXml(1, [
      {
        ref: "A1",
        style: 1,
        type: "s",
        value: String(strings.index("Archive note")),
      },
    ]),
    rowXml(2, [
      {
        ref: "A2",
        type: "s",
        value: String(strings.index("Retained without filtering")),
      },
    ]),
  ];
  return `${XML_DECLARATION}<worksheet xmlns="${NAMESPACE_MAIN}" xmlns:r="${NAMESPACE_RELATIONSHIPS}"><dimension ref="A1:A2"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>${rows.join(
    "",
  )}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
}

function buildStyles(): string {
  return `${XML_DECLARATION}<styleSheet xmlns="${NAMESPACE_MAIN}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font></dxf></dxfs></styleSheet>`;
}

function buildComments(): string {
  return `${XML_DECLARATION}<comments xmlns="${NAMESPACE_MAIN}"><authors><author>ConsultChimps</author></authors><commentList><comment ref="B6" authorId="0"><text><r><t>${escapeXml(
    CORPUS_COMMENT_TEXT,
  )}</t></r></text></comment></commentList></comments>`;
}

function buildVmlDrawing(): string {
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype><v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"/><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill>False</x:AutoFill><x:Row>5</x:Row><x:Column>1</x:Column></x:ClientData></v:shape></xml>`;
}

function buildPivotCacheDefinition(): string {
  const groups = [...new Set(CORPUS_ROWS.map((entry) => entry.group))];
  const clients = CORPUS_ROWS.map((entry) => entry.client);
  const sharedGroups = groups
    .map((group) => `<s v="${escapeXml(group)}"/>`)
    .join("");
  const sharedClients = clients
    .map((client) => `<s v="${escapeXml(client)}"/>`)
    .join("");
  return `${XML_DECLARATION}<pivotCacheDefinition xmlns="${NAMESPACE_MAIN}" xmlns:r="${NAMESPACE_RELATIONSHIPS}" r:id="rId1" refreshOnLoad="0" refreshedBy="ConsultChimps" recordCount="${CORPUS_ROWS.length}" createdVersion="3" refreshedVersion="3" minRefreshableVersion="3"><cacheSource type="worksheet"><worksheetSource ref="A${CORPUS_HEADER_ROW}:F${CORPUS_LAST_DATA_ROW}" sheet="${CORPUS_SHEET}"/></cacheSource><cacheFields count="3"><cacheField name="${CORPUS_SPLIT_COLUMN}" numFmtId="0"><sharedItems count="${groups.length}">${sharedGroups}</sharedItems></cacheField><cacheField name="Client" numFmtId="0"><sharedItems count="${clients.length}">${sharedClients}</sharedItems></cacheField><cacheField name="Amount" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="10" maxValue="60"/></cacheField></cacheFields></pivotCacheDefinition>`;
}

function buildPivotCacheRecords(): string {
  // Records carry literal per-row values rather than shared-item indexes so a
  // confidentiality assertion can grep one group's output for another's data.
  const records = CORPUS_ROWS.map(
    (entry) =>
      `<r><s v="${escapeXml(entry.group)}"/><s v="${escapeXml(
        entry.client,
      )}"/><n v="${entry.amount}"/></r>`,
  ).join("");
  return `${XML_DECLARATION}<pivotCacheRecords xmlns="${NAMESPACE_MAIN}" xmlns:r="${NAMESPACE_RELATIONSHIPS}" count="${CORPUS_ROWS.length}">${records}</pivotCacheRecords>`;
}

function buildPivotTable(): string {
  const groups = [...new Set(CORPUS_ROWS.map((entry) => entry.group))];
  const items = groups
    .map((_group, index) => `<item x="${index}"/>`)
    .concat('<item t="grand"/>')
    .join("");
  const rowItems = groups
    .map((_group, index) => `<i><x v="${index}"/></i>`)
    .concat('<i t="grand"><x/></i>')
    .join("");
  return `${XML_DECLARATION}<pivotTableDefinition xmlns="${NAMESPACE_MAIN}" name="PivotSummary" cacheId="1" dataOnRows="0" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" createdVersion="3" updatedVersion="3" minRefreshableVersion="3" itemPrintTitles="1" useAutoFormatting="1" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"><location ref="A6:B11" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="3"><pivotField axis="axisRow" showAll="0"><items count="${
    groups.length + 1
  }">${items}</items></pivotField><pivotField showAll="0"/><pivotField dataField="1" showAll="0"/></pivotFields><rowFields count="1"><field x="0"/></rowFields><rowItems count="${
    groups.length + 1
  }">${rowItems}</rowItems><colItems count="1"><i/></colItems><dataFields count="1"><dataField name="Sum of Amount" fld="2" baseField="0" baseItem="0"/></dataFields><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`;
}

function buildCalcChain(resolved: ResolvedOptions): string {
  const entries: string[] = [];
  for (const entry of CORPUS_ROWS) {
    entries.push(`<c r="E${entry.row}" i="1"/>`);
    if (resolved.sharedFormula) {
      entries.push(`<c r="F${entry.row}" i="1"/>`);
    }
  }
  if (resolved.arrayFormula) {
    entries.push(`<c r="G${CORPUS_FIRST_DATA_ROW}" i="1"/>`);
  }
  if (resolved.totalsRow) {
    entries.push(`<c r="D${CORPUS_TOTALS_ROW}" i="1"/>`);
  }
  if (resolved.footerBlock) {
    entries.push(`<c r="B${CORPUS_FOOTER_ROW}" i="1"/>`);
  }
  if (resolved.summarySheet) {
    entries.push('<c r="B2" i="2"/>');
    entries.push('<c r="B3" i="2"/>');
  }
  return `${XML_DECLARATION}<calcChain xmlns="${NAMESPACE_MAIN}">${entries.join("")}</calcChain>`;
}

interface SheetEntry {
  name: string;
  partName: string;
  state?: string | undefined;
  xml: string;
  relationships: RelationshipTable;
}

/**
 * Build one corpus workbook package. Options compose: every structure can be
 * switched on independently so a test can isolate the one it pins.
 */
export async function buildCorpusWorkbook(
  options: CorpusWorkbookOptions,
): Promise<Uint8Array> {
  const resolved = resolveOptions(options);
  const strings = new StringTable();
  const sheets: SheetEntry[] = [];

  const dataRelationships = new RelationshipTable();
  const dataXml = buildDataSheet(resolved, strings, dataRelationships);
  sheets.push({
    name: CORPUS_SHEET,
    partName: CORPUS_PARTS.dataSheet,
    xml: dataXml,
    relationships: dataRelationships,
  });

  if (resolved.summarySheet) {
    const summaryRelationships = new RelationshipTable();
    const summaryXml = buildSummarySheet(
      resolved,
      strings,
      summaryRelationships,
    );
    sheets.push({
      name: CORPUS_SUMMARY_SHEET,
      partName: CORPUS_PARTS.summarySheet,
      xml: summaryXml,
      relationships: summaryRelationships,
    });
  }

  if (resolved.hiddenSheets) {
    sheets.push({
      name: CORPUS_HIDDEN_SHEET,
      partName: CORPUS_PARTS.hiddenSheet,
      state: "hidden",
      xml: buildHiddenSheet(strings),
      relationships: new RelationshipTable(),
    });
    sheets.push({
      name: CORPUS_VERY_HIDDEN_SHEET,
      partName: CORPUS_PARTS.veryHiddenSheet,
      state: "veryHidden",
      xml: buildVeryHiddenSheet(strings),
      relationships: new RelationshipTable(),
    });
  }

  const workbookRelationships = new RelationshipTable();
  const sheetElements = sheets
    .map((sheet, index) => {
      const relationshipId = workbookRelationships.add(
        `${RELATIONSHIP_TYPE}/worksheet`,
        sheet.partName.replace(/^xl\//u, ""),
      );
      const state = sheet.state ? ` state="${sheet.state}"` : "";
      return `<sheet name="${escapeXml(sheet.name)}" sheetId="${
        index + 1
      }"${state} r:id="${relationshipId}"/>`;
    })
    .join("");
  workbookRelationships.add(`${RELATIONSHIP_TYPE}/styles`, "styles.xml");
  workbookRelationships.add(
    `${RELATIONSHIP_TYPE}/sharedStrings`,
    "sharedStrings.xml",
  );
  if (resolved.calcChain) {
    workbookRelationships.add(
      `${RELATIONSHIP_TYPE}/calcChain`,
      "calcChain.xml",
    );
  }
  let pivotCacheId: string | undefined;
  if (resolved.pivot) {
    pivotCacheId = workbookRelationships.add(
      `${RELATIONSHIP_TYPE}/pivotCacheDefinition`,
      "pivotCache/pivotCacheDefinition1.xml",
    );
  }
  if (resolved.macro) {
    workbookRelationships.add(VBA_RELATIONSHIP_TYPE, "vbaProject.bin");
  }

  const definedNames = resolved.definedNames
    ? `<definedNames><definedName name="${CORPUS_RANGE_NAME}">${CORPUS_SHEET}!$A$${CORPUS_HEADER_ROW}:$F$${CORPUS_LAST_DATA_ROW}</definedName><definedName name="${CORPUS_LOCAL_NAME}" localSheetId="0">${CORPUS_SHEET}!$A$${CORPUS_FOOTER_ROW}</definedName></definedNames>`
    : "";
  const pivotCaches = pivotCacheId
    ? `<pivotCaches><pivotCache cacheId="1" r:id="${pivotCacheId}"/></pivotCaches>`
    : "";
  const workbookXml = `${XML_DECLARATION}<workbook xmlns="${NAMESPACE_MAIN}" xmlns:r="${NAMESPACE_RELATIONSHIPS}"><sheets>${sheetElements}</sheets>${definedNames}${pivotCaches}</workbook>`;

  const archive = new JSZip();
  const write = (partName: string, content: string | Uint8Array): void => {
    archive.file(partName, content, {
      date: FIXED_ENTRY_DATE,
      createFolders: false,
    });
  };

  const workbookContentType = resolved.macro
    ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
    : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
  const overrides: string[] = [
    `<Override PartName="/xl/workbook.xml" ContentType="${workbookContentType}"/>`,
    ...sheets.map(
      (sheet) =>
        `<Override PartName="/${sheet.partName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    ),
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
  ];
  const defaults: string[] = [
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
  ];

  if (resolved.shape === "table") {
    write(CORPUS_PARTS.table, buildTablePart(resolved));
    overrides.push(
      '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>',
    );
  }
  if (resolved.comments) {
    write(CORPUS_PARTS.comments, buildComments());
    write("xl/drawings/vmlDrawing1.vml", buildVmlDrawing());
    overrides.push(
      '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>',
    );
    defaults.push(
      '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>',
    );
  }
  if (resolved.calcChain) {
    write(CORPUS_PARTS.calcChain, buildCalcChain(resolved));
    overrides.push(
      '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>',
    );
  }
  if (resolved.pivot) {
    write(CORPUS_PARTS.pivotTable, buildPivotTable());
    write(CORPUS_PARTS.pivotCacheDefinition, buildPivotCacheDefinition());
    write(CORPUS_PARTS.pivotCacheRecords, buildPivotCacheRecords());
    const pivotTableRelationships = new RelationshipTable();
    pivotTableRelationships.add(
      `${RELATIONSHIP_TYPE}/pivotCacheDefinition`,
      "../pivotCache/pivotCacheDefinition1.xml",
    );
    write(
      "xl/pivotTables/_rels/pivotTable1.xml.rels",
      pivotTableRelationships.xml(),
    );
    const cacheRelationships = new RelationshipTable();
    cacheRelationships.add(
      `${RELATIONSHIP_TYPE}/pivotCacheRecords`,
      "pivotCacheRecords1.xml",
    );
    write(
      "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
      cacheRelationships.xml(),
    );
    overrides.push(
      '<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>',
      '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>',
      '<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>',
    );
  }
  if (resolved.macro) {
    write(CORPUS_PARTS.vbaProject, Uint8Array.from(CORPUS_VBA_BYTES));
    defaults.push(
      '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>',
    );
  }

  const packageRelationships = new RelationshipTable();
  packageRelationships.add(
    `${RELATIONSHIP_TYPE}/officeDocument`,
    "xl/workbook.xml",
  );
  write("_rels/.rels", packageRelationships.xml());
  write(CORPUS_PARTS.workbook, workbookXml);
  write("xl/_rels/workbook.xml.rels", workbookRelationships.xml());
  write("xl/styles.xml", buildStyles());
  for (const sheet of sheets) {
    write(sheet.partName, sheet.xml);
    if (sheet.relationships.size > 0) {
      const name = sheet.partName.replace(
        /^xl\/worksheets\//u,
        "xl/worksheets/_rels/",
      );
      write(`${name}.rels`, sheet.relationships.xml());
    }
  }
  // The string table is only complete once every worksheet has been built.
  write(CORPUS_PARTS.sharedStrings, strings.xml());
  write(
    CORPUS_PARTS.contentTypes,
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults.join(
      "",
    )}${overrides.join("")}</Types>`,
  );

  return archive.generateAsync({
    compression: "DEFLATE",
    type: "uint8array",
  });
}

/** Read one package part of a corpus workbook (or of an operation's output) as text. */
export async function readPackagePart(
  workbookBytes: Uint8Array,
  partName: string,
): Promise<string> {
  const archive = await JSZip.loadAsync(workbookBytes);
  const entry = archive.file(partName);
  if (!entry) {
    throw new Error(`Workbook package part is missing: ${partName}`);
  }
  return entry.async("text");
}

/** Whether a package part exists, without reading it. */
export async function hasPackagePart(
  workbookBytes: Uint8Array,
  partName: string,
): Promise<boolean> {
  const archive = await JSZip.loadAsync(workbookBytes);
  return archive.file(partName) !== null;
}

/**
 * Every file part path in a workbook package, sorted. Directory entries are
 * excluded: JSZip adds them when it regenerates a package but not when the
 * corpus writes one, and that difference is not a behaviour worth pinning.
 */
export async function packagePartNames(
  workbookBytes: Uint8Array,
): Promise<string[]> {
  const archive = await JSZip.loadAsync(workbookBytes);
  return Object.values(archive.files)
    .filter((entry) => !entry.dir)
    .map((entry) => entry.name)
    .sort();
}

const corpusDirectories: string[] = [];

/** Create a temporary directory registered for `cleanupCorpusDirectories`. */
export async function createCorpusDirectory(): Promise<string> {
  const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-corpus-"));
  corpusDirectories.push(directory);
  return directory;
}

/** Remove every directory created by `createCorpusDirectory` in this test file. */
export async function cleanupCorpusDirectories(): Promise<void> {
  await Promise.all(
    corpusDirectories
      .splice(0)
      .map((directory) => rm(directory, { force: true, recursive: true })),
  );
}

/** Build a corpus workbook and write it to `directory/fileName`. */
export async function writeCorpusWorkbook(
  directory: string,
  fileName: string,
  options: CorpusWorkbookOptions,
): Promise<string> {
  const filePath = path.join(directory, fileName);
  await writeFile(filePath, await buildCorpusWorkbook(options));
  return filePath;
}

/** Read a workbook produced by an operation back into bytes. */
export async function readWorkbookBytes(filePath: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(filePath));
}

/** The row numbers still present in a worksheet part, in document order. */
export function worksheetRowNumbers(worksheetXml: string): number[] {
  return [...worksheetXml.matchAll(/<row\b[^>]*\br="(\d+)"/gu)].map((match) =>
    Number(match[1]),
  );
}

/** Every `<c r="...">` reference in a worksheet part, in document order. */
export function worksheetCellReferences(worksheetXml: string): string[] {
  return [...worksheetXml.matchAll(/<c\b[^>]*\br="([A-Z]+\d+)"/gu)].map(
    (match) => match[1] ?? "",
  );
}

/** The formula text of one cell in a worksheet part, if the cell has one. */
export function worksheetCellFormula(
  worksheetXml: string,
  reference: string,
): string | undefined {
  const pattern = new RegExp(
    `<c\\b[^>]*\\br="${reference}"[^>]*>([\\s\\S]*?)</c>`,
    "u",
  );
  const body = pattern.exec(worksheetXml)?.[1];
  return /<f\b[^>]*>([\s\S]*?)<\/f>/u.exec(body ?? "")?.[1];
}

/** The cached value of one cell in a worksheet part, if the cell has one. */
export function worksheetCellValue(
  worksheetXml: string,
  reference: string,
): string | undefined {
  const pattern = new RegExp(
    `<c\\b[^>]*\\br="${reference}"[^>]*>([\\s\\S]*?)</c>`,
    "u",
  );
  const body = pattern.exec(worksheetXml)?.[1];
  return /<v>([\s\S]*?)<\/v>/u.exec(body ?? "")?.[1];
}

/** The `sqref` of the first `<conditionalFormatting>` block in a worksheet part. */
export function conditionalFormattingSqref(
  worksheetXml: string,
): string | undefined {
  return /<conditionalFormatting\b[^>]*\bsqref="([^"]+)"/u.exec(
    worksheetXml,
  )?.[1];
}

/** The `sqref` of the first `<dataValidation>` in a worksheet part. */
export function dataValidationSqref(worksheetXml: string): string | undefined {
  return /<dataValidation\b[^>]*\bsqref="([^"]+)"/u.exec(worksheetXml)?.[1];
}

/** Every `<mergeCell ref="...">` in a worksheet part, in document order. */
export function mergedCellReferences(worksheetXml: string): string[] {
  return [...worksheetXml.matchAll(/<mergeCell\b[^>]*\bref="([^"]+)"/gu)].map(
    (match) => match[1] ?? "",
  );
}

/** Every `<hyperlink ref="...">` in a worksheet part, in document order. */
export function hyperlinkReferences(worksheetXml: string): string[] {
  return [...worksheetXml.matchAll(/<hyperlink\b[^>]*\bref="([^"]+)"/gu)].map(
    (match) => match[1] ?? "",
  );
}

/** Every `<c r="...">` entry in a calculation chain part, in document order. */
export function calcChainReferences(calcChainXml: string): string[] {
  return [...calcChainXml.matchAll(/<c\b[^>]*\br="([A-Z]+\d+)"/gu)].map(
    (match) => match[1] ?? "",
  );
}
