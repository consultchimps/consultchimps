/**
 * L1 - the workbook document model.
 *
 * `WorkbookModel` owns the worksheets, the tables registry, defined names,
 * shared strings and the calculation chain, and it is the only place row edits
 * happen. A caller cannot delete rows and forget one of the structures that
 * describe them, because there is no other way in.
 *
 * Only parts the model actually changed are written back. Everything else is
 * handed to L0 exactly as it arrived, which is what keeps untouched parts
 * byte-identical through an edit.
 */
import * as XLSX from "xlsx";

import {
  readExcelTableDefinitionsFrom,
  readWorkbookSheetsFrom,
  type ExcelTableDefinition,
  type WorkbookSheetEntry,
} from "../excel-tables.js";
import { WorkbookPackage } from "../package/index.js";
import {
  relocateReference,
  DELETED_REFERENCE,
  type RowRelocation,
} from "./references.js";
import type {
  CellRange,
  DefinedNameEntry,
  DeleteRowsReport,
  RowNumber,
  SheetInfo,
  WorkbookModel as WorkbookModelContract,
  WorkbookTableInfo,
} from "./types.js";
import {
  WorksheetModel,
  type RelocationCounters,
  type WorksheetHost,
} from "./worksheet-model.js";
import { excelSerialToDate, StyleTable } from "./styles.js";
import {
  addAttribute,
  decodeXmlText,
  editElements,
  findElement,
  getAttribute,
  setAttribute,
} from "./xml.js";

const WORKBOOK_PART = "xl/workbook.xml";
const CALC_CHAIN_PART = "xl/calcChain.xml";
const SHARED_STRINGS_PART = "xl/sharedStrings.xml";
const STYLES_PART = "xl/styles.xml";

export interface DeleteRowsOptions {
  /** Close the gaps left behind so surviving rows form a contiguous block. */
  renumber: boolean;
  /**
   * Excel Tables whose `ref` and `autoFilter` follow the edit. Defaults to
   * every table anchored on the worksheet. Pass an empty list for a binding
   * that deliberately leaves a table claiming its original range.
   */
  tables?: readonly ExcelTableDefinition[] | undefined;
}

function visibilityOf(state: string | undefined): SheetInfo["visibility"] {
  return state === "hidden" || state === "veryHidden" ? state : "visible";
}

function decodeTableRange(reference: string): CellRange {
  const range = XLSX.utils.decode_range(reference);
  return {
    start: { row: range.s.r + 1, column: range.s.c },
    end: { row: range.e.r + 1, column: range.e.c },
  };
}

export class WorkbookModel implements WorkbookModelContract, WorksheetHost {
  readonly #package: WorkbookPackage;
  readonly #entries: readonly WorkbookSheetEntry[];
  readonly #tables: readonly ExcelTableDefinition[];
  readonly #worksheets = new Map<string, WorksheetModel>();
  readonly #changedWorksheets = new Set<string>();
  /** Tables the current edit is allowed to resize, when the caller narrowed it. */
  #tableScope: readonly ExcelTableDefinition[] | undefined;
  #calcChain: CalcChainModel | undefined;
  #calcChainChanged = false;
  #sourceBytes: Uint8Array | undefined;
  #values: XLSX.WorkBook | undefined;
  #strings: string[] | undefined;
  #styles: StyleTable | undefined;
  #date1904: boolean | undefined;

  private constructor(
    workbookPackage: WorkbookPackage,
    entries: readonly WorkbookSheetEntry[],
    tables: readonly ExcelTableDefinition[],
  ) {
    this.#package = workbookPackage;
    this.#entries = entries;
    this.#tables = tables;
  }

  static async load(
    bytes: Uint8Array,
    valuesView?: XLSX.WorkBook | undefined,
  ): Promise<WorkbookModel> {
    const model = WorkbookModel.fromPackage(await WorkbookPackage.load(bytes));
    model.#sourceBytes = bytes;
    model.#values = valuesView;
    return model;
  }

  static fromPackage(workbookPackage: WorkbookPackage): WorkbookModel {
    return new WorkbookModel(
      workbookPackage,
      readWorkbookSheetsFrom(workbookPackage),
      readExcelTableDefinitionsFrom(workbookPackage),
    );
  }

  // -- The L1 workbook seam ------------------------------------------------

  get sheets(): readonly SheetInfo[] {
    return this.#entries.map((entry) => ({
      name: entry.name,
      partPath: entry.worksheetPart,
      visibility: visibilityOf(entry.state),
    }));
  }

  /** The parsed worksheet for a sheet name, parsed once and cached. */
  worksheet(name: string): WorksheetModel | undefined {
    const cached = this.#worksheets.get(name);
    if (cached) {
      return cached;
    }
    const entry = this.#entries.find((candidate) => candidate.name === name);
    if (!entry) {
      return undefined;
    }
    const model = WorksheetModel.parse(
      this.#package.requireText(entry.worksheetPart),
      {
        name: entry.name,
        partPath: entry.worksheetPart,
        visibility: visibilityOf(entry.state),
      },
      this,
    );
    this.#worksheets.set(name, model);
    return model;
  }

  /** Workbook-scoped and sheet-scoped defined names. */
  definedNames(): Promise<readonly DefinedNameEntry[]> {
    const xml = this.#package.readText(WORKBOOK_PART);
    if (xml === undefined) {
      return Promise.resolve([]);
    }
    const names: DefinedNameEntry[] = [];
    editElements(xml, "definedName", (element, text) => {
      const name = getAttribute(element.openTag, "name");
      if (name !== undefined) {
        const localSheetId = getAttribute(element.openTag, "localSheetId");
        names.push({
          name,
          localSheetId:
            localSheetId === undefined ? undefined : Number(localSheetId),
          reference: text.slice(
            element.innerStart - element.start,
            element.innerEnd - element.start,
          ),
        });
      }
      return text;
    });
    return Promise.resolve(names);
  }

  tables(): Promise<readonly WorkbookTableInfo[]> {
    return Promise.resolve(
      this.#tables.map((table) => {
        const range = decodeTableRange(table.range);
        return {
          name: table.name,
          sheetName: table.sheet,
          partPath: table.tablePart,
          range,
          // A table with `headerRowCount="0"` has no header row at all. The
          // seam types this as a row number, so headerless reports 0: callers
          // treat "below 1" as "do not associate a detected header with me".
          headerRow: table.headerRow ? range.start.row : 0,
          totalsRow: table.totalsRow,
          columnNames: table.columns,
        };
      }),
    );
  }

  /**
   * Write back the parts the model changed and serialize the package. Parts no
   * edit reached are handed to L0 with their original bytes.
   */
  async save(): Promise<Uint8Array> {
    this.flush();
    return this.#package.save();
  }

  // -- Extensions the seam does not cover ----------------------------------

  /** The underlying package, for parts no model structure covers yet. */
  get package(): WorkbookPackage {
    return this.#package;
  }

  /** Worksheets as the workbook part declares them, with their sheet ids. */
  get sheetEntries(): readonly WorkbookSheetEntry[] {
    return this.#entries;
  }

  /** Excel Table definitions in the package-level shape the readers produce. */
  get tableDefinitions(): readonly ExcelTableDefinition[] {
    return this.#tables;
  }

  /**
   * A cell-value view of the workbook as loaded. Reading values is where a
   * mature spreadsheet reader earns its keep - number formats, dates, inline
   * and shared strings - so the model reads through one rather than
   * reimplementing it. Every *write* still goes through the model itself.
   */
  values(): XLSX.WorkBook {
    if (!this.#values) {
      if (!this.#sourceBytes) {
        throw new Error(
          "This workbook model was built from a package and has no value view.",
        );
      }
      this.#values = XLSX.read(this.#sourceBytes, {
        cellDates: true,
        dense: false,
        type: "array",
      });
    }
    return this.#values;
  }

  /** The shared string table, read far enough to resolve `t="s"` cells. */
  sharedStrings(): string[] {
    const xml = this.#package.readText(SHARED_STRINGS_PART);
    if (xml === undefined) {
      return [];
    }
    const strings: string[] = [];
    editElements(xml, "si", (element, text) => {
      const inner = text.slice(
        element.innerStart - element.start,
        element.innerEnd - element.start,
      );
      const parts: string[] = [];
      editElements(inner, "t", (run, runText) => {
        parts.push(
          run.selfClosing
            ? ""
            : runText.slice(
                run.innerStart - run.start,
                run.innerEnd - run.start,
              ),
        );
        return runText;
      });
      strings.push(decodeXmlText(parts.join("")));
      return text;
    });
    return strings;
  }

  /**
   * Delete rows from one worksheet through the invariant pass. The `tables`
   * option narrows which Excel Tables follow the edit.
   */
  deleteRows(
    sheetName: string,
    rows: Iterable<RowNumber>,
    options: DeleteRowsOptions,
  ): DeleteRowsReport {
    const worksheet = this.#requireWorksheet(sheetName);
    return this.#withTableScope(options.tables, () =>
      worksheet.deleteRows(new Set(rows), { renumber: options.renumber }),
    );
  }

  /**
   * The general form of `deleteRows`: an explicit relocation plan, for edits
   * that move rows rather than only closing gaps - an Excel Table compacting
   * its data rows while the totals row follows and the rows below stay put.
   * Both funnel into the same invariant pass.
   */
  relocateRows(
    sheetName: string,
    relocation: RowRelocation,
    tables: readonly ExcelTableDefinition[] | undefined = undefined,
  ): DeleteRowsReport {
    const worksheet = this.#requireWorksheet(sheetName);
    return this.#withTableScope(tables, () =>
      worksheet.applyRowRelocation(relocation),
    );
  }

  /** Set an Excel Table's range and autoFilter explicitly. */
  setTableRange(
    table: ExcelTableDefinition,
    reference: string,
    autoFilterReference: string | undefined,
  ): void {
    const xml = this.#package.readText(table.tablePart);
    if (xml === undefined) {
      return;
    }
    let rewritten = editElements(xml, "table", (element, text) =>
      getAttribute(element.openTag, "ref") === undefined
        ? text
        : `${setAttribute(element.openTag, "ref", reference)}${text.slice(
            element.openTag.length,
          )}`,
    );
    if (autoFilterReference !== undefined) {
      rewritten = editElements(rewritten, "autoFilter", (element, text) =>
        getAttribute(element.openTag, "ref") === undefined
          ? text
          : `${setAttribute(
              element.openTag,
              "ref",
              autoFilterReference,
            )}${text.slice(element.openTag.length)}`,
      );
    }
    if (rewritten !== xml) {
      this.#package.writeText(table.tablePart, rewritten);
    }
  }

  /** Write changed parts into the package without serializing it. */
  flush(): void {
    for (const sheetName of this.#changedWorksheets) {
      const entry = this.#entries.find(
        (candidate) => candidate.name === sheetName,
      );
      const model = this.#worksheets.get(sheetName);
      if (entry && model) {
        this.#package.writeText(entry.worksheetPart, model.toXml());
      }
    }
    this.#changedWorksheets.clear();

    if (this.#calcChainChanged && this.#calcChain) {
      if (this.#calcChain.isEmpty) {
        this.#package.removePartAndReferences(CALC_CHAIN_PART, WORKBOOK_PART);
      } else {
        this.#package.writeText(CALC_CHAIN_PART, this.#calcChain.toXml());
      }
      this.#calcChainChanged = false;
    }
  }

  // -- WorksheetHost -------------------------------------------------------

  markWorksheetChanged(sheetName: string): void {
    this.#changedWorksheets.add(sheetName);
  }

  sharedString(index: number): string | undefined {
    this.#strings ??= this.sharedStrings();
    return this.#strings[index];
  }

  isDateStyle(styleIndex: number | undefined): boolean {
    this.#styles ??= StyleTable.parse(this.#package.readText(STYLES_PART));
    return this.#styles.isDateStyle(styleIndex);
  }

  serialToDate(serial: number): Date {
    return excelSerialToDate(serial, this.#usesDate1904());
  }

  /** Whether the workbook counts days from 1904 rather than 1900. */
  #usesDate1904(): boolean {
    if (this.#date1904 === undefined) {
      const xml = this.#package.readText(WORKBOOK_PART) ?? "";
      const properties = findElement(xml, "workbookPr");
      const declared = properties
        ? getAttribute(properties.openTag, "date1904")
        : undefined;
      this.#date1904 = declared === "1" || declared === "true";
    }
    return this.#date1904;
  }

  relocateTables(
    sheetName: string,
    relocation: RowRelocation,
    counters: RelocationCounters,
  ): void {
    const scope = this.#tableScope ?? this.#tables;
    for (const table of scope) {
      if (table.sheet === sheetName) {
        this.#relocateTable(table, relocation, counters);
      }
    }
  }

  relocateCalcChain(
    sheetName: string,
    relocation: RowRelocation,
    counters: RelocationCounters,
  ): void {
    const entry = this.#entries.find(
      (candidate) => candidate.name === sheetName,
    );
    const xml = this.#package.readText(CALC_CHAIN_PART);
    if (!entry || xml === undefined) {
      return;
    }
    this.#calcChain ??= new CalcChainModel(xml);
    const adjusted = this.#calcChain.relocate(entry.sheetId, relocation);
    counters.calcChainEntries += adjusted;
    if (adjusted > 0) {
      this.#calcChainChanged = true;
    }
  }

  /**
   * Move cell-comment anchors with the rows they annotate.
   *
   * A comment is anchored twice: by an A1 reference in the comments part, and
   * by a zero-based row in the legacy VML drawing that draws its box. Both
   * have to move together, or Excel shows the note beside the wrong record.
   */
  relocateComments(
    sheetName: string,
    relocation: RowRelocation,
    counters: RelocationCounters,
  ): void {
    const entry = this.#entries.find(
      (candidate) => candidate.name === sheetName,
    );
    if (!entry) {
      return;
    }

    for (const relationship of this.#package.relationshipsOf(
      entry.worksheetPart,
    )) {
      if (relationship.targetMode === "External") {
        continue;
      }
      const partPath = this.#package.resolvePart(
        entry.worksheetPart,
        relationship.target,
      );
      if (relationship.type.endsWith("/comments")) {
        this.#relocateCommentList(partPath, relocation, counters);
      } else if (relationship.type.endsWith("/vmlDrawing")) {
        this.#relocateVmlAnchors(partPath, relocation, counters);
      }
    }
  }

  #relocateCommentList(
    partPath: string,
    relocation: RowRelocation,
    counters: RelocationCounters,
  ): void {
    const xml = this.#package.readText(partPath);
    if (xml === undefined) {
      return;
    }
    const rewritten = editElements(xml, "comment", (element, text) => {
      const reference = getAttribute(element.openTag, "ref");
      if (reference === undefined) {
        return text;
      }
      const relocated = relocateReference(reference, relocation);
      if (relocated.includes(DELETED_REFERENCE)) {
        counters.comments += 1;
        return undefined;
      }
      if (relocated !== reference) {
        counters.comments += 1;
      }
      return `${setAttribute(element.openTag, "ref", relocated)}${text.slice(
        element.openTag.length,
      )}`;
    });
    if (rewritten !== xml) {
      this.#package.writeText(partPath, rewritten);
    }
  }

  #relocateVmlAnchors(
    partPath: string,
    relocation: RowRelocation,
    counters: RelocationCounters,
  ): void {
    const xml = this.#package.readText(partPath);
    if (xml === undefined) {
      return;
    }
    const rewritten = editElements(xml, "shape", (shape, shapeText) => {
      const anchor = findElement(shapeText, "Row");
      if (!anchor || anchor.selfClosing) {
        return shapeText;
      }
      // The VML anchor counts rows from zero; the relocation plan from one.
      const sourceRow =
        Number(shapeText.slice(anchor.innerStart, anchor.innerEnd)) + 1;
      if (!Number.isInteger(sourceRow)) {
        return shapeText;
      }
      const destination = relocation.target(sourceRow);
      if (destination === null) {
        counters.comments += 1;
        return undefined;
      }
      if (destination === sourceRow) {
        return shapeText;
      }
      counters.comments += 1;
      return `${shapeText.slice(0, anchor.innerStart)}${destination - 1}${shapeText.slice(
        anchor.innerEnd,
      )}`;
    });
    if (rewritten !== xml) {
      this.#package.writeText(partPath, rewritten);
    }
  }

  // -- Internals -----------------------------------------------------------

  #requireWorksheet(sheetName: string): WorksheetModel {
    const worksheet = this.worksheet(sheetName);
    if (!worksheet) {
      throw new Error(`Worksheet "${sheetName}" is not in this workbook.`);
    }
    return worksheet;
  }

  #withTableScope<T>(
    tables: readonly ExcelTableDefinition[] | undefined,
    edit: () => T,
  ): T {
    this.#tableScope = tables;
    try {
      return edit();
    } finally {
      this.#tableScope = undefined;
    }
  }

  #relocateTable(
    table: ExcelTableDefinition,
    relocation: RowRelocation,
    counters: RelocationCounters,
  ): void {
    const xml = this.#package.readText(table.tablePart);
    if (xml === undefined) {
      return;
    }

    let rewritten = editElements(xml, "table", (element, text) => {
      const reference = getAttribute(element.openTag, "ref");
      if (reference === undefined) {
        return text;
      }
      return `${setAttribute(
        element.openTag,
        "ref",
        relocateReference(reference, relocation),
      )}${text.slice(element.openTag.length)}`;
    });
    rewritten = editElements(rewritten, "autoFilter", (element, text) => {
      const reference = getAttribute(element.openTag, "ref");
      if (reference === undefined) {
        return text;
      }
      const relocated = relocateReference(reference, relocation);
      return relocated.includes(DELETED_REFERENCE)
        ? undefined
        : `${setAttribute(element.openTag, "ref", relocated)}${text.slice(
            element.openTag.length,
          )}`;
    });

    if (rewritten !== xml) {
      this.#package.writeText(table.tablePart, rewritten);
      counters.tableRefs += 1;
    }
  }
}

/**
 * The calculation chain, parsed far enough to keep it honest across a row
 * edit. A chain entry's sheet index is inherited from the previous entry when
 * the attribute is absent, so dropping an entry has to resolve it first.
 */
class CalcChainModel {
  readonly #header: string;
  readonly #footer: string;
  #entries: Array<{ openTag: string; reference: string; sheetId: number }>;

  constructor(xml: string) {
    const entries: Array<{
      openTag: string;
      reference: string;
      sheetId: number;
    }> = [];
    let inheritedSheetId = 0;
    let firstStart = -1;
    let lastEnd = -1;

    editElements(xml, "c", (element, text) => {
      const reference = getAttribute(element.openTag, "r");
      const declared = getAttribute(element.openTag, "i");
      if (declared !== undefined) {
        inheritedSheetId = Number(declared);
      }
      if (reference !== undefined) {
        if (firstStart < 0) {
          firstStart = element.start;
        }
        lastEnd = element.end;
        entries.push({
          openTag:
            declared === undefined
              ? addAttribute(element.openTag, "i", String(inheritedSheetId))
              : element.openTag,
          reference,
          sheetId: inheritedSheetId,
        });
      }
      return text;
    });

    this.#entries = entries;
    this.#header = firstStart < 0 ? xml : xml.slice(0, firstStart);
    this.#footer = lastEnd < 0 ? "" : xml.slice(lastEnd);
  }

  get isEmpty(): boolean {
    return this.#entries.length === 0;
  }

  /** Drop entries for deleted cells and renumber the rest; count the changes. */
  relocate(sheetId: number, relocation: RowRelocation): number {
    let adjusted = 0;
    this.#entries = this.#entries.filter((chainEntry) => {
      if (chainEntry.sheetId !== sheetId) {
        return true;
      }
      const relocated = relocateReference(chainEntry.reference, relocation);
      if (relocated.includes(DELETED_REFERENCE)) {
        adjusted += 1;
        return false;
      }
      if (relocated !== chainEntry.reference) {
        chainEntry.openTag = setAttribute(chainEntry.openTag, "r", relocated);
        chainEntry.reference = relocated;
        adjusted += 1;
      }
      return true;
    });
    return adjusted;
  }

  toXml(): string {
    return `${this.#header}${this.#entries
      .map((chainEntry) => chainEntry.openTag)
      .join("")}${this.#footer}`;
  }
}
