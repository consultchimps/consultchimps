/**
 * L3: the streaming workbook reader (ADR 0006). Consolidation reads through it,
 * and so do the table readers, through `../sheet-grid.ts`.
 *
 * The package layer streams a worksheet's markup; this module decides what each
 * cell holds. Every rule below is the rule the SheetJS-backed table reader
 * applied, cell for cell, so a read gives the values it gave before, with
 * these deliberate differences:
 *
 * - an error cell holds its text (`#DIV/0!`) as a `CellError`, not the number
 *   the engine coded it as;
 * - text keeps a carriage return before a line feed (`_x000D_` then a new
 *   line), which the engine folded into the line feed alone;
 * - a formula's cached text is unescaped once, as every other text is, where
 *   the engine unescaped it twice and so read a literal `&amp;lt;` as `<`;
 * - a declared date with no text reads as blank, where the engine dropped the
 *   whole worksheet;
 * - a row number or cell reference that is present but unreadable refuses the
 *   worksheet, the document model's rule, where the engine skipped the cell.
 *
 * Dates follow the document model, as they did before: a cell is a date when it
 * declares `t="d"` or when its style formats a number as one, and its value is
 * the model's ISO timestamp, through the one calendar route.
 */
import {
  ConsultChimpsError,
  type RandomAccessSource,
} from "@consultchimps/core";

import { XLSX_ERRORS } from "../../errors.js";
import {
  type ExcelTableDefinition,
  readExcelTableDefinitionsFrom,
  readWorkbookSheetsFrom,
} from "../../excel-tables.js";
import {
  calendarIsoText,
  serialMoment,
  utcCalendarParts,
  worksheetDateValue,
} from "../../model/calendar.js";
import { StyleTable } from "../../model/styles.js";
import { findElement, getAttribute } from "../../model/xml.js";
import {
  decodeEscapes,
  forEachDefinedName,
  forEachOpenTag,
  PreloadedParts,
  readSharedStrings,
  readWorksheetEvents,
  tagAttribute,
  ZipReader,
  type RawCell,
} from "../../package/index.js";
import { CellError } from "../../package/cell-error.js";
import { readFailure, type WorkbookReadContext } from "../read-model.js";

/** A value a streamed cell holds: a table value, or an error cell's text. */
export type StreamedValue = string | number | boolean | CellError;

/** A zero-based rectangle, in the engine's numbering. */
export interface CellRectangle {
  readonly startRow: number;
  readonly startColumn: number;
  readonly endRow: number;
  readonly endColumn: number;
}

/** One cell that holds something, empty text included. */
export interface StreamedCell {
  readonly column: number;
  readonly value: StreamedValue;
}

/** Receives a worksheet's rows top to bottom, cells left to right. */
export interface WorksheetConsumer {
  /** Called before the first row, and again if the read starts over. */
  begin(): void;
  /** A zero-based row holding at least one cell. */
  row(row: number, cells: readonly StreamedCell[]): void;
}

export interface WorksheetRead {
  /** The used range, or undefined when the worksheet declares none and holds nothing. */
  readonly range: CellRectangle | undefined;
  /** Merged ranges as the worksheet declares them. */
  readonly merges: readonly CellRectangle[];
  /**
   * Whether the rows had to be gathered before they could be delivered: a part
   * whose rows or cells are out of order, repeated, or filed under the wrong
   * row is read whole and delivered in order, as the engine indexed it.
   */
  readonly gathered: boolean;
}

/** A worksheet the workbook lists. */
export interface StreamedSheet {
  readonly name: string;
  readonly visible: boolean;
  /**
   * The worksheet part; undefined for a sheet that holds no cells to read,
   * such as a chart sheet; empty when nothing says where the sheet is.
   */
  readonly part: string | undefined;
}

/** A defined name as the workbook part declares it. */
export interface StreamedDefinedName {
  readonly name: string;
  /** The formula the name stands for, such as `Data!$A$1:$C$9`. */
  readonly reference: string;
}

/** How one workbook is named in errors and in the output. */
export interface StreamedWorkbookContext extends WorkbookReadContext {
  /** The name a table records as its source file. */
  readonly file: string;
}

const WORKBOOK_PART = "xl/workbook.xml";
const STYLES_PART = "xl/styles.xml";
const CONTENT_TYPES_PART = "[Content_Types].xml";
const CONTENT_TYPES_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/content-types";
const SHARED_STRINGS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const WORKSHEET_RELATIONSHIP_SUFFIX = "/worksheet";
const TABLE_RELATIONSHIP_SUFFIX = "/table";

/** The error values a cell can hold, as the file format spells them. */
const ERROR_VALUES = new Set([
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#GETTING_DATA",
  "#WTF?",
]);

/** Cell types whose text the model never reads as a date. */
const NON_DATE_TYPES = new Set(["b", "d", "s", "str", "inlineStr", "e"]);

/**
 * A cell reference read the way the engine reads one: the leading capital
 * letters as a column, whatever follows.
 */
function referenceColumn(ref: string): number {
  let index = 0;
  for (let position = 0; position < ref.length; position += 1) {
    const code = ref.charCodeAt(position) - 64;
    if (code < 1 || code > 26) break;
    index = 26 * index + code;
  }
  return index - 1;
}

const CANONICAL_REFERENCE = /^([A-Z]+)([1-9]\d*)$/u;

/** The zero-based row of a reference the engine can look up, or undefined. */
function referenceRow(ref: string): number | undefined {
  const match = CANONICAL_REFERENCE.exec(ref);
  return match ? Number(match[2]) - 1 : undefined;
}

/**
 * A range reference read leniently, the way the engine reads `<dimension>`
 * and every other range a workbook declares.
 */
export function decodeRange(range: string): CellRectangle {
  let index: number;
  let code = 0;
  let position = 0;
  const length = range.length;
  for (index = 0; position < length; position += 1) {
    code = range.charCodeAt(position) - 64;
    if (code < 1 || code > 26) break;
    index = 26 * index + code;
  }
  const startColumn = index - 1;
  for (index = 0; position < length; position += 1) {
    code = range.charCodeAt(position) - 48;
    if (code < 0 || code > 9) break;
    index = 10 * index + code;
  }
  const startRow = index - 1;
  if (position === length || code !== 10) {
    return { startRow, startColumn, endRow: startRow, endColumn: startColumn };
  }
  position += 1;
  for (index = 0; position !== length; position += 1) {
    code = range.charCodeAt(position) - 64;
    if (code < 1 || code > 26) break;
    index = 26 * index + code;
  }
  const endColumn = index - 1;
  for (index = 0; position !== length; position += 1) {
    code = range.charCodeAt(position) - 48;
    if (code < 0 || code > 9) break;
    index = 10 * index + code;
  }
  return { startRow, startColumn, endRow: index - 1, endColumn };
}

function declaredDimension(ref: string): CellRectangle | undefined {
  if (!/^\w*:\w*$/u.test(ref)) return undefined;
  const range = decodeRange(ref);
  return range.startRow <= range.endRow &&
    range.startColumn <= range.endColumn &&
    range.startRow >= 0 &&
    range.startColumn >= 0
    ? range
    : undefined;
}

function parseBoolean(text: string | undefined): boolean {
  return text === "1" || text === "true";
}

function isoText(moment: Date): string {
  return calendarIsoText(utcCalendarParts(moment));
}

/** What the engine stores for a cell, before the model's dates are applied. */
type EngineCell =
  /** The cell is not stored and does not count towards the used range. */
  | { readonly kind: "absent" }
  /** The cell counts towards the used range but holds nothing. */
  | { readonly kind: "counted" }
  | { readonly kind: "stored"; readonly value: StreamedValue | null };

const ABSENT: EngineCell = { kind: "absent" };
const COUNTED: EngineCell = { kind: "counted" };

/** The listed-but-unreadable refusal the table reader raises. */
function unreadableWorksheet(
  sheet: string,
  file: string,
  cause?: unknown,
): ConsultChimpsError {
  return new ConsultChimpsError(
    XLSX_ERRORS.XLSX_READ_FAILED,
    `Worksheet "${sheet}" is listed in ${file} but could not be read from it, so what it holds is unknown.`,
    { cause, details: { source: file, worksheet: sheet } },
  );
}

/** A cell reference the document model can read; see `decodeCell`. */
const MODEL_REFERENCE = /^[A-Za-z]{1,3}\d+$/u;

/**
 * Refuse what the document model refuses: a row number or a cell reference
 * that is present but unreadable marks a damaged worksheet, where an absent
 * one means document order. Every reader the model backed refused these.
 */
function assertWellFormedRow(
  ref: string | undefined,
  cells: readonly RawCell[],
): void {
  if (ref !== undefined) {
    const number = Number(ref);
    if (!Number.isInteger(number) || number < 1) {
      throw new Error(
        `Encountered a worksheet row with an invalid row number: ${ref}`,
      );
    }
  }
  for (const cell of cells) {
    if (cell.ref !== undefined && !MODEL_REFERENCE.test(cell.ref)) {
      throw new Error(`Encountered an invalid cell reference: ${cell.ref}`);
    }
  }
}

class UnresolvedString extends Error {}
class OutOfOrder extends Error {}

/**
 * One workbook opened for streaming: its sheets, styles, date system and
 * shared-string location are read up front; worksheets are read on request,
 * as many times as a caller asks.
 */
export class StreamedWorkbook {
  readonly sheets: readonly StreamedSheet[];
  /** The Excel Tables the worksheets carry, as the document model reads them. */
  readonly tables: readonly ExcelTableDefinition[];
  /** The defined names, in workbook order, built-in names included. */
  readonly names: readonly StreamedDefinedName[];
  readonly #zip: ZipReader;

  /** Changes whenever the package's contents change; see `ZipReader`. */
  get fingerprint(): number {
    return this.#zip.fingerprint;
  }
  readonly #context: StreamedWorkbookContext;
  readonly #styles: StyleTable;
  readonly #date1904: boolean;
  readonly #stringsPart: string | undefined;
  #strings: ReadonlyArray<string | undefined> | undefined;

  private constructor(
    zip: ZipReader,
    context: StreamedWorkbookContext,
    sheets: readonly StreamedSheet[],
    tables: readonly ExcelTableDefinition[],
    names: readonly StreamedDefinedName[],
    styles: StyleTable,
    date1904: boolean,
    stringsPart: string | undefined,
  ) {
    this.#zip = zip;
    this.#context = context;
    this.sheets = sheets;
    this.tables = tables;
    this.names = names;
    this.#styles = styles;
    this.#date1904 = date1904;
    this.#stringsPart = stringsPart;
  }

  /**
   * Open a workbook, refusing one that cannot be read with the stable read
   * error: not a zip, encrypted, or missing or malformed structural parts,
   * including an Excel Table definition the model cannot parse.
   */
  static async open(
    source: RandomAccessSource,
    context: StreamedWorkbookContext,
  ): Promise<StreamedWorkbook> {
    try {
      return await StreamedWorkbook.#open(source, context);
    } catch (error) {
      throw readFailure(context, undefined, error);
    }
  }

  static async #open(
    source: RandomAccessSource,
    context: StreamedWorkbookContext,
  ): Promise<StreamedWorkbook> {
    const zip = await ZipReader.open(source);
    const contentTypes = await zip.readText(CONTENT_TYPES_PART);
    if (contentTypes === undefined) {
      throw new Error("The package has no [Content_Types].xml part.");
    }
    let namespace: string | undefined;
    let stringsPart: string | undefined;
    forEachOpenTag(contentTypes, CONTENT_TYPES_PART, (tag) => {
      if (tag.local === "Types") {
        namespace = tag.uri;
      } else if (
        tag.local === "Override" &&
        stringsPart === undefined &&
        tagAttribute(tag, "ContentType") === SHARED_STRINGS_CONTENT_TYPE
      ) {
        stringsPart = tagAttribute(tag, "PartName")?.replace(/^\//u, "");
      }
    });
    if (namespace !== CONTENT_TYPES_NAMESPACE) {
      throw new Error(`Unknown content types namespace: ${namespace ?? ""}`);
    }

    const parts = new PreloadedParts(zip, context.source);
    const workbookXml = (await parts.load(WORKBOOK_PART)) ?? "";
    await parts.loadRelationships(WORKBOOK_PART);
    // The model's reading of the structure, which a read has always had to
    // pass: the worksheet parts the relationships name, and every Excel Table
    // definition those worksheets carry.
    const entries = readWorkbookSheetsFrom(parts);
    for (const entry of entries) {
      await parts.loadRelationships(entry.worksheetPart);
      for (const relationship of parts.relationshipsOf(entry.worksheetPart)) {
        if (relationship.type.endsWith(TABLE_RELATIONSHIP_SUFFIX)) {
          await parts.load(
            parts.resolvePart(entry.worksheetPart, relationship.target),
          );
        }
      }
    }
    const tables = readExcelTableDefinitionsFrom(parts);

    const relationships = new Map(
      parts
        .relationshipsOf(WORKBOOK_PART)
        .map((relationship) => [relationship.id, relationship] as const),
    );
    const sheets: StreamedSheet[] = [];
    forEachOpenTag(workbookXml, WORKBOOK_PART, (tag) => {
      if (tag.local !== "sheet") return;
      const name = tagAttribute(tag, "name");
      if (name === undefined) return;
      const id = tagAttribute(tag, "id");
      const relationship = id === undefined ? undefined : relationships.get(id);
      const state = tagAttribute(tag, "state");
      let part: string | undefined;
      if (relationship === undefined) {
        // Nothing says where the sheet's cells are, so they cannot be read;
        // `readWorksheet` reports it rather than treating it as empty.
        part = "";
      } else if (relationship.type.endsWith(WORKSHEET_RELATIONSHIP_SUFFIX)) {
        part = parts.resolvePart(WORKBOOK_PART, relationship.target);
      }
      sheets.push({
        name: decodeEscapes(name),
        visible: state !== "hidden" && state !== "veryHidden",
        part,
      });
    });

    const names: StreamedDefinedName[] = [];
    forEachDefinedName(workbookXml, WORKBOOK_PART, (name, reference) => {
      names.push({ name, reference: decodeEscapes(reference) });
    });

    const properties = findElement(workbookXml, "workbookPr");
    const declared = properties
      ? getAttribute(properties.openTag, "date1904")
      : undefined;
    return new StreamedWorkbook(
      zip,
      context,
      sheets,
      tables,
      names,
      StyleTable.parse(await zip.readText(STYLES_PART)),
      declared === "1" || declared === "true",
      stringsPart,
    );
  }

  /**
   * Load the shared strings, which a worksheet read needs; `releaseStrings`
   * lets them go. A table the engine could not read leaves every string
   * unresolved, and a worksheet that uses one is then refused.
   */
  async loadStrings(): Promise<void> {
    if (this.#strings !== undefined) return;
    this.#strings = [];
    if (this.#stringsPart === undefined || !this.#zip.has(this.#stringsPart)) {
      return;
    }
    try {
      this.#strings = await readSharedStrings(this.#zip, this.#stringsPart);
    } catch {
      this.#strings = [];
    }
  }

  releaseStrings(): void {
    this.#strings = undefined;
  }

  /**
   * Read one worksheet into `consumer`. Rows are delivered as the part is
   * read; a part whose rows arrive out of order is read again whole and
   * delivered in order, after `begin` is called a second time. Pass `gather`
   * to read whole from the start, as a second read of a worksheet the first
   * read had to gather should. Cells outside the used range are left out
   * unless `clip` is false, for a reader whose rectangle is declared rather
   * than taken from the used range, such as an Excel Table or a named range.
   */
  async readWorksheet(
    sheet: StreamedSheet,
    consumer: WorksheetConsumer,
    options: {
      gather?: boolean;
      clip?: boolean;
      between?: () => Promise<void>;
    } = {},
  ): Promise<WorksheetRead> {
    const file = this.#context.file;
    if (sheet.part === undefined) {
      consumer.begin();
      return { range: undefined, merges: [], gathered: false };
    }
    if (sheet.part === "" || !this.#zip.has(sheet.part)) {
      throw unreadableWorksheet(sheet.name, file);
    }
    await this.loadStrings();
    const clip = options.clip ?? true;
    try {
      if (options.gather !== true) {
        try {
          return await this.#read(
            sheet.part,
            consumer,
            false,
            clip,
            options.between,
          );
        } catch (error) {
          if (!(error instanceof OutOfOrder)) throw error;
        }
      }
      return await this.#read(
        sheet.part,
        consumer,
        true,
        clip,
        options.between,
      );
    } catch (error) {
      if (error instanceof UnresolvedString) {
        throw unreadableWorksheet(sheet.name, file, error);
      }
      if (error instanceof ConsultChimpsError) throw error;
      throw readFailure(this.#context, sheet.name, error);
    }
  }

  /** What the engine stores for a cell, before dates. */
  #engineCell(raw: RawCell, formula: boolean): EngineCell {
    const type = raw.type;
    if (type === undefined && raw.value === undefined) {
      // A formula with no cached value reads as nothing, but the engine counts
      // it towards the used range all the same.
      return formula ? COUNTED : ABSENT;
    }
    const text = raw.value === undefined ? undefined : decodeEscapes(raw.value);
    switch (type ?? "n") {
      case "n": {
        if (text === undefined) return COUNTED;
        const number = Number.parseFloat(text);
        return {
          kind: "stored",
          value: Number.isFinite(number) ? number : null,
        };
      }
      case "s": {
        if (text === undefined) return COUNTED;
        const strings = this.#strings ?? [];
        const index = Number.parseInt(text, 10);
        if (!(index >= 0 && index < strings.length)) {
          throw new UnresolvedString(
            `Shared string ${text} is not in the workbook's string table.`,
          );
        }
        return { kind: "stored", value: strings[index] ?? null };
      }
      case "str":
        return { kind: "stored", value: text ?? "" };
      case "inlineStr":
        return {
          kind: "stored",
          value: raw.inline === null ? "" : (raw.inline ?? null),
        };
      case "b":
        return { kind: "stored", value: parseBoolean(text) };
      case "e":
        return {
          kind: "stored",
          value:
            text !== undefined && ERROR_VALUES.has(text)
              ? new CellError(text)
              : null,
        };
      case "d":
        // The model answers for every declared date it can read; one it reads
        // as blank is blank to the engine too.
        return { kind: "stored", value: null };
      default:
        return { kind: "stored", value: text ?? null };
    }
  }

  /** The model's date for a cell, or undefined when it is not a date. */
  #modelDate(raw: RawCell): string | undefined {
    const text = raw.value;
    if (text === undefined) return undefined;
    if (raw.type === "d") {
      if (text.trim() === "") return undefined;
      const moment = worksheetDateValue(text);
      return moment === undefined ? text : isoText(moment);
    }
    if (raw.type !== undefined && NON_DATE_TYPES.has(raw.type)) {
      return undefined;
    }
    const style = raw.style === undefined ? undefined : Number(raw.style);
    if (!this.#styles.isDateStyle(style)) return undefined;
    const trimmed = text.trim();
    if (trimmed === "") return undefined;
    const numeric = Number(trimmed);
    if (!Number.isFinite(numeric)) return undefined;
    const moment = serialMoment(numeric, this.#date1904);
    return moment === undefined ? undefined : isoText(moment);
  }

  async #read(
    part: string,
    consumer: WorksheetConsumer,
    gather: boolean,
    clip: boolean,
    between: (() => Promise<void>) | undefined,
  ): Promise<WorksheetRead> {
    consumer.begin();
    let dimension: CellRectangle | undefined;
    let guessStartRow = 2_000_000;
    let guessEndRow = 0;
    let guessStartColumn = 2_000_000;
    let guessEndColumn = 0;
    let rowTag = 0;
    let lastRow = -1;
    let pendingSelfClosing: string | undefined | null = null;
    const merges: CellRectangle[] = [];
    // Shared formulas defined so far, and array formula ranges, which make a
    // cell with no value of its own a formula cell, as the engine reads them.
    const sharedFormulas = new Set<number>();
    const arrayRanges: CellRectangle[] = [];
    const hasFormula = (
      raw: RawCell,
      row: number | undefined,
      column: number,
    ): boolean => {
      const formula = raw.formula;
      let found = false;
      if (formula !== undefined) {
        const shared =
          formula.sharedIndex === undefined
            ? Number.NaN
            : Number.parseInt(formula.sharedIndex, 10);
        if (formula.text !== "") {
          found = true;
          if (formula.type === "array" && formula.ref?.includes(":")) {
            arrayRanges.push(decodeRange(formula.ref));
          } else if (formula.type === "shared" && !Number.isNaN(shared)) {
            sharedFormulas.add(shared);
          }
        } else if (sharedFormulas.has(shared)) {
          found = true;
        }
      }
      if (!found && row !== undefined) {
        found = arrayRanges.some(
          (range) =>
            row >= range.startRow &&
            row <= range.endRow &&
            column >= range.startColumn &&
            column <= range.endColumn,
        );
      }
      return found;
    };
    // The whole worksheet, keyed the way the engine keys it, when gathering.
    const gathered = gather
      ? new Map<number, Map<number, StreamedValue | null>>()
      : undefined;

    const inDimension = (row: number, column: number): boolean =>
      !clip ||
      dimension === undefined ||
      (row >= dimension.startRow &&
        row <= dimension.endRow &&
        column >= dimension.startColumn &&
        column <= dimension.endColumn);

    const processRow = (
      ref: string | undefined,
      cells: readonly RawCell[],
    ): void => {
      rowTag = ref === undefined ? rowTag + 1 : Number.parseInt(ref, 10);
      const tagValid = Number.isInteger(rowTag) && rowTag >= 1;
      if (tagValid) {
        guessStartRow = Math.min(guessStartRow, rowTag - 1);
        guessEndRow = Math.max(guessEndRow, rowTag - 1);
      }
      let columnTag = -1;
      let rowCells: StreamedCell[] | undefined;
      let previousColumn = -1;
      for (const raw of cells) {
        let row: number | undefined;
        if (raw.ref !== undefined) {
          columnTag = referenceColumn(raw.ref);
          row = referenceRow(raw.ref);
        } else {
          columnTag += 1;
          row = tagValid ? rowTag - 1 : undefined;
        }
        const engine = this.#engineCell(raw, hasFormula(raw, row, columnTag));
        if (engine.kind === "absent") continue;
        if (columnTag >= 0) {
          guessStartColumn = Math.min(guessStartColumn, columnTag);
          guessEndColumn = Math.max(guessEndColumn, columnTag);
        }
        if (engine.kind === "counted" || row === undefined || columnTag < 0) {
          continue;
        }
        const value = this.#modelDate(raw) ?? engine.value;
        if (gathered) {
          let cellsOfRow = gathered.get(row);
          if (cellsOfRow === undefined) {
            cellsOfRow = new Map();
            gathered.set(row, cellsOfRow);
          }
          cellsOfRow.set(columnTag, value);
          continue;
        }
        if (
          row !== rowTag - 1 ||
          row <= lastRow ||
          columnTag <= previousColumn
        ) {
          throw new OutOfOrder();
        }
        previousColumn = columnTag;
        if (value === null || !inDimension(row, columnTag)) continue;
        (rowCells ??= []).push({ column: columnTag, value });
      }
      if (!gathered && previousColumn >= 0) {
        lastRow = rowTag - 1;
      }
      if (rowCells !== undefined) {
        consumer.row(rowTag - 1, rowCells);
      }
    };

    await readWorksheetEvents(
      this.#zip,
      part,
      {
        dimension(ref) {
          dimension = declaredDimension(ref);
        },
        row(ref, selfClosing, cells) {
          assertWellFormedRow(ref, cells);
          // The engine reads a self-closing row only when nothing but other
          // self-closing rows follows it, and then only the last of them.
          if (selfClosing) {
            pendingSelfClosing = ref;
            return;
          }
          pendingSelfClosing = null;
          processRow(ref, cells);
        },
        endRows() {
          if (pendingSelfClosing !== null) {
            processRow(pendingSelfClosing, []);
            pendingSelfClosing = null;
          }
        },
        mergeCell(attributes) {
          const keys = Object.keys(attributes);
          const ref = attributes["ref"];
          if (
            keys.length === 1 &&
            ref !== undefined &&
            /^[A-Z0-9:]+$/u.test(ref)
          ) {
            merges.push(decodeRange(ref));
          }
        },
      },
      between,
    );

    const range =
      dimension ??
      (guessEndColumn >= guessStartColumn && guessEndRow >= guessStartRow
        ? {
            startRow: guessStartRow,
            startColumn: guessStartColumn,
            endRow: guessEndRow,
            endColumn: guessEndColumn,
          }
        : undefined);

    if (gathered && range !== undefined) {
      const rows = [...gathered.keys()]
        .filter(
          (row) => !clip || (row >= range.startRow && row <= range.endRow),
        )
        .sort((left, right) => left - right);
      for (const row of rows) {
        const cellsOfRow = gathered.get(row)!;
        const delivered: StreamedCell[] = [];
        for (const column of [...cellsOfRow.keys()].sort((a, b) => a - b)) {
          const value = cellsOfRow.get(column);
          if (
            value === null ||
            value === undefined ||
            (clip && (column < range.startColumn || column > range.endColumn))
          ) {
            continue;
          }
          delivered.push({ column, value });
        }
        if (delivered.length > 0) consumer.row(row, delivered);
      }
    }
    return { range, merges, gathered: gather };
  }
}
