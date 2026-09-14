import type { FileEntry } from "@zip.js/zip.js";
import { throwIfAborted } from "@consultchimps/core";

import type { SharedStrings } from "./strings.js";
import type { WorkbookStyles } from "./styles.js";
import type {
  StreamCell,
  StreamRow,
  StreamScalarCell,
  WorkbookStreamOptions,
} from "./types.js";
import {
  calendarIsoText,
  utcCalendarParts,
  worksheetDateValue,
} from "../model/calendar.js";
import { decodeCell, encodeCell } from "../model/references.js";
import { BoundedXmlText } from "./xml.js";
import { createElementParser, unqualifiedAttribute } from "./xml-elements.js";
import { entryChunks, type StreamLimits } from "./zip.js";

type PendingScalarCell =
  StreamScalarCell | { readonly kind: "shared"; readonly index: number };

type CellType = "b" | "d" | "e" | "inlineStr" | "n" | "s" | "str";

type PendingCell =
  | PendingScalarCell
  | {
      readonly kind: "formula";
      readonly formula?: string | undefined;
      readonly cached: PendingScalarCell | { readonly kind: "missing" };
    };

interface PendingRow {
  readonly sourceRow: number;
  readonly cells: Readonly<Record<string, PendingCell>>;
}

interface CurrentCell {
  readonly reference: string;
  readonly column: number;
  readonly type: CellType | undefined;
  readonly style: number;
  value: string;
  formula: string;
  inline: string;
  valueOpen: boolean;
  formulaOpen: boolean;
  inlineTextOpen: boolean;
  inlineContainerOpen: boolean;
  hasValue: boolean;
  hasFormula: boolean;
  hasInlineContainer: boolean;
  hasInlineText: boolean;
  valueBytes: number;
  formulaBytes: number;
  inlineBytes: number;
}

function cellRecord<Cell>(): Record<string, Cell> {
  return Object.create(null) as Record<string, Cell>;
}

function cellType(value: string | undefined): CellType | undefined {
  switch (value) {
    case undefined:
    case "b":
    case "d":
    case "e":
    case "inlineStr":
    case "n":
    case "s":
    case "str":
      return value;
    default:
      throw new Error(`A worksheet cell has unsupported cell type "${value}".`);
  }
}

function scalarCell(
  current: CurrentCell,
  styles: WorkbookStyles,
): PendingScalarCell {
  const raw = current.type === "inlineStr" ? current.inline : current.value;
  if (current.type === "s") {
    if (!current.hasValue) return { kind: "blank" };
    if (!/^\d+$/u.test(raw)) {
      throw new Error(
        `Cell ${current.reference} has an invalid shared-string index.`,
      );
    }
    return { kind: "shared", index: Number(raw) };
  }
  if (current.type === "b") {
    if (!current.hasValue || raw === "") return { kind: "blank" };
    if (raw !== "0" && raw !== "1") {
      throw new Error(
        `Cell ${current.reference} has an invalid Boolean value.`,
      );
    }
    return { kind: "boolean", value: raw === "1" };
  }
  if (current.type === "e") {
    return !current.hasValue || raw === ""
      ? { kind: "blank" }
      : { kind: "error", error: raw };
  }
  if (current.type === "d") {
    if (!current.hasValue || raw.trim() === "") return { kind: "blank" };
    const date = worksheetDateValue(raw);
    return date === undefined
      ? { kind: "string", value: raw }
      : {
          kind: "date",
          raw,
          iso: calendarIsoText(utcCalendarParts(date)),
        };
  }
  if (current.type === "str") {
    return current.hasValue
      ? { kind: "string", value: raw }
      : { kind: "blank" };
  }
  if (current.type === "inlineStr") {
    if (!current.hasInlineText) return { kind: "blank" };
    return { kind: "string", value: raw };
  }
  if (!current.hasValue || raw === "") return { kind: "blank" };
  if (!/^-?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/u.test(raw)) {
    throw new Error(`Cell ${current.reference} has an invalid numeric value.`);
  }
  const iso = styles.dateValue(raw, current.style);
  return iso === undefined
    ? { kind: "number", raw }
    : { kind: "date", raw, iso };
}

async function hydrateScalar(
  cell: PendingScalarCell,
  sharedStrings: SharedStrings,
): Promise<StreamScalarCell> {
  return cell.kind === "shared"
    ? { kind: "string", value: await sharedStrings.value(cell.index) }
    : cell;
}

async function hydrateCell(
  cell: PendingCell,
  sharedStrings: SharedStrings,
): Promise<StreamCell> {
  if (cell.kind !== "formula") return hydrateScalar(cell, sharedStrings);
  return {
    kind: "formula",
    ...(cell.formula === undefined ? {} : { formula: cell.formula }),
    cached:
      cell.cached.kind === "missing"
        ? cell.cached
        : await hydrateScalar(cell.cached, sharedStrings),
  };
}

async function hydrateRows(
  rows: readonly PendingRow[],
  sharedStrings: SharedStrings,
): Promise<readonly StreamRow[]> {
  const result: StreamRow[] = [];
  for (const row of rows) {
    const cells = cellRecord<StreamCell>();
    for (const [name, cell] of Object.entries(row.cells)) {
      cells[name] = await hydrateCell(cell, sharedStrings);
    }
    const hasSelectedValue = Object.values(cells).some(
      (cell) =>
        cell.kind === "formula" ||
        (cell.kind === "string"
          ? cell.value.length > 0
          : cell.kind !== "blank"),
    );
    if (hasSelectedValue) result.push({ sourceRow: row.sourceRow, cells });
  }
  return result;
}

export async function* parseWorksheetBatches(
  entry: FileEntry,
  options: {
    readonly firstRow: number;
    readonly lastRow: number;
    readonly columns?: ReadonlyMap<number, string> | undefined;
    readonly batchSize: number;
    readonly sharedStrings: SharedStrings;
    readonly styles: WorkbookStyles;
    readonly limits: StreamLimits;
    readonly signal?: AbortSignal | undefined;
    readonly onProgress?: WorkbookStreamOptions["onProgress"];
  },
): AsyncIterable<readonly StreamRow[]> {
  const ready: PendingRow[] = [];
  let activeRow = 0;
  let rowOpen = false;
  let previousRow = 0;
  let nextImplicitRow = 1;
  let nextImplicitColumn = 0;
  let cells = cellRecord<PendingCell>();
  let seenColumns = new Set<number>();
  let current: CurrentCell | undefined;
  let passedLastRow = false;
  const decoder = new TextDecoder("utf-8", { fatal: true });
  const encoder = new TextEncoder();
  const textLimit = new BoundedXmlText(
    new Set(["t", "v", "f"]),
    options.limits.maximumCellBytes,
    (element) =>
      `A worksheet ${element} text node exceeds the configured ${options.limits.maximumCellBytes}-byte cell limit.`,
  );

  const parser = createElementParser({
    root: "worksheet",
    open(tag, path) {
      if (path.inDocument && tag.local === "row") {
        if (!path.is("worksheet", "sheetData", "row")) {
          throw new Error("A worksheet row cannot contain another row.");
        }
        if (rowOpen) {
          throw new Error("A worksheet row cannot contain another row.");
        }
        const rawRow = unqualifiedAttribute(tag, "r");
        if (
          rawRow !== undefined &&
          (rawRow.length === 0 || /\D/u.test(rawRow))
        ) {
          throw new Error("A worksheet row has an invalid row number.");
        }
        activeRow = rawRow === undefined ? nextImplicitRow : Number(rawRow);
        if (
          !Number.isSafeInteger(activeRow) ||
          activeRow < 1 ||
          activeRow > 1_048_576 ||
          activeRow <= previousRow
        ) {
          throw new Error("A worksheet row has an invalid row number.");
        }
        rowOpen = true;
        previousRow = activeRow;
        nextImplicitRow = activeRow + 1;
        nextImplicitColumn = 0;
        cells = cellRecord<PendingCell>();
        seenColumns = new Set<number>();
        if (activeRow > options.lastRow) passedLastRow = true;
        return;
      }
      if (path.inDocument && tag.local === "c") {
        if (!path.is("worksheet", "sheetData", "row", "c")) {
          throw new Error(
            current === undefined
              ? "A worksheet cell must be inside a row."
              : "A worksheet cell cannot contain another cell.",
          );
        }
        if (!rowOpen) {
          throw new Error("A worksheet cell must be inside a row.");
        }
        if (current !== undefined) {
          throw new Error("A worksheet cell cannot contain another cell.");
        }
        const type = cellType(unqualifiedAttribute(tag, "t"));
        const explicitReference = unqualifiedAttribute(tag, "r");
        const parsedReference =
          explicitReference === undefined
            ? { column: nextImplicitColumn, row: activeRow }
            : decodeCell(explicitReference);
        if (
          parsedReference === undefined ||
          parsedReference.column < 0 ||
          parsedReference.column >= 16_384 ||
          !Number.isSafeInteger(parsedReference.row) ||
          parsedReference.row < 1 ||
          parsedReference.row > 1_048_576 ||
          parsedReference.row !== activeRow ||
          (explicitReference !== undefined &&
            encodeCell(parsedReference.column, parsedReference.row) !==
              explicitReference.toUpperCase())
        ) {
          throw new Error("A worksheet cell has an invalid reference.");
        }
        const parsedColumn = parsedReference.column;
        if (seenColumns.has(parsedColumn)) {
          throw new Error(
            `Worksheet row ${activeRow} contains column ${parsedColumn + 1} more than once.`,
          );
        }
        seenColumns.add(parsedColumn);
        nextImplicitColumn = parsedColumn + 1;
        const explicitStyle = unqualifiedAttribute(tag, "s");
        const style = Number(explicitStyle ?? "0");
        if (!Number.isSafeInteger(style) || style < 0) {
          throw new Error(
            `Cell ${explicitReference ?? "(implicit)"} has an invalid style.`,
          );
        }
        if (explicitStyle !== undefined && !options.styles.hasStyle(style)) {
          throw new Error(
            `Cell ${explicitReference ?? "(implicit)"} has a style outside the workbook cell formats.`,
          );
        }
        current = {
          reference: explicitReference ?? encodeCell(parsedColumn, activeRow),
          column: parsedColumn,
          type,
          style,
          value: "",
          formula: "",
          inline: "",
          valueOpen: false,
          formulaOpen: false,
          inlineTextOpen: false,
          inlineContainerOpen: false,
          hasValue: false,
          hasFormula: false,
          hasInlineContainer: false,
          hasInlineText: false,
          valueBytes: 0,
          formulaBytes: 0,
          inlineBytes: 0,
        };
        return;
      }
      if (!current || !path.inDocument) return;
      if (current.valueOpen || current.formulaOpen) {
        throw new Error(
          `Cell ${current.reference} has nested markup inside its value or formula.`,
        );
      }
      if (tag.local === "v") {
        if (!path.is("worksheet", "sheetData", "row", "c", "v")) {
          throw new Error(
            `Cell ${current.reference} value must be a direct child of the cell.`,
          );
        }
        if (current.type === "inlineStr") {
          throw new Error(
            `An inline string cannot use a value element in cell ${current.reference}.`,
          );
        }
        if (current.hasValue) {
          throw new Error(
            `Cell ${current.reference} has a value more than once.`,
          );
        }
        current.valueOpen = true;
        current.hasValue = true;
      } else if (tag.local === "f") {
        if (!path.is("worksheet", "sheetData", "row", "c", "f")) {
          throw new Error(
            `Cell ${current.reference} formula must be a direct child of the cell.`,
          );
        }
        if (current.hasFormula) {
          throw new Error(
            `Cell ${current.reference} has a formula more than once.`,
          );
        }
        current.formulaOpen = true;
        current.hasFormula = true;
      } else if (tag.local === "is") {
        if (
          !path.is("worksheet", "sheetData", "row", "c", "is") ||
          current.type !== "inlineStr"
        ) {
          throw new Error(
            `Cell ${current.reference} inline string must be a direct child of an inline-string cell.`,
          );
        }
        if (current.hasInlineContainer) {
          throw new Error(
            `Cell ${current.reference} has an inline string more than once.`,
          );
        }
        current.inlineContainerOpen = true;
        current.hasInlineContainer = true;
      } else if (tag.local === "t") {
        const inlineText =
          path.is("worksheet", "sheetData", "row", "c", "is", "t") ||
          path.is("worksheet", "sheetData", "row", "c", "is", "r", "t");
        if (inlineText) {
          if (current.type !== "inlineStr" || !current.inlineContainerOpen) {
            throw new Error(
              `Cell ${current.reference} has invalid inline-string text.`,
            );
          }
          current.inlineTextOpen = true;
          current.hasInlineText = true;
        } else if (
          !path.within("worksheet", "sheetData", "row", "c", "is", "rPh")
        ) {
          throw new Error(
            `Cell ${current.reference} has invalid inline-string text.`,
          );
        }
      }
    },
    text(text, path) {
      if (!current) return;
      const valueText = path.is("worksheet", "sheetData", "row", "c", "v");
      const formulaText = path.is("worksheet", "sheetData", "row", "c", "f");
      const inlineText =
        path.is("worksheet", "sheetData", "row", "c", "is", "t") ||
        path.is("worksheet", "sheetData", "row", "c", "is", "r", "t");
      if (!valueText && !formulaText && !inlineText) return;
      const bytes = encoder.encode(text).byteLength;
      if (valueText && current.valueOpen) {
        current.valueBytes += bytes;
        if (current.valueBytes > options.limits.maximumCellBytes)
          throw new Error(
            `Cell ${current.reference} value exceeds the configured ${options.limits.maximumCellBytes}-byte limit.`,
          );
        current.value += text;
      }
      if (formulaText && current.formulaOpen) {
        current.formulaBytes += bytes;
        if (current.formulaBytes > options.limits.maximumCellBytes)
          throw new Error(
            `Cell ${current.reference} formula exceeds the configured ${options.limits.maximumCellBytes}-byte limit.`,
          );
        current.formula += text;
      }
      if (inlineText && current.inlineTextOpen) {
        current.inlineBytes += bytes;
        if (current.inlineBytes > options.limits.maximumCellBytes)
          throw new Error(
            `Cell ${current.reference} inline string exceeds the configured ${options.limits.maximumCellBytes}-byte limit.`,
          );
        current.inline += text;
      }
    },
    close(_tag, path) {
      if (current && path.is("worksheet", "sheetData", "row", "c", "v")) {
        current.valueOpen = false;
      } else if (
        current &&
        path.is("worksheet", "sheetData", "row", "c", "f")
      ) {
        current.formulaOpen = false;
      } else if (
        current &&
        (path.is("worksheet", "sheetData", "row", "c", "is", "t") ||
          path.is("worksheet", "sheetData", "row", "c", "is", "r", "t"))
      ) {
        current.inlineTextOpen = false;
      } else if (
        current &&
        path.is("worksheet", "sheetData", "row", "c", "is")
      ) {
        current.inlineContainerOpen = false;
      } else if (current && path.is("worksheet", "sheetData", "row", "c")) {
        const selectedName = options.columns?.get(current.column);
        if (
          activeRow >= options.firstRow &&
          activeRow <= options.lastRow &&
          (options.columns === undefined || selectedName !== undefined)
        ) {
          const key = selectedName ?? current.reference;
          if (current.hasFormula) {
            cells[key] = {
              kind: "formula",
              ...(current.formula === "" ? {} : { formula: current.formula }),
              cached: current.hasValue
                ? scalarCell(current, options.styles)
                : { kind: "missing" },
            };
          } else {
            cells[key] = scalarCell(current, options.styles);
          }
        }
        current = undefined;
      } else if (path.is("worksheet", "sheetData", "row")) {
        if (!rowOpen || current !== undefined) {
          throw new Error("A worksheet row has invalid cell structure.");
        }
        if (activeRow >= options.firstRow && activeRow <= options.lastRow) {
          ready.push({ sourceRow: activeRow, cells });
        }
        rowOpen = false;
        activeRow = 0;
        cells = cellRecord<PendingCell>();
        seenColumns = new Set<number>();
      }
    },
  });

  for await (const chunk of entryChunks(entry, {
    signal: options.signal,
    onProgress: options.onProgress,
    stage: "worksheet",
  })) {
    if (!passedLastRow) {
      textLimit.consume(chunk);
      parser.write(decoder.decode(chunk, { stream: true }));
    }
    while (ready.length >= options.batchSize) {
      throwIfAborted(options.signal, "xlsx.stream", "memory");
      const rows = await hydrateRows(
        ready.splice(0, options.batchSize),
        options.sharedStrings,
      );
      if (rows.length > 0) yield rows;
    }
  }
  if (!passedLastRow) {
    parser.write(decoder.decode());
    parser.close();
  }
  if (ready.length > 0) {
    throwIfAborted(options.signal, "xlsx.stream", "memory");
    const rows = await hydrateRows(ready, options.sharedStrings);
    if (rows.length > 0) yield rows;
  }
}

export async function readWorksheetRow(
  entry: FileEntry,
  row: number,
  options: {
    readonly sharedStrings: SharedStrings;
    readonly styles: WorkbookStyles;
    readonly limits: StreamLimits;
    readonly signal?: AbortSignal | undefined;
    readonly onProgress?: WorkbookStreamOptions["onProgress"];
  },
): Promise<StreamRow | undefined> {
  for await (const batch of parseWorksheetBatches(entry, {
    firstRow: row,
    lastRow: row,
    batchSize: 1,
    ...options,
  })) {
    return batch[0];
  }
  return undefined;
}
