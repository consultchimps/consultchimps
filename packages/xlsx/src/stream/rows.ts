import type { FileEntry } from "@zip.js/zip.js";
import { SaxesParser } from "saxes";
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
  attribute,
  BoundedXmlText,
  cellRow,
  columnIndex,
  localName,
} from "./xml.js";
import { entryChunks, type StreamLimits } from "./zip.js";

type PendingScalarCell =
  StreamScalarCell | { readonly kind: "shared"; readonly index: number };

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
  readonly type: string | undefined;
  readonly style: number;
  value: string;
  formula: string;
  inline: string;
  valueOpen: boolean;
  formulaOpen: boolean;
  inlineTextOpen: boolean;
  hasValue: boolean;
  hasFormula: boolean;
  valueBytes: number;
  formulaBytes: number;
  inlineBytes: number;
}

function scalarCell(
  current: CurrentCell,
  styles: WorkbookStyles,
): PendingScalarCell {
  const raw = current.type === "inlineStr" ? current.inline : current.value;
  if (current.type === "s") {
    if (!/^\d+$/u.test(raw)) {
      throw new Error(
        `Cell ${current.reference} has an invalid shared-string index.`,
      );
    }
    return { kind: "shared", index: Number(raw) };
  }
  if (current.type === "b") {
    if (raw !== "0" && raw !== "1") {
      throw new Error(
        `Cell ${current.reference} has an invalid Boolean value.`,
      );
    }
    return { kind: "boolean", value: raw === "1" };
  }
  if (current.type === "e") return { kind: "error", error: raw };
  if (current.type === "d") return { kind: "date", raw, iso: raw };
  if (current.type === "str" || current.type === "inlineStr") {
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
    const cells: Record<string, StreamCell> = {};
    for (const [name, cell] of Object.entries(row.cells)) {
      cells[name] = await hydrateCell(cell, sharedStrings);
    }
    result.push({ sourceRow: row.sourceRow, cells });
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
  let nextImplicitRow = 1;
  let nextImplicitColumn = 0;
  let cells: Record<string, PendingCell> = {};
  let current: CurrentCell | undefined;
  let passedLastRow = false;
  const parser = new SaxesParser();
  const decoder = new TextDecoder("utf-8", { fatal: true });
  const encoder = new TextEncoder();
  const textLimit = new BoundedXmlText(
    new Set(["t", "v", "f"]),
    options.limits.maximumCellBytes,
    (element) =>
      `A worksheet ${element} text node exceeds the configured ${options.limits.maximumCellBytes}-byte cell limit.`,
  );

  parser.on("opentag", (tag) => {
    const name = localName(tag.name);
    if (name === "row") {
      const rawRow = attribute(tag, "r");
      activeRow = rawRow === undefined ? nextImplicitRow : Number(rawRow);
      if (
        !Number.isSafeInteger(activeRow) ||
        activeRow < 1 ||
        activeRow > 1_048_576
      ) {
        throw new Error("A worksheet row has an invalid row number.");
      }
      nextImplicitRow = activeRow + 1;
      nextImplicitColumn = 0;
      cells = {};
      if (activeRow > options.lastRow) passedLastRow = true;
    } else if (name === "c") {
      const explicitReference = attribute(tag, "r");
      const parsedColumn = explicitReference
        ? columnIndex(explicitReference)
        : nextImplicitColumn;
      const parsedRow = explicitReference
        ? cellRow(explicitReference)
        : activeRow;
      if (parsedColumn === undefined || parsedRow !== activeRow) {
        throw new Error("A worksheet cell has an invalid reference.");
      }
      nextImplicitColumn = parsedColumn + 1;
      const rawStyle = attribute(tag, "s") ?? "0";
      const style = Number(rawStyle);
      if (!Number.isSafeInteger(style) || style < 0) {
        throw new Error(
          `Cell ${explicitReference ?? "(implicit)"} has an invalid style.`,
        );
      }
      current = {
        reference: explicitReference ?? `${parsedColumn}:${activeRow}`,
        column: parsedColumn,
        type: attribute(tag, "t"),
        style,
        value: "",
        formula: "",
        inline: "",
        valueOpen: false,
        formulaOpen: false,
        inlineTextOpen: false,
        hasValue: false,
        hasFormula: false,
        valueBytes: 0,
        formulaBytes: 0,
        inlineBytes: 0,
      };
    } else if (current && name === "v") {
      current.valueOpen = true;
      current.hasValue = true;
    } else if (current && name === "f") {
      current.formulaOpen = true;
      current.hasFormula = true;
    } else if (current && name === "t" && current.type === "inlineStr") {
      current.inlineTextOpen = true;
    }
  });
  const appendText = (text: string) => {
    if (!current) return;
    const bytes = encoder.encode(text).byteLength;
    if (current.valueOpen) {
      current.valueBytes += bytes;
      if (current.valueBytes > options.limits.maximumCellBytes)
        throw new Error(
          `Cell ${current.reference} value exceeds the configured ${options.limits.maximumCellBytes}-byte limit.`,
        );
      current.value += text;
    }
    if (current.formulaOpen) {
      current.formulaBytes += bytes;
      if (current.formulaBytes > options.limits.maximumCellBytes)
        throw new Error(
          `Cell ${current.reference} formula exceeds the configured ${options.limits.maximumCellBytes}-byte limit.`,
        );
      current.formula += text;
    }
    if (current.inlineTextOpen) {
      current.inlineBytes += bytes;
      if (current.inlineBytes > options.limits.maximumCellBytes)
        throw new Error(
          `Cell ${current.reference} inline string exceeds the configured ${options.limits.maximumCellBytes}-byte limit.`,
        );
      current.inline += text;
    }
  };
  parser.on("text", appendText);
  parser.on("cdata", appendText);
  parser.on("closetag", (tag) => {
    const name = localName(tag.name);
    if (current && name === "v") current.valueOpen = false;
    else if (current && name === "f") current.formulaOpen = false;
    else if (current && name === "t") current.inlineTextOpen = false;
    else if (current && name === "c") {
      const selectedName = options.columns?.get(current.column);
      if (
        activeRow >= options.firstRow &&
        activeRow <= options.lastRow &&
        (options.columns === undefined || selectedName !== undefined)
      ) {
        const scalar = scalarCell(current, options.styles);
        const key = selectedName ?? current.reference;
        cells[key] = current.hasFormula
          ? {
              kind: "formula",
              ...(current.formula === "" ? {} : { formula: current.formula }),
              cached: current.hasValue ? scalar : { kind: "missing" },
            }
          : scalar;
      }
      current = undefined;
    } else if (name === "row") {
      if (activeRow >= options.firstRow && activeRow <= options.lastRow) {
        ready.push({ sourceRow: activeRow, cells });
      }
    }
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
      yield await hydrateRows(
        ready.splice(0, options.batchSize),
        options.sharedStrings,
      );
    }
  }
  if (!passedLastRow) {
    parser.write(decoder.decode());
    parser.close();
  }
  if (ready.length > 0) {
    throwIfAborted(options.signal, "xlsx.stream", "memory");
    yield await hydrateRows(ready, options.sharedStrings);
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
