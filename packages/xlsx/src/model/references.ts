/**
 * A1 reference arithmetic and the row relocation plan the invariant pass is
 * driven by.
 *
 * A `RowRelocation` describes, for every row of one worksheet, where that row
 * ends up: a new row number, or nothing at all. Every dependent structure -
 * cell references, formula text, merged ranges, sqrefs, hyperlinks, table
 * ranges and calculation-chain entries - is adjusted through this one object,
 * which is why a row edit cannot update some of them and forget the rest.
 */

/** The Excel error a reference to a deleted row collapses to. */
export const DELETED_REFERENCE = "#REF!";

export interface CellReference {
  column: number;
  row: number;
}

/** Decode "AB" to a zero-based column index. */
export function decodeColumn(letters: string): number {
  let column = 0;
  for (const character of letters.toUpperCase()) {
    column = column * 26 + (character.charCodeAt(0) - 64);
  }
  return column - 1;
}

/** Encode a zero-based column index to "AB". */
export function encodeColumn(column: number): string {
  let letters = "";
  for (let remaining = column + 1; remaining > 0;) {
    const remainder = (remaining - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    remaining = Math.floor((remaining - remainder) / 26);
  }
  return letters;
}

const CELL_REFERENCE = /^([A-Za-z]{1,3})(\d+)$/u;

/** Decode "D4" to zero-based column and one-based row, or undefined. */
export function decodeCell(reference: string): CellReference | undefined {
  const match = CELL_REFERENCE.exec(reference);
  if (!match?.[1] || !match[2]) {
    return undefined;
  }
  return { column: decodeColumn(match[1]), row: Number(match[2]) };
}

/** Encode zero-based column and one-based row as "D4". */
export function encodeCell(column: number, row: number): string {
  return `${encodeColumn(column)}${row}`;
}

/**
 * A plan mapping every old row of one worksheet to its destination.
 *
 * Rows up to `boundary` are listed explicitly; rows past it are unchanged
 * except for a constant shift, which keeps whole-column ranges such as
 * `D1:D1048576` computable without walking a million rows.
 */
export class RowRelocation {
  readonly #explicit: ReadonlyMap<number, number | null>;
  readonly #boundary: number;
  readonly #shiftBeyond: number;

  private constructor(
    explicit: ReadonlyMap<number, number | null>,
    boundary: number,
    shiftBeyond: number,
  ) {
    this.#explicit = explicit;
    this.#boundary = boundary;
    this.#shiftBeyond = shiftBeyond;
  }

  /**
   * Delete `removed` and, when `renumber` is set, close the gaps so surviving
   * rows keep their relative order without leaving holes.
   */
  static compacting(
    removed: ReadonlySet<number>,
    lastRow: number,
    renumber: boolean,
  ): RowRelocation {
    const explicit = new Map<number, number | null>();
    let deletedBefore = 0;

    for (let row = 1; row <= lastRow; row += 1) {
      if (removed.has(row)) {
        explicit.set(row, null);
        deletedBefore += 1;
        continue;
      }
      explicit.set(row, renumber ? row - deletedBefore : row);
    }

    return new RowRelocation(
      explicit,
      lastRow,
      renumber ? -[...removed].filter((row) => row <= lastRow).length : 0,
    );
  }

  /**
   * An explicit plan: rows listed in `entries` move (or vanish), and every row
   * past the highest listed row stays exactly where it is.
   */
  static explicit(
    entries: Iterable<readonly [number, number | null]>,
  ): RowRelocation {
    const explicit = new Map<number, number | null>(entries);
    const boundary = Math.max(0, ...explicit.keys());
    return new RowRelocation(explicit, boundary, 0);
  }

  /** Where a row ends up, or null when the row is removed. */
  target(row: number): number | null {
    if (row <= this.#boundary) {
      const destination = this.#explicit.get(row);
      // A listed row mapped to null is deleted; only an unlisted row is
      // unchanged, so this cannot collapse into a nullish default.
      return destination === undefined ? row : destination;
    }
    return row + this.#shiftBeyond;
  }

  /** Whether the plan leaves every row exactly where it was. */
  get isIdentity(): boolean {
    if (this.#shiftBeyond !== 0) {
      return false;
    }
    for (const [row, destination] of this.#explicit) {
      if (destination !== row) {
        return false;
      }
    }
    return true;
  }

  /** The rows the plan removes. */
  removedRows(): number[] {
    return [...this.#explicit]
      .filter(([, destination]) => destination === null)
      .map(([row]) => row);
  }

  /**
   * The row span a range covers after relocation, or null when every row it
   * covered was removed. Ranges shrink to their surviving extent rather than
   * keeping stale endpoints.
   */
  mapRowSpan(startRow: number, endRow: number): [number, number] | null {
    let lowest: number | undefined;
    let highest: number | undefined;

    const explicitEnd = Math.min(endRow, this.#boundary);
    for (let row = startRow; row <= explicitEnd; row += 1) {
      const destination = this.target(row);
      if (destination === null) {
        continue;
      }
      lowest =
        lowest === undefined ? destination : Math.min(lowest, destination);
      highest =
        highest === undefined ? destination : Math.max(highest, destination);
    }

    const beyondStart = Math.max(startRow, this.#boundary + 1);
    if (beyondStart <= endRow) {
      const first = beyondStart + this.#shiftBeyond;
      const last = endRow + this.#shiftBeyond;
      lowest = lowest === undefined ? first : Math.min(lowest, first);
      highest = highest === undefined ? last : Math.max(highest, last);
    }

    return lowest === undefined || highest === undefined
      ? null
      : [lowest, highest];
  }
}

/**
 * Relocate an "A1" or "A1:B2" reference, keeping absolute markers. Returns
 * `#REF!` when nothing it pointed at survives.
 */
export function relocateReference(
  reference: string,
  relocation: RowRelocation,
): string {
  const parts = reference.split(":");
  if (parts.length === 1) {
    return relocateSingleReference(parts[0] ?? "", relocation);
  }
  if (parts.length !== 2) {
    return reference;
  }

  const start = parseAnchoredReference(parts[0] ?? "");
  const end = parseAnchoredReference(parts[1] ?? "");
  if (!start || !end) {
    return reference;
  }

  const span = relocation.mapRowSpan(
    Math.min(start.row, end.row),
    Math.max(start.row, end.row),
  );
  if (!span) {
    return DELETED_REFERENCE;
  }

  const ordered = start.row <= end.row ? span : [span[1], span[0]];
  return `${formatAnchored(start, ordered[0]!)}:${formatAnchored(end, ordered[1]!)}`;
}

interface AnchoredReference {
  columnAbsolute: string;
  columnLetters: string;
  rowAbsolute: string;
  row: number;
}

const ANCHORED_REFERENCE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/u;

function parseAnchoredReference(
  reference: string,
): AnchoredReference | undefined {
  const match = ANCHORED_REFERENCE.exec(reference);
  if (!match) {
    return undefined;
  }
  return {
    columnAbsolute: match[1] ?? "",
    columnLetters: match[2] ?? "",
    rowAbsolute: match[3] ?? "",
    row: Number(match[4]),
  };
}

function formatAnchored(reference: AnchoredReference, row: number): string {
  return `${reference.columnAbsolute}${reference.columnLetters}${reference.rowAbsolute}${row}`;
}

function relocateSingleReference(
  reference: string,
  relocation: RowRelocation,
): string {
  const parsed = parseAnchoredReference(reference);
  if (!parsed) {
    return reference;
  }
  const destination = relocation.target(parsed.row);
  return destination === null
    ? DELETED_REFERENCE
    : formatAnchored(parsed, destination);
}

/**
 * Relocate every range in an `sqref` list, dropping the ones whose rows all
 * disappeared. Returns undefined when nothing is left to describe.
 */
export function relocateSqref(
  sqref: string,
  relocation: RowRelocation,
): string | undefined {
  const relocated = sqref
    .split(/\s+/u)
    .filter((reference) => reference !== "")
    .map((reference) => relocateReference(reference, relocation))
    .filter((reference) => !reference.includes(DELETED_REFERENCE));
  return relocated.length === 0 ? undefined : relocated.join(" ");
}

/**
 * Rewrite the unqualified A1 references inside formula text.
 *
 * Sheet-qualified references are deliberately left alone: a formula on another
 * worksheet keeps describing the source workbook's geometry, and rewriting it
 * here would silently restate a cross-sheet aggregate. String literals and
 * structured-reference brackets are copied through untouched.
 */
export function relocateFormulaRows(
  formula: string,
  relocation: RowRelocation,
): string {
  if (relocation.isIdentity) {
    return formula;
  }

  const pieces: string[] = [];
  let index = 0;

  while (index < formula.length) {
    const character = formula[index]!;

    if (character === '"' || character === "'") {
      const end = skipQuoted(formula, index, character);
      pieces.push(formula.slice(index, end));
      index = end;
      continue;
    }
    if (character === "[") {
      const end = skipBrackets(formula, index);
      pieces.push(formula.slice(index, end));
      index = end;
      continue;
    }

    const token = readReferenceToken(formula, index);
    if (token) {
      // A token that is not ours to rewrite is still consumed whole, so the
      // far end of "Data!D4:D9" cannot be picked up as a fresh reference on
      // the next pass.
      pieces.push(
        token.rewritable
          ? relocateReference(token.text, relocation)
          : token.text,
      );
      index = token.end;
      continue;
    }

    pieces.push(character);
    index += 1;
  }

  return pieces.join("");
}

function skipQuoted(text: string, start: number, quote: string): number {
  for (let index = start + 1; index < text.length; index += 1) {
    if (text[index] !== quote) {
      continue;
    }
    // A doubled quote is an escaped quote, not the end of the literal.
    if (text[index + 1] === quote) {
      index += 1;
      continue;
    }
    return index + 1;
  }
  return text.length;
}

function skipBrackets(text: string, start: number): number {
  let depth = 0;
  for (let index = start; index < text.length; index += 1) {
    if (text[index] === "[") {
      depth += 1;
    } else if (text[index] === "]") {
      depth -= 1;
      if (depth === 0) {
        return index + 1;
      }
    }
  }
  return text.length;
}

const REFERENCE_TOKEN = /^\$?[A-Za-z]{1,3}\$?\d+/u;
// A character that makes a match part of a larger name rather than a reference.
const IDENTIFIER_BEFORE = /[A-Za-z0-9_.$!]/u;
const IDENTIFIER_AFTER = /[A-Za-z0-9_(.[]/u;

function readReferenceToken(
  formula: string,
  start: number,
): { end: number; text: string; rewritable: boolean } | undefined {
  const first = REFERENCE_TOKEN.exec(formula.slice(start))?.[0];
  if (!first) {
    return undefined;
  }
  let end = start + first.length;

  // A range is one token: its endpoints must move together, not separately.
  if (formula[end] === ":") {
    const second = REFERENCE_TOKEN.exec(formula.slice(end + 1))?.[0];
    if (second) {
      end = end + 1 + second.length;
    }
  }

  const following = formula[end];
  if (following !== undefined && IDENTIFIER_AFTER.test(following)) {
    // Part of a longer name: a function like LOG10( or a table like Tbl1[.
    return undefined;
  }

  // A reference behind "!" belongs to another sheet or an external workbook,
  // and one behind an identifier character is part of a longer name. Both are
  // consumed whole but copied through unchanged.
  const previous = formula[start - 1];
  const rewritable =
    previous === undefined || !IDENTIFIER_BEFORE.test(previous);

  return { end, rewritable, text: formula.slice(start, end) };
}
