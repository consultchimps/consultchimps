/**
 * Rewriting the names a formula refers to.
 *
 * A transplanted worksheet keeps its cells exactly where they were, so no cell
 * reference moves. What does move is the NAME a reference is qualified by: a
 * merge renames colliding worksheets and colliding Excel Tables, and every
 * formula, defined name, conditional-formatting rule and data-validation
 * constraint that named them has to follow. Leaving them alone is worse than
 * dropping them, because `SUM(Data!D4:D9)` in the second workbook's Summary
 * would silently start reading the FIRST workbook's Data.
 *
 * The scanner walks the formula once, skipping string literals so a quoted
 * `"Data!"` inside a formula is never mistaken for a reference.
 */
import { decodeXmlText, editElements } from "../model/xml.js";
import { escapeXmlText } from "./ooxml.js";

/** Renames a transplant has to propagate into formula text. */
export interface NameRewrites {
  /** Lower-cased original worksheet name to its name in the merged workbook. */
  readonly sheets: ReadonlyMap<string, string>;
  /** Lower-cased original table name to its name in the merged workbook. */
  readonly tables: ReadonlyMap<string, string>;
}

export function hasRewrites(rewrites: NameRewrites): boolean {
  return rewrites.sheets.size > 0 || rewrites.tables.size > 0;
}

/** A cell reference in disguise, which a bare sheet name must never be. */
const CELL_REFERENCE_SHAPED = /^[A-Za-z]{1,3}\d+$/u;
const BARE_SHEET_NAME = /^[A-Za-z_][A-Za-z0-9_.]*$/u;
const RESERVED_BARE_NAMES = new Set(["c", "r", "true", "false"]);

// Sheet, table and defined names may use letters from any script, so the
// identifier classes are expressed as Unicode letter properties, not as ASCII.
const IDENTIFIER_START = /[\p{L}_\\]/u;
const IDENTIFIER_PART = /[\p{L}\p{Nd}_.\\]/u;

/** A worksheet name as a formula writes it, quoted only when it has to be. */
export function encodeSheetName(name: string): string {
  if (
    BARE_SHEET_NAME.test(name) &&
    !CELL_REFERENCE_SHAPED.test(name) &&
    !RESERVED_BARE_NAMES.has(name.toLocaleLowerCase())
  ) {
    return name;
  }
  return `'${name.replaceAll("'", "''")}'`;
}

interface ScannedName {
  /** The name with its quoting removed. */
  readonly name: string;
  /** Offset just past the name in the source text. */
  readonly end: number;
}

/** Read a `'quoted name'` at `start`, where `''` stands for one quote. */
function readQuotedName(text: string, start: number): ScannedName | undefined {
  let index = start + 1;
  let name = "";
  while (index < text.length) {
    if (text[index] === "'") {
      if (text[index + 1] === "'") {
        name += "'";
        index += 2;
        continue;
      }
      return { name, end: index + 1 };
    }
    name += text[index];
    index += 1;
  }
  return undefined;
}

/** Read a bare identifier at `start`, or nothing when none begins there. */
function readBareName(text: string, start: number): ScannedName | undefined {
  const first = text[start];
  if (first === undefined || !IDENTIFIER_START.test(first)) {
    return undefined;
  }
  let index = start + 1;
  while (index < text.length && IDENTIFIER_PART.test(text[index] ?? "")) {
    index += 1;
  }
  return { name: text.slice(start, index), end: index };
}

function readName(text: string, start: number): ScannedName | undefined {
  return text[start] === "'"
    ? readQuotedName(text, start)
    : readBareName(text, start);
}

/** Skip a `"literal"`, in which `""` stands for one quotation mark. */
function skipStringLiteral(text: string, start: number): number {
  let index = start + 1;
  while (index < text.length) {
    if (text[index] === '"') {
      if (text[index + 1] === '"') {
        index += 2;
        continue;
      }
      return index + 1;
    }
    index += 1;
  }
  return text.length;
}

/**
 * Rewrite every worksheet-qualified and table-qualified name in one formula.
 * Returns the text unchanged when no name it mentions was renamed, so a
 * transplant that renamed nothing leaves its formulas byte-identical.
 */
export function rewriteFormulaNames(
  formula: string,
  rewrites: NameRewrites,
): string {
  if (!hasRewrites(rewrites)) {
    return formula;
  }

  const pieces: string[] = [];
  let index = 0;
  let changed = false;

  const renamedSheet = (name: string): string | undefined =>
    rewrites.sheets.get(name.toLocaleLowerCase());

  while (index < formula.length) {
    const character = formula[index] ?? "";

    if (character === '"') {
      const end = skipStringLiteral(formula, index);
      pieces.push(formula.slice(index, end));
      index = end;
      continue;
    }

    const scanned = readName(formula, index);
    if (!scanned) {
      pieces.push(character);
      index += 1;
      continue;
    }

    // `Sheet!ref` and `'Sheet Name'!ref`.
    if (formula[scanned.end] === "!") {
      const renamed = renamedSheet(scanned.name);
      if (renamed === undefined) {
        pieces.push(formula.slice(index, scanned.end + 1));
      } else {
        pieces.push(`${encodeSheetName(renamed)}!`);
        changed = true;
      }
      index = scanned.end + 1;
      continue;
    }

    // A 3-D reference spanning a sheet range: `First:Last!ref`.
    const second =
      formula[scanned.end] === ":"
        ? readName(formula, scanned.end + 1)
        : undefined;
    if (second && formula[second.end] === "!") {
      const first = renamedSheet(scanned.name);
      const last = renamedSheet(second.name);
      if (first === undefined && last === undefined) {
        pieces.push(formula.slice(index, second.end + 1));
      } else {
        pieces.push(
          `${encodeSheetName(first ?? scanned.name)}:${encodeSheetName(
            last ?? second.name,
          )}!`,
        );
        changed = true;
      }
      index = second.end + 1;
      continue;
    }

    // A structured reference: `TableName[...]`. Only a bare name can carry one.
    if (formula[scanned.end] === "[" && formula[index] !== "'") {
      const renamed = rewrites.tables.get(scanned.name.toLocaleLowerCase());
      if (renamed !== undefined) {
        pieces.push(renamed);
        changed = true;
        index = scanned.end;
        continue;
      }
    }

    pieces.push(formula.slice(index, scanned.end));
    index = scanned.end;
  }

  return changed ? pieces.join("") : formula;
}

/** Elements whose text content is a formula in the parts a merge transplants. */
const FORMULA_ELEMENTS = [
  "f",
  "formula",
  "formula1",
  "formula2",
  "calculatedColumnFormula",
  "totalsRowFormula",
] as const;

/**
 * Rewrite the names in every formula-bearing element of one part. The text is
 * decoded before scanning and re-escaped only when the scan changed it, so a
 * part whose formulas mention nothing renamed travels through byte-identical.
 */
export function rewriteFormulaElements(
  xml: string,
  rewrites: NameRewrites,
): string {
  if (!hasRewrites(rewrites)) {
    return xml;
  }

  let rewritten = xml;
  for (const localName of FORMULA_ELEMENTS) {
    rewritten = editElements(rewritten, localName, (element, text) => {
      if (element.selfClosing) {
        return text;
      }
      const innerStart = element.innerStart - element.start;
      const innerEnd = element.innerEnd - element.start;
      const decoded = decodeXmlText(text.slice(innerStart, innerEnd));
      const updated = rewriteFormulaNames(decoded, rewrites);
      if (updated === decoded) {
        return text;
      }
      return `${text.slice(0, innerStart)}${escapeXmlText(updated)}${text.slice(
        innerEnd,
      )}`;
    });
  }
  return rewritten;
}
