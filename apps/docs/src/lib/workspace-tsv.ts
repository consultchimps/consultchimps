/**
 * The clipboard grammar the record grid reads and writes, in one place.
 *
 * Copy and paste are the same grammar asked in two directions, so they are one
 * module and one set of rules. It is Excel's own: cells separated by tabs, rows
 * separated by line breaks, and a field wrapped in double quotes when it holds
 * one of the characters that would otherwise end it, with an internal quote
 * doubled. A block encoded here pastes into Excel as the block it is, and a
 * block copied from Excel parses back here as the block it was.
 *
 * Two rules are the whole reason this is ours rather than the grid library's.
 *
 * **Line endings.** Excel on Windows writes CRLF. Tabulator's built-in range
 * parser splits on "\n" and leaves the "\r" attached to the last field of every
 * row, so a multi-row paste quietly corrupts its last column. This parser
 * accepts CRLF, a bare CR, and a bare LF, in any mix, and drops one trailing
 * separator, because a copy usually ends with one and it is not an empty row.
 *
 * **Quoting.** A text column here may hold a tab or a newline, and the built-in
 * copy joins with those characters and quotes nothing, so such a cell would
 * corrupt the clipboard on the way out as surely as the "\r" corrupts it on the
 * way in. Encoding and parsing quotes is what makes a round trip a round trip.
 *
 * What a cell contributes is its stored value's text, never its rendered label.
 * A foreign-key cell copies the Record ID it stores rather than the name it
 * shows, because the Record ID is what pastes back into that column and
 * resolves; the label would be refused. That is `cellText`, which the fill
 * handle uses as well, so one gesture cannot read a cell differently from
 * another.
 */
import type { CellValue } from "@consultchimps/tabular";

/** Cells across a row, joined by this. */
const CELL_SEPARATOR = "\t";

/**
 * Rows joined by this. CRLF because that is what Excel writes and what every
 * spreadsheet reads; the parser below is the lenient half of the pair.
 */
const ROW_SEPARATOR = "\r\n";

/** The characters that make a field ambiguous unless it is quoted. */
const NEEDS_QUOTING = /["\t\r\n]/u;

/**
 * How a stored value reads on the clipboard, and to the fill handle.
 *
 * An empty cell is empty text rather than the word "null": that is what a
 * spreadsheet shows and what pastes back as an empty cell.
 */
export function cellText(value: CellValue | undefined): string {
  return value === null || value === undefined ? "" : String(value);
}

/** Encode one field, quoting it only when it would otherwise be ambiguous. */
function encodeField(field: string): string {
  if (!NEEDS_QUOTING.test(field)) {
    return field;
  }
  return `"${field.replaceAll('"', '""')}"`;
}

/**
 * Encode a rectangular block as clipboard text. Rows are given top to bottom
 * and cells left to right, exactly as they sit in the grid.
 */
export function encodeTsv(rows: ReadonlyArray<readonly string[]>): string {
  return rows
    .map((row) => row.map(encodeField).join(CELL_SEPARATOR))
    .join(ROW_SEPARATOR);
}

/**
 * Parse clipboard text into a block of fields.
 *
 * Written as a scan rather than a split because a quoted field may contain both
 * separators, so no amount of splitting can find the boundaries first. Anything
 * the grammar does not describe is read the way a spreadsheet reads it rather
 * than refused: a quote inside an unquoted field is a literal quote (`12" pipe`
 * is a length, not a mistake), text after a closing quote continues the same
 * field, and a quote that is never closed takes the rest of the text with it.
 * Refusing those would turn an ordinary paste into an error message.
 */
export function parseTsv(text: string): string[][] {
  if (text === "") {
    return [];
  }

  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  let index = 0;

  const endField = (): void => {
    row.push(field);
    field = "";
  };
  const endRow = (): void => {
    endField();
    rows.push(row);
    row = [];
  };

  while (index < text.length) {
    const character = text[index] as string;

    if (quoted) {
      if (character === '"') {
        // A doubled quote is one quote; a single one closes the field.
        if (text[index + 1] === '"') {
          field += '"';
          index += 2;
          continue;
        }
        quoted = false;
        index += 1;
        continue;
      }
      field += character;
      index += 1;
      continue;
    }

    if (character === '"' && field === "") {
      quoted = true;
      index += 1;
      continue;
    }
    if (character === CELL_SEPARATOR) {
      endField();
      index += 1;
      continue;
    }
    if (character === "\r" || character === "\n") {
      endRow();
      // CRLF is one separator, not two.
      index += character === "\r" && text[index + 1] === "\n" ? 2 : 1;
      continue;
    }
    field += character;
    index += 1;
  }

  // Whatever is left is the last field, unless the text ended on a separator:
  // a copy usually ends with one, and it means the end of the last row rather
  // than an empty row after it.
  if (field !== "" || row.length > 0 || rows.length === 0) {
    endRow();
  }

  return rows;
}
