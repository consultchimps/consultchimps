/**
 * L1: a worksheet part read as the document model reads it, one row at a time
 * (ADR 0006).
 *
 * `WorksheetModel.parse` divides a part into the text before `sheetData`'s
 * content, the rows and the text between them, and the text from `sheetData`'s
 * closing tag on. This reads the same division from a stream, with the same
 * scanning rules as `findElement` (markup constructs skipped, quoted `>` inside
 * an opening tag, same-name nesting counted), so each row arrives as the
 * `WorksheetRow` the model would have built and nothing but the row being read
 * is held. The text around the rows is handed over whole: it holds no rows.
 */
import type { ZipReader } from "../package/zip-reader.js";
import { WorksheetRow } from "./worksheet-model.js";
import { getAttribute } from "./xml.js";

/**
 * One row as written. Its number is read from its opening tag; the full parse
 * the model makes of it happens only when asked for, since a reader that drops
 * most rows has no use for theirs.
 */
export class StreamedRow {
  readonly text: string;
  /** The number the model gives the row when it carries none. */
  readonly implied: number;
  readonly number: number;
  #parsed: WorksheetRow | undefined;

  constructor(text: string, openTag: string, implied: number) {
    this.text = text;
    this.implied = implied;
    // Read, not judged: a number the model refuses is refused when the row is
    // parsed, so a reader that copies rows as written copies this one too.
    const written = getAttribute(openTag, "r");
    this.number = written === undefined ? implied : Number(written);
  }

  /** The row as the model parses it, refused as the model refuses it. */
  parse(): WorksheetRow {
    this.#parsed ??= new WorksheetRow(this.text, this.implied);
    return this.#parsed;
  }
}

/** Receives a worksheet part in document order. */
export interface WorksheetPartHandlers {
  /** Everything up to the end of `sheetData`'s opening tag, once, first. */
  prefix(text: string): void;
  /** One row, in the order the part lists them. */
  row(row: StreamedRow): void;
  /** Text between rows inside `sheetData`, as the model keeps it. */
  text(text: string): void;
  /** Everything from `sheetData`'s closing tag on, once, last. */
  suffix(text: string): void;
}

const NEED_MORE = Symbol("need more");
type Scan<T> = T | typeof NEED_MORE;

function localNameOf(name: string): string {
  const colon = name.indexOf(":");
  return (colon < 0 ? name : name.slice(colon + 1)).toLocaleLowerCase();
}

/** Whether `text` from `start` could still turn into `literal`. */
function couldStart(text: string, start: number, literal: string): boolean {
  const available = text.length - start;
  return available < literal.length && literal.startsWith(text.slice(start));
}

/**
 * Where a markup construct at `start` ends, as `skipMarkupConstruct` reads it,
 * undefined when none starts there.
 */
function skipMarkup(
  text: string,
  start: number,
  eof: boolean,
): Scan<number | undefined> {
  // Every construct starts "<!" or "<?"; anything else is none.
  const next = text.charCodeAt(start + 1);
  if (next !== 0x21 && next !== 0x3f) {
    return start + 1 >= text.length && !eof ? NEED_MORE : undefined;
  }
  const constructs: Array<[string, string, number]> = [
    ["<!--", "-->", 4],
    ["<![CDATA[", "]]>", 9],
    ["<?", "?>", 2],
    ["<!", ">", 2],
  ];
  for (const [open, close, skip] of constructs) {
    if (text.startsWith(open, start)) {
      const end = text.indexOf(close, start + skip);
      if (end >= 0) return end + close.length;
      return eof ? text.length : NEED_MORE;
    }
    if (!eof && couldStart(text, start, open)) return NEED_MORE;
  }
  return undefined;
}

interface OpenTag {
  readonly end: number;
  readonly name: string;
  readonly selfClosing: boolean;
}

const TAG_NAME = /[^\s/>!?][^\s/>]*/uy;

/** An opening tag at `start`, as `readOpenTag` reads one. */
function readOpenTag(
  text: string,
  start: number,
  eof: boolean,
): Scan<OpenTag | undefined> {
  TAG_NAME.lastIndex = start + 1;
  const match = TAG_NAME.exec(text);
  if (!match) {
    // A name that reaches the end of what is buffered may go on.
    return !eof && text.length - start < 2 ? NEED_MORE : undefined;
  }
  const name = match[0];
  let index = start + 1 + name.length;
  if (index >= text.length && !eof) return NEED_MORE;
  let quote: string | undefined;
  while (index < text.length) {
    const character = text[index];
    if (quote) {
      if (character === quote) quote = undefined;
    } else if (character === '"' || character === "'") {
      quote = character;
    } else if (character === ">") {
      return {
        end: index + 1,
        name,
        selfClosing: text[index - 1] === "/",
      };
    }
    index += 1;
  }
  return eof ? undefined : NEED_MORE;
}

/**
 * Where the element opened by `name` before `from` closes, as
 * `findMatchingClose` finds it; undefined when it never does.
 */
function matchingClose(
  text: string,
  from: number,
  name: string,
  eof: boolean,
): Scan<number | undefined> {
  const closeTag = `</${name}>`;
  // The usual case answered directly. The scan below only finds a closing tag
  // in the buffer, so without one there is nothing to find yet. With one, and
  // no markup construct or same-name element before it, the scan stops there:
  // in well-formed XML no attribute value holds a "<".
  const first = text.indexOf(closeTag, from);
  if (first < 0 && !eof) return NEED_MORE;
  if (first >= 0) {
    const between = text.slice(from, first);
    if (
      !between.includes("<!") &&
      !between.includes("<?") &&
      !between.includes(`<${name}`)
    ) {
      return first;
    }
  }
  let depth = 1;
  for (let index = text.indexOf("<", from); index >= 0;) {
    const skipped = skipMarkup(text, index, eof);
    if (skipped === NEED_MORE) return NEED_MORE;
    if (skipped !== undefined) {
      index = text.indexOf("<", skipped);
      continue;
    }
    if (text.startsWith(closeTag, index)) {
      depth -= 1;
      if (depth === 0) return index;
      index = text.indexOf("<", index + closeTag.length);
      continue;
    }
    if (!eof && couldStart(text, index, closeTag)) return NEED_MORE;
    const openTag = readOpenTag(text, index, eof);
    if (openTag === NEED_MORE) return NEED_MORE;
    if (openTag) {
      if (openTag.name === name && !openTag.selfClosing) depth += 1;
      index = text.indexOf("<", openTag.end);
      continue;
    }
    index = text.indexOf("<", index + 1);
  }
  return eof ? undefined : NEED_MORE;
}

/** Delivers a part's bytes in order to `onData`. */
export type PartFeed = (onData: (chunk: Uint8Array) => void) => Promise<void>;

/**
 * A part read from a package. `between` runs after each compressed chunk is
 * read, so a long part yields to the event loop.
 */
export function zipPartFeed(
  zip: ZipReader,
  part: string,
  between?: () => Promise<void>,
): PartFeed {
  return (onData) => zip.stream(part, onData, between);
}

/** A part already in memory. */
export function bytesPartFeed(bytes: Uint8Array): PartFeed {
  return (onData) => {
    onData(bytes);
    return Promise.resolve();
  };
}

/** Read a worksheet part as `WorksheetModel.parse` divides it. */
export async function readWorksheetPart(
  feed: PartFeed,
  handlers: WorksheetPartHandlers,
): Promise<void> {
  const decoder = new TextDecoder();
  let buffer = "";
  let state: "prefix" | "rows" | "suffix" = "prefix";
  // In the prefix: where scanning resumes. In the rows: where the next
  // segment starts; text before it has been handed over.
  let cursor = 0;
  let scan = 0;
  let sheetDataName = "";
  let sheetDataDepth = 1;
  let impliedRowNumber = 1;
  let suffix: string[] = [];

  const process = (eof: boolean): void => {
    if (state === "prefix") {
      for (let index = buffer.indexOf("<", scan); index >= 0;) {
        const skipped = skipMarkup(buffer, index, eof);
        if (skipped === NEED_MORE) {
          scan = index;
          return;
        }
        if (skipped !== undefined) {
          index = buffer.indexOf("<", skipped);
          continue;
        }
        const openTag = readOpenTag(buffer, index, eof);
        if (openTag === NEED_MORE) {
          scan = index;
          return;
        }
        if (!openTag) {
          index = buffer.indexOf("<", index + 1);
          continue;
        }
        if (localNameOf(openTag.name) !== "sheetdata") {
          index = buffer.indexOf("<", openTag.end);
          continue;
        }
        handlers.prefix(buffer.slice(0, openTag.end));
        buffer = buffer.slice(openTag.end);
        cursor = 0;
        scan = 0;
        if (openTag.selfClosing) {
          state = "suffix";
          suffix = [buffer];
          buffer = "";
          return;
        }
        state = "rows";
        sheetDataName = openTag.name;
        break;
      }
      if (state === "prefix") {
        if (eof) {
          throw new Error("Worksheet package part does not contain sheetData.");
        }
        // Keep a possible partial tag; everything before it has been scanned.
        const last = buffer.lastIndexOf("<");
        scan = last >= 0 ? last : buffer.length;
        return;
      }
    }

    if (state === "rows") {
      const closeTag = `</${sheetDataName}>`;
      let index = buffer.indexOf("<", scan);
      while (index >= 0) {
        const skipped = skipMarkup(buffer, index, eof);
        if (skipped === NEED_MORE) break;
        if (skipped !== undefined) {
          index = buffer.indexOf("<", skipped);
          continue;
        }
        if (buffer.startsWith(closeTag, index)) {
          sheetDataDepth -= 1;
          if (sheetDataDepth === 0) {
            if (index > cursor) handlers.text(buffer.slice(cursor, index));
            state = "suffix";
            suffix = [buffer.slice(index)];
            buffer = "";
            return;
          }
          index = buffer.indexOf("<", index + closeTag.length);
          continue;
        }
        if (!eof && couldStart(buffer, index, closeTag)) break;
        const openTag = readOpenTag(buffer, index, eof);
        if (openTag === NEED_MORE) break;
        if (!openTag) {
          index = buffer.indexOf("<", index + 1);
          continue;
        }
        if (openTag.name === sheetDataName && !openTag.selfClosing) {
          sheetDataDepth += 1;
        }
        if (localNameOf(openTag.name) !== "row") {
          index = buffer.indexOf("<", openTag.end);
          continue;
        }
        let end: number;
        if (openTag.selfClosing) {
          end = openTag.end;
        } else {
          const close = matchingClose(buffer, openTag.end, openTag.name, eof);
          if (close === NEED_MORE) break;
          if (close === undefined) {
            // Never closed: the model reads no further rows, and what is
            // left is text up to sheetData's end.
            index = buffer.indexOf("<", openTag.end);
            continue;
          }
          end = close + `</${openTag.name}>`.length;
        }
        if (index > cursor) handlers.text(buffer.slice(cursor, index));
        const row = new StreamedRow(
          buffer.slice(index, end),
          buffer.slice(index, openTag.end),
          impliedRowNumber,
        );
        impliedRowNumber = row.number + 1;
        handlers.row(row);
        cursor = end;
        index = buffer.indexOf("<", end);
      }
      if (eof) {
        throw new Error("Worksheet package part does not contain sheetData.");
      }
      // Drop what has been handed over, keep the rest.
      const keepFrom = Math.min(cursor, index >= 0 ? index : buffer.length);
      buffer = buffer.slice(keepFrom);
      cursor -= keepFrom;
      scan = index >= 0 ? index - keepFrom : buffer.length;
      return;
    }

    suffix.push(buffer);
    buffer = "";
  };

  await feed((chunk) => {
    buffer += decoder.decode(chunk, { stream: true });
    process(false);
  });
  buffer += decoder.decode();
  process(true);
  // The closures above move the state on, which the checker cannot follow.
  if ((state as string) !== "suffix") {
    throw new Error("Worksheet package part does not contain sheetData.");
  }
  handlers.suffix(suffix.join(""));
}
