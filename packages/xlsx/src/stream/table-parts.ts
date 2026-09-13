import type { FileEntry } from "@zip.js/zip.js";

import type { WorkbookStreamOptions } from "./types.js";
import { createElementParser, relationshipAttribute } from "./xml-elements.js";
import { entryChunks } from "./zip.js";

const MAXIMUM_WORKSHEET_TAG_BYTES = 64 * 1024;

function asciiEquals(
  bytes: readonly number[],
  start: number,
  end: number,
  expected: string,
): boolean {
  if (end - start !== expected.length) return false;
  for (let index = 0; index < expected.length; index += 1) {
    if (bytes[start + index] !== expected.charCodeAt(index)) return false;
  }
  return true;
}

class TablePartScanner {
  #mode: "text" | "tag" | "comment" | "cdata" | "instruction" = "text";
  #tag: number[] = [];
  #quote: number | undefined;
  #previousByte = 0;
  #penultimateByte = 0;
  readonly #activeRelationshipIds: string[] = [];
  readonly #seenRelationshipIds = new Set<string>();
  readonly #decoder = new TextDecoder("utf-8", { fatal: true });
  readonly #parser;

  constructor(private readonly relationshipIds: ReadonlySet<string>) {
    this.#parser = createElementParser({
      root: "worksheet",
      open: (tag, path) => {
        if (!path.is("worksheet", "tableParts", "tablePart")) return;
        const id = relationshipAttribute(tag, path);
        if (!id) {
          throw new Error(
            "A worksheet tablePart is missing its relationship ID.",
          );
        }
        if (!this.relationshipIds.has(id)) {
          throw new Error(
            `A worksheet tablePart references missing relationship "${id}".`,
          );
        }
        if (this.#seenRelationshipIds.has(id)) {
          throw new Error(
            `Worksheet table relationship "${id}" is referenced more than once.`,
          );
        }
        this.#seenRelationshipIds.add(id);
        this.#activeRelationshipIds.push(id);
      },
    });
  }

  consume(chunk: Uint8Array): void {
    for (const byte of chunk) this.#consumeByte(byte);
  }

  finish(): readonly string[] {
    if (this.#mode !== "text") {
      throw new Error(
        "A worksheet ended inside XML markup while reading its table references.",
      );
    }
    // Row reading validates worksheet completeness; this scan discovers references.
    return this.#activeRelationshipIds;
  }

  #consumeByte(byte: number): void {
    if (
      this.#mode === "comment" ||
      this.#mode === "cdata" ||
      this.#mode === "instruction"
    ) {
      this.#skipMarkup(byte);
      return;
    }
    if (this.#mode === "text") {
      if (byte === 0x3c) {
        this.#mode = "tag";
        this.#tag = [byte];
        this.#quote = undefined;
      }
      return;
    }

    this.#tag.push(byte);
    if (this.#tag.length > MAXIMUM_WORKSHEET_TAG_BYTES) {
      throw new Error(
        `A worksheet tag exceeds the ${MAXIMUM_WORKSHEET_TAG_BYTES}-byte parser limit.`,
      );
    }
    if (this.#tag.length === 4 && asciiEquals(this.#tag, 0, 4, "<!--")) {
      this.#startSkipping("comment");
      return;
    }
    if (this.#tag.length === 9 && asciiEquals(this.#tag, 0, 9, "<![CDATA[")) {
      this.#startSkipping("cdata");
      return;
    }
    if (this.#tag.length === 2 && asciiEquals(this.#tag, 0, 2, "<?")) {
      this.#startSkipping("instruction");
      return;
    }
    if (
      this.#tag.length === 9 &&
      String.fromCharCode(...this.#tag).toUpperCase() === "<!DOCTYPE"
    ) {
      throw new Error(
        "Document type declarations are not allowed in XLSX XML.",
      );
    }
    if (this.#quote !== undefined) {
      if (byte === this.#quote) this.#quote = undefined;
      return;
    }
    if (byte === 0x22 || byte === 0x27) {
      this.#quote = byte;
      return;
    }
    if (byte !== 0x3e) return;
    this.#readTag(this.#tag);
    this.#mode = "text";
    this.#tag = [];
  }

  #startSkipping(mode: "comment" | "cdata" | "instruction"): void {
    this.#mode = mode;
    this.#tag = [];
    this.#previousByte = 0;
    this.#penultimateByte = 0;
  }

  #skipMarkup(byte: number): void {
    const finished =
      this.#mode === "instruction"
        ? this.#previousByte === 0x3f && byte === 0x3e
        : this.#mode === "comment"
          ? this.#penultimateByte === 0x2d &&
            this.#previousByte === 0x2d &&
            byte === 0x3e
          : this.#penultimateByte === 0x5d &&
            this.#previousByte === 0x5d &&
            byte === 0x3e;
    this.#penultimateByte = this.#previousByte;
    this.#previousByte = byte;
    if (finished) {
      this.#mode = "text";
      this.#previousByte = 0;
      this.#penultimateByte = 0;
    }
  }

  #readTag(bytes: readonly number[]): void {
    this.#parser.write(this.#decoder.decode(new Uint8Array(bytes)));
  }
}

export async function activeTableRelationshipIds(
  entry: FileEntry,
  relationshipIds: ReadonlySet<string>,
  options: WorkbookStreamOptions,
): Promise<readonly string[]> {
  const scanner = new TablePartScanner(relationshipIds);
  for await (const chunk of entryChunks(entry, {
    signal: options.signal,
    onProgress: options.onProgress,
    stage: "worksheet-metadata",
  })) {
    scanner.consume(chunk);
  }
  return scanner.finish();
}
