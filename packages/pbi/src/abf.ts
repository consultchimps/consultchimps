import type { PipelineBudget } from "./budget.js";
import { unreadableModel } from "./errors.js";

/**
 * The Analysis Services backup container: a header page, a UTF-8 virtual
 * directory, and a UTF-16 backup log that names the VertiPaq members. The XML
 * is flat and entity-light, and `DOMParser` does not exist in a worker, so a
 * tokenizer reads it. Unlike the reference implementation this one is bounded:
 * a crafted header cannot expand a node tree without limit, and every member
 * offset is checked against the image before it is sliced.
 */

const utf8 = new TextDecoder("utf-8");
const utf16 = new TextDecoder("utf-16le");

/** The signature and BOM occupy the first 72 bytes; the header page ends at 4 KiB. */
const HEADER_OFFSET = 72;
const HEADER_PAGE = 0x1000;

/** Ceilings so a crafted header cannot expand the node tree without bound. */
const MAX_NODES = 200_000;
const MAX_DEPTH = 64;
/** What one parsed element costs: the object, its two fields and its array. */
const NODE_BYTES = 128;

export interface XmlNode {
  readonly tag: string;
  readonly children: XmlNode[];
  text: string;
}

const ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

function unescape(text: string): string {
  if (!text.includes("&")) return text;
  return text.replace(
    /&(#x?[0-9a-fA-F]+|amp|lt|gt|quot|apos);/g,
    (match, group: string) => {
      if (group[0] !== "#") return ENTITIES[group] ?? match;
      const hex = group[1] === "x";
      const point = Number.parseInt(
        hex ? group.slice(2) : group.slice(1),
        hex ? 16 : 10,
      );
      return Number.isInteger(point) && point >= 0 && point <= 0x10ffff
        ? String.fromCodePoint(point)
        : match;
    },
  );
}

/**
 * A bounded element tree. Processing instructions, comments and DTDs are
 * skipped. When a budget is supplied the tree is charged to it as it grows, so
 * the accounting covers what the parser actually holds rather than only the
 * ceilings refusing the absurd case.
 */
export function parseXml(source: string, budget?: PipelineBudget): XmlNode {
  let cursor = 0;
  const length = source.length;
  const root: XmlNode = { tag: "#root", children: [], text: "" };
  const stack: XmlNode[] = [root];
  let nodes = 0;
  while (cursor < length) {
    const open = source.indexOf("<", cursor);
    if (open < 0) break;
    if (open > cursor) {
      const text = source.slice(cursor, open);
      const top = stack[stack.length - 1]!;
      if (text.trim()) top.text += text;
    }
    const close = source.indexOf(">", open);
    if (close < 0) break;
    const raw = source.slice(open + 1, close);
    cursor = close + 1;
    if (raw[0] === "?" || raw[0] === "!") continue;
    if (raw[0] === "/") {
      if (stack.length > 1) stack.pop();
      continue;
    }
    const selfClosing = raw.endsWith("/");
    const body = selfClosing ? raw.slice(0, -1) : raw;
    const space = body.search(/\s/);
    const element: XmlNode = {
      tag: space < 0 ? body : body.slice(0, space),
      children: [],
      text: "",
    };
    if (++nodes > MAX_NODES) throw unreadableModel("backup");
    budget?.reserve("backupHeaderNodes", NODE_BYTES, "backup");
    stack[stack.length - 1]!.children.push(element);
    if (!selfClosing) {
      if (stack.length >= MAX_DEPTH) throw unreadableModel("backup");
      stack.push(element);
    }
  }
  return root.children[0] ?? root;
}

function childText(element: XmlNode, tag: string): string | null {
  for (const child of element.children)
    if (child.tag === tag) return unescape(child.text);
  return null;
}

function childrenNamed(element: XmlNode, tag: string): XmlNode[] {
  return element.children.filter((child) => child.tag === tag);
}

function descend(element: XmlNode, path: string): XmlNode[] {
  let current = [element];
  for (const segment of path.split("/")) {
    const next: XmlNode[] = [];
    for (const node of current)
      for (const child of node.children)
        if (child.tag === segment) next.push(child);
    current = next;
  }
  return current;
}

/** UTF-16LE text that may carry a BOM and trailing NUL padding. */
function utf16Text(bytes: Uint8Array): string {
  let view = bytes;
  if (view.length >= 2 && view[0] === 0xff && view[1] === 0xfe)
    view = view.subarray(2);
  const text = utf16.decode(view);
  const nul = text.indexOf("\0");
  return nul >= 0 ? text.slice(0, nul) : text;
}

export interface AbfMember {
  /** The member path relative to the backup's persist root. */
  readonly path: string;
  readonly fileName: string;
  readonly storagePath: string;
  readonly size: number;
  readonly offset: number;
}

export interface AbfImage {
  readonly members: readonly AbfMember[];
  /** True when every member carries a four-byte trailer. */
  readonly errorCode: boolean;
}

function positiveInteger(value: string | null): number {
  if (value === null) throw unreadableModel("backup");
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < 0)
    throw unreadableModel("backup");
  return parsed;
}

/** Read the backup's virtual directory and backup log into a member list. */
export function parseAbf(image: Uint8Array, budget?: PipelineBudget): AbfImage {
  if (image.length <= HEADER_PAGE) throw unreadableModel("backup");
  const header = parseXml(
    utf16Text(image.subarray(HEADER_OFFSET, HEADER_PAGE)),
    budget,
  );
  const errorCode = childText(header, "ErrorCode") === "true";
  // Per-file XPress8 inside the backup is a different compression the reader
  // does not attempt. It has no refusal code of its own, so it reads as an
  // unreadable model rather than a warned-about unverified path.
  if (childText(header, "ApplyCompression") === "true")
    throw unreadableModel("backup");
  const offsetHeader = positiveInteger(childText(header, "m_cbOffsetHeader"));
  const dataSize = positiveInteger(childText(header, "DataSize"));
  positiveInteger(childText(header, "Files"));
  if (offsetHeader > image.length - dataSize) throw unreadableModel("backup");

  // The decoded text is a live copy of the member, charged before it is made.
  budget?.reserve("backupHeaderText", dataSize * 2, "backup");
  const directory = parseXml(
    utf8.decode(image.subarray(offsetHeader, offsetHeader + dataSize)),
    budget,
  );
  const entries = childrenNamed(directory, "BackupFile").map((element) => ({
    path: childText(element, "Path"),
    size: positiveInteger(childText(element, "Size")),
    offset: positiveInteger(childText(element, "m_cbOffsetHeader")),
  }));
  const last = entries[entries.length - 1];
  if (last === undefined) throw unreadableModel("backup");
  if (last.offset > image.length - last.size) throw unreadableModel("backup");

  let logBytes = image.subarray(last.offset, last.offset + last.size);
  if (errorCode) {
    if (logBytes.length < 4) throw unreadableModel("backup");
    logBytes = logBytes.subarray(0, logBytes.length - 4);
  }
  budget?.reserve("backupHeaderText", logBytes.length, "backup");
  const log = parseXml(utf16Text(logBytes), budget);
  const groups = descend(log, "FileGroups/FileGroup");
  // The reference implementation indexes the second group unconditionally. A
  // backup with fewer groups is an unreadable model, not a crash.
  if (groups.length < 2) throw unreadableModel("backup");
  const persistLocation = childText(groups[1]!, "PersistLocationPath");
  if (persistLocation === null) throw unreadableModel("backup");
  const persistRoot = `${persistLocation}\\`;

  const byStoragePath = new Map(
    entries
      .filter((entry) => entry.path !== null)
      .map((entry) => [entry.path!, entry]),
  );
  const members: AbfMember[] = [];
  for (const group of groups) {
    for (const file of descend(group, "FileList/BackupFile")) {
      const storagePath = childText(file, "StoragePath");
      if (storagePath === null) continue;
      const entry = byStoragePath.get(storagePath);
      if (entry === undefined) continue;
      const full = childText(file, "Path");
      if (full === null) continue;
      const relative = full.startsWith(persistRoot)
        ? full.slice(persistRoot.length)
        : full;
      members.push({
        path: relative,
        fileName: relative.split("\\").pop() ?? relative,
        storagePath,
        size: entry.size,
        offset: entry.offset,
      });
    }
  }
  return { members, errorCode };
}

/**
 * The bytes of one backup member. Every offset is checked against the image,
 * which the reference implementation does not do.
 */
export function memberBytes(
  image: Uint8Array,
  backup: AbfImage,
  fileName: string,
): Uint8Array {
  const member = backup.members.find((entry) => entry.fileName === fileName);
  if (member === undefined) throw unreadableModel("backup");
  const trailer = backup.errorCode ? 4 : 0;
  const end = member.offset + member.size - trailer;
  if (
    member.offset < 0 ||
    end < member.offset ||
    member.offset + member.size > image.length
  )
    throw unreadableModel("backup");
  return image.subarray(member.offset, end);
}

/** True when a member of that exact name exists. */
export function hasMember(backup: AbfImage, fileName: string): boolean {
  return backup.members.some((member) => member.fileName === fileName);
}
