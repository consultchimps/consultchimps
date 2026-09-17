import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { ConsultChimpsError } from "@consultchimps/core";
import { hasMember, memberBytes, parseAbf, parseXml } from "../src/abf.js";
import { PipelineBudget, validateExportOptions } from "../src/budget.js";
import { readPbiModelPart } from "../src/container.js";
import { decompressModelPart } from "../src/xpress9/stream.js";
import { FIXTURES } from "./oracle.js";

/** Section C, "Backup container (ABF)". */

const utf16 = (text: string): Uint8Array => {
  const bytes = new Uint8Array(text.length * 2);
  for (let index = 0; index < text.length; index++) {
    const unit = text.charCodeAt(index);
    bytes[index * 2] = unit & 0xff;
    bytes[index * 2 + 1] = unit >> 8;
  }
  return bytes;
};
const utf8 = (text: string): Uint8Array => new TextEncoder().encode(text);

interface Member {
  readonly path: string;
  readonly storagePath: string;
  readonly bytes: Uint8Array;
  readonly group: number;
}

interface BuildOptions {
  readonly members?: readonly Member[];
  readonly groups?: number;
  readonly errorCode?: boolean;
  readonly applyCompression?: boolean;
  readonly omitHeaderField?: string;
  readonly directoryOutside?: boolean;
  readonly memberOutside?: boolean;
}

/**
 * The smallest ABF image the reader accepts, so each refusal row can be
 * triggered by changing exactly one thing about a valid image.
 */
function buildAbf(options: BuildOptions = {}): Uint8Array {
  const members =
    options.members ??
    ([
      {
        path: "metadata.sqlitedb",
        storagePath: "AAAA",
        bytes: utf8("catalog"),
        group: 1,
      },
    ] as const);
  const groupCount = options.groups ?? 2;
  const root = "C:\\backup";
  const blocks: { storagePath: string; bytes: Uint8Array; offset: number }[] =
    [];
  let cursor = 0x1000;
  const trailer = (options.errorCode ?? true) ? 4 : 0;
  for (const member of members) {
    blocks.push({
      storagePath: member.storagePath,
      bytes: member.bytes,
      offset: cursor,
    });
    cursor += member.bytes.length + trailer;
  }
  const groups: string[] = [];
  for (let index = 0; index < groupCount; index++) {
    const list = members
      .filter((member) => member.group === index)
      .map(
        (member) =>
          `<BackupFile><Path>${root}\\${member.path}</Path><StoragePath>${member.storagePath}</StoragePath><Size>${member.bytes.length}</Size></BackupFile>`,
      )
      .join("");
    groups.push(
      `<FileGroup><PersistLocationPath>${root}</PersistLocationPath><FileList>${list}</FileList></FileGroup>`,
    );
  }
  const log = utf16(
    `<BackupLog><FileGroups>${groups.join("")}</FileGroups></BackupLog>`,
  );
  const logOffset = cursor;
  cursor += log.length + trailer;

  const entries = [
    ...blocks.map(
      (block) =>
        `<BackupFile><Path>${block.storagePath}</Path><Size>${block.bytes.length + trailer}</Size><m_cbOffsetHeader>${options.memberOutside ? 1 << 30 : block.offset}</m_cbOffsetHeader></BackupFile>`,
    ),
    `<BackupFile><Path>LOG</Path><Size>${log.length + trailer}</Size><m_cbOffsetHeader>${logOffset}</m_cbOffsetHeader></BackupFile>`,
  ].join("");
  const directory = utf8(`<VirtualDirectory>${entries}</VirtualDirectory>`);
  const directoryOffset = options.directoryOutside ? 1 << 30 : cursor;
  cursor += directory.length;

  const fields: Record<string, string> = {
    ErrorCode: String(options.errorCode ?? true),
    ApplyCompression: String(options.applyCompression ?? false),
    m_cbOffsetHeader: String(directoryOffset),
    DataSize: String(directory.length),
    Files: String(members.length + 1),
  };
  const header = utf16(
    `<BackupHeader>${Object.entries(fields)
      .filter(([name]) => name !== options.omitHeaderField)
      .map(([name, value]) => `<${name}>${value}</${name}>`)
      .join("")}</BackupHeader>`,
  );

  const image = new Uint8Array(cursor);
  image.set(header, 72);
  for (const block of blocks) image.set(block.bytes, block.offset);
  image.set(log, logOffset);
  if (!options.directoryOutside) image.set(directory, directoryOffset);
  return image;
}

function refusal(run: () => unknown): ConsultChimpsError {
  try {
    run();
  } catch (error) {
    if (error instanceof ConsultChimpsError) return error;
    throw error;
  }
  throw new Error("expected a refusal");
}

describe("the XML tokenizer", () => {
  it("reads a flat tree, skipping declarations and comments", () => {
    const tree = parseXml('<?xml version="1.0"?><!-- c --><a><b>1</b><c/></a>');
    expect(tree.tag).toBe("a");
    expect(tree.children.map((child) => child.tag)).toEqual(["b", "c"]);
  });

  it("refuses a header that nests past the depth ceiling", () => {
    const deep = "<a>".repeat(200) + "</a>".repeat(200);
    expect(refusal(() => parseXml(deep)).code).toBe("PBI_MODEL_UNREADABLE");
  });

  it("refuses a header with more nodes than the ceiling", () => {
    const wide = `<a>${"<b/>".repeat(200_001)}</a>`;
    expect(refusal(() => parseXml(wide)).code).toBe("PBI_MODEL_UNREADABLE");
  });
});

describe("the backup container", () => {
  it("reads a member and trims the four-byte trailer", () => {
    const image = buildAbf();
    const backup = parseAbf(image);
    expect(backup.errorCode).toBe(true);
    expect(hasMember(backup, "metadata.sqlitedb")).toBe(true);
    expect(
      new TextDecoder().decode(memberBytes(image, backup, "metadata.sqlitedb")),
    ).toBe("catalog");
  });

  it("refuses a header missing a required field", () => {
    for (const field of ["m_cbOffsetHeader", "DataSize", "Files"])
      expect(
        refusal(() => parseAbf(buildAbf({ omitHeaderField: field }))).code,
      ).toBe("PBI_MODEL_UNREADABLE");
  });

  it("refuses a virtual directory outside the image", () => {
    expect(
      refusal(() => parseAbf(buildAbf({ directoryOutside: true }))).code,
    ).toBe("PBI_MODEL_UNREADABLE");
  });

  it("refuses per-file XPress8 compression rather than attempting it", () => {
    expect(
      refusal(() => parseAbf(buildAbf({ applyCompression: true }))).code,
    ).toBe("PBI_MODEL_UNREADABLE");
  });

  it("refuses fewer than two file groups instead of throwing", () => {
    // The reference implementation indexes the second group unconditionally.
    expect(refusal(() => parseAbf(buildAbf({ groups: 1 }))).code).toBe(
      "PBI_MODEL_UNREADABLE",
    );
  });

  it("refuses a missing catalog member", () => {
    const image = buildAbf({
      members: [
        { path: "other.bin", storagePath: "AAAA", bytes: utf8("x"), group: 1 },
      ],
    });
    const backup = parseAbf(image);
    expect(hasMember(backup, "metadata.sqlitedb")).toBe(false);
    expect(
      refusal(() => memberBytes(image, backup, "metadata.sqlitedb")).code,
    ).toBe("PBI_MODEL_UNREADABLE");
  });

  it("refuses a member whose offset and size fall outside the image", () => {
    const image = buildAbf({ memberOutside: true });
    const backup = parseAbf(image);
    expect(
      refusal(() => memberBytes(image, backup, "metadata.sqlitedb")).code,
    ).toBe("PBI_MODEL_UNREADABLE");
  });
});

describe("the real corpus backup", () => {
  it("lists exactly the members the independent reader lists", async () => {
    const container = new Uint8Array(
      readFileSync(path.join(FIXTURES, "a-2018-fuzzy.pbix")),
    );
    const budget = new PipelineBudget(validateExportOptions({}, true));
    const stream = await decompressModelPart(
      readPbiModelPart(container),
      budget,
      undefined,
    );
    const backup = parseAbf(stream.bytes);
    const expected = JSON.parse(
      readFileSync(
        path.join(FIXTURES, "oracle", "a-2018-fuzzy", "file-log.json"),
        "utf8",
      ),
    ) as { FileName: string; Size: number; m_cbOffsetHeader: number }[];
    expect(backup.members).toHaveLength(expected.length);
    expect(
      backup.members.map((member) => [
        member.fileName,
        member.size,
        member.offset,
      ]),
    ).toEqual(
      expected.map((entry) => [
        entry.FileName,
        entry.Size,
        entry.m_cbOffsetHeader,
      ]),
    );
  }, 120_000);
});
