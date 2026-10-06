import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { consolidateWorkbooks } from "@consultchimps/xlsx";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import type { NamedFile } from "./operation-tasks";
import {
  OUTPUT_DIRECTORY,
  OUTPUT_RETENTION_MS,
  removeOutputs,
  type OutputDirectoryHandle,
  type OutputFileHandle,
  type OutputLocks,
  type OutputStorage,
} from "./output-storage";
import { consolidateFiles } from "./streamed-consolidation";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** A one-sheet workbook of `rows` rows, stored uncompressed so it is large. */
async function workbook(prefix: string, rows: number): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  zip.file(
    "xl/workbook.xml",
    `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Log" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  );
  const cell = (ref: string, text: string) =>
    `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;
  const body = [
    `<row r="1">${cell("A1", "Case_ID")}${cell("B1", "Note")}</row>`,
  ];
  for (let row = 2; row <= rows; row += 1) {
    body.push(
      `<row r="${row}">${cell(`A${row}`, `${prefix}-${row}`)}${cell(`B${row}`, `A longer note for row ${row}`)}</row>`,
    );
  }
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<worksheet xmlns="${MAIN}"><sheetData>${body.join("")}</sheetData></worksheet>`,
  );
  return zip.generateAsync({ type: "uint8array", compression: "STORE" });
}

/** A file whose reads fail when one asks for more than 1 MiB at once. */
function guardedFile(name: string, bytes: Uint8Array): NamedFile {
  const blob = new Blob([bytes.slice()]);
  const guarded = {
    size: blob.size,
    slice: (start: number, end: number) => {
      if (end - start > 1024 * 1024) {
        throw new Error(`Read ${end - start} of ${blob.size} bytes at once`);
      }
      return blob.slice(start, end);
    },
    arrayBuffer: () => Promise.reject(new Error("Read the whole file")),
  };
  return { name, file: guarded as unknown as Blob };
}

interface FakeFile {
  data: Uint8Array;
  open: boolean;
}

/**
 * A stand-in for OPFS whose sync access handle writes at most 1,000 bytes per
 * call, as the standard allows, unless `stall` makes it write nothing.
 */
function fakeStorage(
  options: { readonly stall?: boolean; readonly onWrite?: () => void } = {},
) {
  const files = new Map<string, FakeFile>();
  const fileHandle = (name: string): OutputFileHandle => ({
    kind: "file",
    name,
    getFile: () =>
      Promise.resolve(new File([files.get(name)!.data.slice()], name)),
    createSyncAccessHandle: () => {
      const file = files.get(name)!;
      file.open = true;
      return Promise.resolve({
        write: (buffer, { at }) => {
          options.onWrite?.();
          if (options.stall) return 0;
          const count = Math.min(buffer.length, 1000);
          const end = at + count;
          if (end > file.data.length) {
            const grown = new Uint8Array(end);
            grown.set(file.data);
            file.data = grown;
          }
          file.data.set(buffer.subarray(0, count), at);
          return count;
        },
        truncate: (size) => {
          file.data = file.data.slice(0, size);
        },
        flush: () => undefined,
        close: () => {
          file.open = false;
        },
      });
    },
  });
  const directory: OutputDirectoryHandle = {
    kind: "directory",
    name: OUTPUT_DIRECTORY,
    async *values() {
      for (const name of files.keys()) yield fileHandle(name);
    },
    getDirectoryHandle: () => Promise.resolve(directory),
    getFileHandle: (name, { create } = {}) => {
      if (!files.has(name)) {
        if (!create) return Promise.reject(new Error("NotFoundError"));
        files.set(name, { data: new Uint8Array(0), open: false });
      }
      return Promise.resolve(fileHandle(name));
    },
    removeEntry: (name) => {
      files.delete(name);
      return Promise.resolve();
    },
  };
  const storage: OutputStorage = {
    getDirectory: () =>
      Promise.resolve({
        ...directory,
        name: "",
        getDirectoryHandle: () => Promise.resolve(directory),
      }),
  };
  return { storage, files, place: { storage, locks: fakeLocks() } };
}

/** Web Locks for one process: a held name refuses an `ifAvailable` request. */
function fakeLocks(
  held = new Set<string>(),
): OutputLocks & { held: Set<string> } {
  return {
    held,
    request: async (name, options, callback) => {
      if (held.has(name)) {
        if (options.ifAvailable === true) return callback(null);
        throw new Error(`${name} is held`);
      }
      held.add(name);
      try {
        return await callback({ name });
      } finally {
        held.delete(name);
      }
    },
  };
}

const controls = () => ({
  onProgress: () => undefined,
  signal: new AbortController().signal,
});

async function commandLineBytes(
  inputs: ReadonlyArray<readonly [string, Uint8Array]>,
): Promise<Buffer> {
  const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-docs-"));
  try {
    const paths = [];
    for (const [name, bytes] of inputs) {
      const target = path.join(directory, name);
      await writeFile(target, bytes);
      paths.push(target);
    }
    const output = path.join(directory, "out.xlsx");
    await consolidateWorkbooks({ inputs: paths, output });
    return await readFile(output);
  } finally {
    await rm(directory, { force: true, recursive: true });
  }
}

describe("consolidateFiles", () => {
  it("reads in pieces, writes to OPFS, and gives the command line's bytes", async () => {
    const inputs = [
      ["north.xlsx", await workbook("N", 12_000)],
      ["south.xlsx", await workbook("S", 50)],
    ] as const;
    expect(inputs[0][1].length).toBeGreaterThan(1024 * 1024);
    const { place, files } = fakeStorage();
    const created = new Set<string>();

    const { outputs, result } = await consolidateFiles(
      inputs.map(([name, bytes]) => guardedFile(name, bytes)),
      { outputName: "out" },
      controls(),
      place,
      created,
    );

    expect(result.metrics.outputRows).toBe(12_048);
    expect(outputs).toHaveLength(1);
    expect(outputs[0]!.name).toBe("out.xlsx");
    expect(outputs[0]!.inMemory).toBeUndefined();
    expect(created.size).toBe(1);
    expect([...files.values()].every((file) => !file.open)).toBe(true);
    expect(Buffer.from(await outputs[0]!.blob.arrayBuffer())).toEqual(
      await commandLineBytes(inputs),
    );
  }, 60_000);

  it("collects the output in memory where OPFS refuses", async () => {
    const bytes = await workbook("N", 20);
    const storage: OutputStorage = {
      getDirectory: () => Promise.reject(new Error("SecurityError")),
    };
    const { outputs } = await consolidateFiles(
      [guardedFile("north.xlsx", bytes)],
      { outputName: "out" },
      controls(),
      { storage, locks: undefined },
      new Set(),
    );
    expect(outputs[0]!.inMemory).toBe(true);
    expect(Buffer.from(await outputs[0]!.blob.arrayBuffer())).toEqual(
      await commandLineBytes([["north.xlsx", bytes]]),
    );
  });

  it("removes the OPFS file when the run is cancelled midway", async () => {
    const controller = new AbortController();
    const { place, files } = fakeStorage({
      onWrite: () => controller.abort(),
    });
    const created = new Set<string>();
    await expect(
      consolidateFiles(
        [guardedFile("north.xlsx", await workbook("N", 3_000))],
        {},
        { onProgress: () => undefined, signal: controller.signal },
        place,
        created,
      ),
    ).rejects.toMatchObject({ code: "OPERATION_ABORTED" });
    expect(files.size).toBe(0);
    expect(created.size).toBe(0);
  });

  it("fails rather than spinning when OPFS writes nothing", async () => {
    const { place, files } = fakeStorage({ stall: true });
    await expect(
      consolidateFiles(
        [guardedFile("north.xlsx", await workbook("N", 20))],
        {},
        controls(),
        place,
        new Set(),
      ),
    ).rejects.toMatchObject({ code: "OUTPUT_STORAGE_FULL" });
    expect(files.size).toBe(0);
  });
});

describe("unreadable inputs", () => {
  it("fails as an unreadable file, not a damaged workbook", async () => {
    const bytes = await workbook("N", 20);
    const blob = new Blob([bytes.slice()]);
    const gone = {
      size: blob.size,
      slice: () => ({
        arrayBuffer: () => Promise.reject(new Error("NotReadableError")),
      }),
    } as unknown as Blob;
    const { place } = fakeStorage();
    await expect(
      consolidateFiles(
        [{ name: "north.xlsx", file: gone }],
        {},
        controls(),
        place,
        new Set(),
      ),
    ).rejects.toMatchObject({
      code: "FILE_UNREADABLE",
      details: { source: "north.xlsx" },
    });
  });
});

describe("removeOutputs", () => {
  it("removes this worker's outputs and unheld ones, and keeps another tab's", async () => {
    const { storage, files } = fakeStorage();
    const now = 10 * OUTPUT_RETENTION_MS;
    for (const name of [
      "1-mine.part",
      "2-other-tab.part",
      `${now}-closed-tab.part`,
    ]) {
      files.set(name, { data: new Uint8Array(1), open: false });
    }
    // Another tab holds its output's lock however old the file is.
    const locks = fakeLocks(new Set(["consultchimps-output:2-other-tab.part"]));
    const created = new Set(["1-mine.part"]);
    await removeOutputs({ storage, locks }, created, now);
    expect([...files.keys()]).toEqual(["2-other-tab.part"]);
    expect(created.size).toBe(0);
  });

  it("falls back to age where the browser has no locks", async () => {
    const { storage, files } = fakeStorage();
    const now = 10 * OUTPUT_RETENTION_MS;
    for (const name of [
      `${now - 2000}-recent.part`,
      `${now - OUTPUT_RETENTION_MS - 1}-abandoned.part`,
    ]) {
      files.set(name, { data: new Uint8Array(1), open: false });
    }
    await removeOutputs({ storage, locks: undefined }, new Set(), now);
    expect([...files.keys()]).toEqual([`${now - 2000}-recent.part`]);
  });
});
