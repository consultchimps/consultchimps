/**
 * The command line merges through a bounded number of open inputs, opening an
 * input again when its rows are written and refusing one that changed.
 */
import { mkdtemp, readFile, rm, utimes, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import JSZip from "jszip";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { mergeWorkbooksBytes } from "../src/bytes.js";
import { mergeWorkbooks } from "../src/index.js";
import { MergeInputs, OPEN_INPUT_LIMIT } from "../src/merge/input-handles.js";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

async function workbook(label: string): Promise<Uint8Array> {
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
    `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${label}</t></is></c></row></sheetData></worksheet>`,
  );
  return zip.generateAsync({ type: "uint8array" });
}

let folder: string;
beforeEach(async () => {
  folder = await mkdtemp(path.join(tmpdir(), "merge-inputs-"));
});
afterEach(async () => {
  await rm(folder, { recursive: true, force: true });
});

describe("merging through a bounded number of open inputs", () => {
  it("merges more inputs than it keeps open, as the byte surface does", async () => {
    const count = OPEN_INPUT_LIMIT + 4;
    const inputs: Array<{ name: string; bytes: Uint8Array }> = [];
    for (let index = 0; index < count; index += 1) {
      const name = `input-${String(index).padStart(2, "0")}.xlsx`;
      const bytes = await workbook(`Input ${index}`);
      await writeFile(path.join(folder, name), bytes);
      inputs.push({ name, bytes });
    }
    const output = path.join(folder, "merged.xlsx");
    await mergeWorkbooks(
      inputs.map((input) => path.join(folder, input.name)),
      output,
    );
    const reference = await mergeWorkbooksBytes({ inputs });
    expect(
      Buffer.from(await readFile(output)).equals(
        Buffer.from(reference.outputs[0]!.bytes),
      ),
    ).toBe(true);
  }, 60_000);

  it("reads an input again after closing it, and refuses it once changed", async () => {
    const first = path.join(folder, "first.xlsx");
    const second = path.join(folder, "second.xlsx");
    await writeFile(first, await workbook("First"));
    await writeFile(second, await workbook("Second"));
    const inputs = new MergeInputs(1);
    try {
      const a = await inputs.open(first);
      const b = await inputs.open(second);
      const head = await a.readAt(0, 4);
      expect([...head]).toEqual([0x50, 0x4b, 0x03, 0x04]);
      expect([...(await b.readAt(0, 4))]).toEqual([...head]);

      const later = new Date(Date.now() + 60_000);
      await utimes(first, later, later);
      await expect(a.readAt(0, 4)).rejects.toMatchObject({
        code: "FILES_SOURCE_CHANGED",
      });
      await expect(a.verifyUnchanged()).rejects.toMatchObject({
        code: "FILES_SOURCE_CHANGED",
      });
      await expect(b.verifyUnchanged()).resolves.toBeUndefined();
    } finally {
      await inputs.close();
    }
  });
});
