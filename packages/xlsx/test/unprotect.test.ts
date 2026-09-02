import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import {
  isConsultChimpsError,
  type OperationProgress,
} from "@consultchimps/core";
import { unprotectWorkbookBytes } from "../src/bytes.js";

function workbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new();
  const visible = XLSX.utils.aoa_to_sheet([
    ["Amount", "Formula"],
    [10, { f: "A2*2" }],
  ]);
  const hidden = XLSX.utils.aoa_to_sheet([["Confidential", 42]]);
  XLSX.utils.book_append_sheet(workbook, visible, "Summary");
  XLSX.utils.book_append_sheet(workbook, hidden, "Hidden");
  workbook.Workbook = {
    ...workbook.Workbook,
    Sheets: [
      { name: "Summary", Hidden: 0 },
      { name: "Hidden", Hidden: 1 },
    ],
  };
  return new Uint8Array(
    XLSX.write(workbook, {
      bookType: "xlsx",
      cellStyles: true,
      compression: true,
      type: "array",
    }) as ArrayBuffer,
  );
}

async function protectedWorkbook(): Promise<Uint8Array> {
  const archive = await JSZip.loadAsync(workbookBytes());
  const workbookXml = await archive.file("xl/workbook.xml")!.async("text");
  archive.file(
    "xl/workbook.xml",
    workbookXml.replace(
      "</workbook>",
      '<workbookProtection lockStructure="1" workbookPassword="ABCD"/></workbook>',
    ),
  );
  for (const sheetPath of [
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/sheet2.xml",
  ]) {
    const xml = await archive.file(sheetPath)!.async("text");
    archive.file(
      sheetPath,
      xml.replace(
        "</worksheet>",
        '<sheetProtection sheet="1" objects="1" scenarios="1"/></worksheet>',
      ),
    );
  }
  return archive.generateAsync({ compression: "DEFLATE", type: "uint8array" });
}

async function packageText(bytes: Uint8Array, part: string): Promise<string> {
  const archive = await JSZip.loadAsync(bytes);
  return archive.file(part)!.async("text");
}

describe("Excel workbook unprotection", () => {
  it("removes worksheet and workbook protection while preserving workbook parts", async () => {
    const input = await protectedWorkbook();
    const source = new Uint8Array(input);
    const events: OperationProgress[] = [];

    const outcome = await unprotectWorkbookBytes({
      input: { name: "protected.xlsx", bytes: input },
      onProgress: (progress) => events.push(progress),
      outputName: "unprotected.xlsx",
    });

    expect(input).toEqual(source);
    expect(outcome.result.metrics).toEqual({
      sheetProtectionsRemoved: 2,
      workbookProtectionsRemoved: 1,
    });
    expect(outcome.result.artifacts[0]?.path).toBe("unprotected.xlsx");
    expect(events).toEqual([
      expect.objectContaining({ stage: "writing-output", completed: 1 }),
    ]);

    const output = outcome.outputs[0]!.bytes;
    expect(await packageText(output, "xl/workbook.xml")).not.toContain(
      "workbookProtection",
    );
    expect(await packageText(output, "xl/worksheets/sheet1.xml")).not.toContain(
      "sheetProtection",
    );
    expect(await packageText(output, "xl/worksheets/sheet2.xml")).not.toContain(
      "sheetProtection",
    );
    expect(await packageText(output, "xl/styles.xml")).toBe(
      await packageText(input, "xl/styles.xml"),
    );
    expect(XLSX.read(output, { type: "array" }).SheetNames).toEqual([
      "Summary",
      "Hidden",
    ]);
  });

  it("reports an already-unprotected workbook without changing its contents", async () => {
    const input = workbookBytes();
    const outcome = await unprotectWorkbookBytes({
      input: { name: "plain.xlsx", bytes: input },
    });

    expect(outcome.result.metrics).toEqual({
      sheetProtectionsRemoved: 0,
      workbookProtectionsRemoved: 0,
    });
    expect(
      XLSX.read(outcome.outputs[0]!.bytes, { type: "array" }).SheetNames,
    ).toEqual(["Summary", "Hidden"]);
  });

  it("preserves a macro-enabled package and removes its protection", async () => {
    const archive = await JSZip.loadAsync(await protectedWorkbook());
    const vba = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3]);
    archive.file("xl/vbaProject.bin", vba);
    const contentTypes = await archive
      .file("[Content_Types].xml")!
      .async("text");
    archive.file(
      "[Content_Types].xml",
      contentTypes
        .replace(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
          "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        )
        .replace(
          "</Types>",
          '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
        ),
    );
    const input = await archive.generateAsync({
      compression: "DEFLATE",
      type: "uint8array",
    });

    const outcome = await unprotectWorkbookBytes({
      input: { name: "macros.xlsm", bytes: input },
    });
    const outputArchive = await JSZip.loadAsync(outcome.outputs[0]!.bytes);
    expect(
      await outputArchive.file("xl/vbaProject.bin")!.async("uint8array"),
    ).toEqual(vba);
    expect(
      await outputArchive.file("[Content_Types].xml")!.async("text"),
    ).toContain("macroEnabled.main+xml");
    expect(outcome.result.metrics).toEqual({
      sheetProtectionsRemoved: 2,
      workbookProtectionsRemoved: 1,
    });
  });

  it.each([
    ["notes.txt", new Uint8Array([1, 2, 3])],
    ["broken.xlsx", new Uint8Array([0x50, 0x4b, 0x03, 0x04, 9, 9])],
  ])("rejects unsupported input %s", async (name, bytes) => {
    await expect(
      unprotectWorkbookBytes({ input: { name, bytes } }),
    ).rejects.toSatisfy((error: unknown) => {
      return (
        isConsultChimpsError(error) &&
        error.code === "XLSX_UNPROTECT_UNSUPPORTED_FILE"
      );
    });
  });
});

async function macroProtectedWorkbook(): Promise<Uint8Array> {
  const archive = await JSZip.loadAsync(await protectedWorkbook());
  archive.file("xl/vbaProject.bin", new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]));
  const contentTypes = await archive.file("[Content_Types].xml")!.async("text");
  archive.file(
    "[Content_Types].xml",
    contentTypes
      .replace(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
      )
      .replace(
        "</Types>",
        '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
      ),
  );
  return archive.generateAsync({ compression: "DEFLATE", type: "uint8array" });
}

const ORDINARY_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const MACRO_MEDIA_TYPE = "application/vnd.ms-excel.sheet.macroEnabled.12";

describe("unprotect output media types", () => {
  it("derives the media type from the package, not the output name", async () => {
    // An ordinary package named .xlsm is still ordinary: unprotect adds no
    // VBA project, so the name must not make the bytes claim to be macro-enabled.
    const outcome = await unprotectWorkbookBytes({
      input: { name: "book.xlsx", bytes: await protectedWorkbook() },
      outputName: "book.xlsm",
    });
    expect(outcome.outputs[0]!.name).toBe("book.xlsm");
    expect(outcome.outputs[0]!.mediaType).toBe(ORDINARY_MEDIA_TYPE);
    expect(outcome.result.artifacts[0]!.mediaType).toBe(ORDINARY_MEDIA_TYPE);
  });

  it("reports a macro-enabled package with the macro media type", async () => {
    // A macro package keeps the macro media type even under an .xlsx name.
    const outcome = await unprotectWorkbookBytes({
      input: { name: "macros.xlsm", bytes: await macroProtectedWorkbook() },
      outputName: "macros.xlsx",
    });
    expect(outcome.outputs[0]!.mediaType).toBe(MACRO_MEDIA_TYPE);
    expect(outcome.result.artifacts[0]!.mediaType).toBe(MACRO_MEDIA_TYPE);
  });
});
