import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import JSZip from "jszip";
import { afterEach, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

const execFileAsync = promisify(execFile);
const cliPath = fileURLToPath(new URL("../dist/index.js", import.meta.url));
const temporaryDirectories: string[] = [];

async function protectedWorkbook(): Promise<Uint8Array> {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([["Value"], [42]]),
    "Summary",
  );
  const bytes = XLSX.write(workbook, {
    bookType: "xlsx",
    compression: true,
    type: "array",
  }) as ArrayBuffer;
  const archive = await JSZip.loadAsync(bytes);
  const workbookXml = await archive.file("xl/workbook.xml")!.async("text");
  archive.file(
    "xl/workbook.xml",
    workbookXml.replace(
      "</workbook>",
      '<workbookProtection lockStructure="1"/></workbook>',
    ),
  );
  const sheetXml = await archive
    .file("xl/worksheets/sheet1.xml")!
    .async("text");
  archive.file(
    "xl/worksheets/sheet1.xml",
    sheetXml.replace(
      "</worksheet>",
      '<sheetProtection sheet="1"/></worksheet>',
    ),
  );
  return archive.generateAsync({ compression: "DEFLATE", type: "uint8array" });
}

async function runCli(args: string[]): Promise<{
  stderr: string;
  stdout: string;
}> {
  const result = await execFileAsync(process.execPath, [cliPath, ...args], {
    encoding: "utf8",
  });
  return { stderr: result.stderr, stdout: result.stdout };
}

describe("built Excel unprotect command", () => {
  afterEach(async () => {
    await Promise.all(
      temporaryDirectories
        .splice(0)
        .map((directory) => rm(directory, { force: true, recursive: true })),
    );
  });

  it("writes a new workbook and reports removed protections as JSON", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-cli-"));
    temporaryDirectories.push(directory);
    const input = path.join(directory, "protected.xlsx");
    const output = path.join(directory, "unprotected.xlsx");
    const source = await protectedWorkbook();
    await writeFile(input, source);

    const result = await runCli([
      "--json",
      "sheets",
      "unprotect",
      input,
      "--output",
      output,
    ]);

    expect(result.stderr).toBe("");
    expect(JSON.parse(result.stdout)).toMatchObject({
      ok: true,
      result: {
        metrics: {
          sheetProtectionsRemoved: 1,
          workbookProtectionsRemoved: 1,
        },
      },
    });
    expect(new Uint8Array(await readFile(input))).toEqual(source);
    const outputArchive = await JSZip.loadAsync(await readFile(output));
    expect(
      await outputArchive.file("xl/workbook.xml")!.async("text"),
    ).not.toContain("workbookProtection");
    expect(
      await outputArchive.file("xl/worksheets/sheet1.xml")!.async("text"),
    ).not.toContain("sheetProtection");
  });

  it("returns a stable error for a non-OOXML input", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-cli-"));
    temporaryDirectories.push(directory);
    const input = path.join(directory, "encrypted.xlsx");
    await writeFile(input, Buffer.from("not an Office package"));

    await expect(
      runCli([
        "--json",
        "sheets",
        "unprotect",
        input,
        "-o",
        path.join(directory, "out.xlsx"),
      ]),
    ).rejects.toMatchObject({
      code: 1,
      stdout: expect.stringContaining("XLSX_UNPROTECT_UNSUPPORTED_FILE"),
    });
  });
});
