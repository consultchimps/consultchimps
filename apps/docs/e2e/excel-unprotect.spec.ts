import { expect, test } from "@playwright/test";
import JSZip from "jszip";
import { readFile } from "node:fs/promises";
import {
  createWorkbookUpload,
  fileInput,
  resultArtifacts,
  resultsPanel,
  type UploadFile,
} from "./fixtures";

async function protectedWorkbookUpload(): Promise<UploadFile> {
  const source = await createWorkbookUpload(
    "protected.xlsm",
    [
      { name: "Summary", rows: [["Value"], [42]] },
      { name: "Hidden", rows: [["Private"], [99]], state: "hidden" },
    ],
    { macroEnabled: true },
  );
  const archive = await JSZip.loadAsync(source.buffer);
  const workbookXml = await archive.file("xl/workbook.xml")!.async("text");
  archive.file(
    "xl/workbook.xml",
    workbookXml.replace(
      "</workbook>",
      '<workbookProtection lockStructure="1"/></workbook>',
    ),
  );
  for (const sheetPath of [
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/sheet2.xml",
  ]) {
    const sheetXml = await archive.file(sheetPath)!.async("text");
    archive.file(
      sheetPath,
      sheetXml.replace(
        "</worksheet>",
        '<sheetProtection sheet="1"/></worksheet>',
      ),
    );
  }
  return {
    ...source,
    buffer: await archive.generateAsync({ type: "nodebuffer" }),
  };
}

test.describe("/tools/excel-unprotect", () => {
  test("runs in the operation worker and downloads the unprotected workbook", async ({
    page,
  }) => {
    await page.goto("/tools/excel-unprotect");
    await expect(
      page.getByRole("heading", {
        level: 1,
        name: "Unprotect an Excel workbook",
      }),
    ).toBeVisible();

    await fileInput(page).setInputFiles(await protectedWorkbookUpload());
    await expect(page.getByTestId("chosen-file")).toContainText(
      "protected.xlsm",
    );
    await page.getByTestId("run-button").click();

    await expect(resultsPanel(page)).toBeVisible();
    await expect(resultArtifacts(page)).toHaveCount(1);
    await expect(resultArtifacts(page).first()).toContainText(
      "protected-unprotected.xlsm",
    );
    await expect(page.getByTestId("result-message")).toContainText("finished");
    const downloadPromise = page.waitForEvent("download");
    await resultArtifacts(page).getByTestId("artifact-download").click();
    const download = await downloadPromise;
    expect(download.suggestedFilename()).toBe("protected-unprotected.xlsm");
    const outputArchive = await JSZip.loadAsync(
      await readFile(await download.path()),
    );
    expect(
      await outputArchive.file("xl/workbook.xml")!.async("text"),
    ).not.toContain("workbookProtection");
    expect(
      await outputArchive.file("xl/worksheets/sheet1.xml")!.async("text"),
    ).not.toContain("sheetProtection");
    expect(
      await outputArchive.file("xl/worksheets/sheet2.xml")!.async("text"),
    ).not.toContain("sheetProtection");
    expect(
      await outputArchive.file("xl/vbaProject.bin")!.async("uint8array"),
    ).toEqual(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]));
  });
});
