import { expect, test } from "@playwright/test";
import JSZip from "jszip";

import { createWorkbookUpload, type UploadFile } from "./fixtures";

async function malformedStylesWorkbook(): Promise<UploadFile> {
  const workbook = await createWorkbookUpload("malformed-styles.xlsx", [
    { name: "Inventory", rows: [["Region"], ["North"]] },
  ]);
  const archive = await JSZip.loadAsync(workbook.buffer);
  archive.file(
    "xl/styles.xml",
    "<styleSheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><cellXfs count='1'><xf numFmtId='0'></cellXfs></styleSheet>",
  );
  return {
    ...workbook,
    buffer: await archive.generateAsync({ type: "nodebuffer" }),
  };
}

test("releases a workbook after lazy stream initialization fails", async ({
  page,
}) => {
  const malformed = await malformedStylesWorkbook();
  const valid = await createWorkbookUpload("valid.xlsx", [
    { name: "Inventory", rows: [["Region"], ["South"]] },
  ]);
  await page.goto("/tools/db");
  await page
    .getByTestId("workspace-new-name")
    .fill(`lazy-stream-${crypto.randomUUID()}.sqlite`);
  await page.getByTestId("workspace-new").click();

  const input = page.getByTestId("workspace-import-input");
  await input.setInputFiles(malformed);
  await page.getByTestId("workspace-import-prepare").click();
  await expect(page.getByTestId("workspace-error")).toContainText(
    'Could not read workbook "malformed-styles.xlsx"',
  );
  await expect(page.getByTestId("workspace-error")).not.toContainText(
    "cleanup",
  );
  await expect(page.getByTestId("workspace-import-review")).toHaveCount(0);
  await expect
    .poll(() =>
      page.evaluate(async () => {
        const root = await navigator.storage.getDirectory();
        const scratch: string[] = [];
        for await (const name of root.keys()) {
          if (name.startsWith(".consultchimps-scratch-")) scratch.push(name);
        }
        return scratch;
      }),
    )
    .toEqual([]);

  await input.setInputFiles(valid);
  await page.getByTestId("workspace-import-prepare").click();
  await expect(page.getByTestId("workspace-error")).toHaveCount(0);
  await expect(page.getByTestId("workspace-import-review")).toBeVisible();
  await expect(page.getByTestId("workspace-import-review")).toContainText(
    "valid.xlsx: Inventory",
  );
});
