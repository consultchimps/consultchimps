import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import { convertWorkbookToValues } from "../src/values-only.js";

describe("convertWorkbookToValues", () => {
  it("removes formulas and calculation metadata without changing cell formatting", async () => {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([
      ["Amount", "Tax", "Total"],
      [100, 5, { f: "A2+B2", t: "n", v: 105, z: "$#,##0.00" }],
    ]);
    worksheet["!cols"] = [{ wch: 18 }, { wch: 12 }, { wch: 22 }];
    worksheet["!rows"] = [{ hpt: 28 }, { hpt: 20 }];
    XLSX.utils.book_append_sheet(workbook, worksheet, "Summary");
    const sourceBytes = XLSX.write(workbook, {
      bookType: "xlsx",
      cellStyles: true,
      type: "buffer",
    });

    const sourceArchive = await JSZip.loadAsync(sourceBytes);
    const sourceSheetXml = await sourceArchive
      .file("xl/worksheets/sheet1.xml")!
      .async("text");
    const sourceCellOpeningTag = /<c\b[^>]*\br="C2"[^>]*>/u.exec(
      sourceSheetXml,
    )?.[0];
    const sourceColumns = /<cols>[\s\S]*?<\/cols>/u.exec(sourceSheetXml)?.[0];
    const sourceFirstRowOpeningTag = /<row\b[^>]*\br="1"[^>]*>/u.exec(
      sourceSheetXml,
    )?.[0];
    expect(sourceCellOpeningTag).toBeDefined();
    expect(sourceSheetXml).toContain("<f>A2+B2</f><v>105</v>");

    const outputBytes = await convertWorkbookToValues(sourceBytes);
    const outputArchive = await JSZip.loadAsync(outputBytes);
    const outputSheetXml = await outputArchive
      .file("xl/worksheets/sheet1.xml")!
      .async("text");
    const outputCellOpeningTag = /<c\b[^>]*\br="C2"[^>]*>/u.exec(
      outputSheetXml,
    )?.[0];
    const outputColumns = /<cols>[\s\S]*?<\/cols>/u.exec(outputSheetXml)?.[0];
    const outputFirstRowOpeningTag = /<row\b[^>]*\br="1"[^>]*>/u.exec(
      outputSheetXml,
    )?.[0];

    expect(outputSheetXml).not.toContain("<f>");
    expect(outputSheetXml).toContain("<v>105</v>");
    expect(outputCellOpeningTag).toBe(sourceCellOpeningTag);
    expect(outputColumns).toBe(sourceColumns);
    expect(outputFirstRowOpeningTag).toBe(sourceFirstRowOpeningTag);
  });

  it("removes table formulas and stale calculation-chain references", async () => {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([["Value"], [1]]),
      "Data",
    );
    const sourceArchive = await JSZip.loadAsync(
      XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
    );
    sourceArchive.file(
      "xl/tables/table1.xml",
      "<table><tableColumns><tableColumn><calculatedColumnFormula>[@Value]*2</calculatedColumnFormula><totalsRowFormula>SUM([Value])</totalsRowFormula></tableColumn></tableColumns></table>",
    );
    sourceArchive.file("xl/calcChain.xml", "<calcChain />");
    const relationships = await sourceArchive
      .file("xl/_rels/workbook.xml.rels")!
      .async("text");
    sourceArchive.file(
      "xl/_rels/workbook.xml.rels",
      relationships.replace(
        "</Relationships>",
        '<Relationship Id="rCalc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>',
      ),
    );
    const contentTypes = await sourceArchive
      .file("[Content_Types].xml")!
      .async("text");
    sourceArchive.file(
      "[Content_Types].xml",
      contentTypes.replace(
        "</Types>",
        '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
      ),
    );

    const outputArchive = await JSZip.loadAsync(
      await convertWorkbookToValues(
        await sourceArchive.generateAsync({ type: "nodebuffer" }),
      ),
    );
    const tableXml = await outputArchive
      .file("xl/tables/table1.xml")!
      .async("text");
    const outputRelationships = await outputArchive
      .file("xl/_rels/workbook.xml.rels")!
      .async("text");
    const outputContentTypes = await outputArchive
      .file("[Content_Types].xml")!
      .async("text");

    expect(tableXml).not.toContain("calculatedColumnFormula");
    expect(tableXml).not.toContain("totalsRowFormula");
    expect(outputArchive.file("xl/calcChain.xml")).toBeNull();
    expect(outputRelationships).not.toContain("calcChain");
    expect(outputContentTypes).not.toContain("calcChain");
  });
});
