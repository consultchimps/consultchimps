import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import {
  convertWorkbookToValues,
  convertWorkbookToValuesWithReport,
} from "../src/values-only.js";
import {
  buildSheetFixture,
  buildWorkbookFixture,
} from "./support/workbook-fixture.js";

describe("convertWorkbookToValues", () => {
  it("removes formulas and calculation metadata without changing cell formatting", async () => {
    // Column widths and row heights are added to the worksheet by hand.
    const built = await JSZip.loadAsync(
      await buildWorkbookFixture({
        sheets: [
          {
            name: "Summary",
            rows: [
              ["Amount", "Tax", "Total"],
              [100, 5, { formula: "A2+B2", value: 105, format: "$#,##0.00" }],
            ],
          },
        ],
      }),
    );
    const builtSheet = await built
      .file("xl/worksheets/sheet1.xml")!
      .async("text");
    built.file(
      "xl/worksheets/sheet1.xml",
      builtSheet
        .replace(
          "<sheetData>",
          '<cols><col min="1" max="1" width="18.7109375" customWidth="1"/><col min="2" max="2" width="12.7109375" customWidth="1"/><col min="3" max="3" width="22.7109375" customWidth="1"/></cols><sheetData>',
        )
        .replace('<row r="1">', '<row r="1" ht="28" customHeight="1">')
        .replace('<row r="2">', '<row r="2" ht="20" customHeight="1">'),
    );
    const sourceBytes = await built.generateAsync({ type: "uint8array" });

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
    const sourceArchive = await JSZip.loadAsync(
      await buildSheetFixture("Data", [["Value"], [1]]),
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

  it("reports formulas whose cached values are unavailable", async () => {
    const conversion = await convertWorkbookToValuesWithReport(
      await buildSheetFixture("Formulas", [
        ["Cached", "Missing"],
        [{ formula: "1+1", value: 2 }, { formula: "2+2" }],
      ]),
    );
    expect(conversion.formulasConverted).toBe(2);
    expect(conversion.formulasWithoutCachedValues).toEqual([
      { cell: "B2", worksheetPart: "xl/worksheets/sheet1.xml" },
    ]);
  });
});
