import { mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import JSZip from "jszip";
import { afterEach, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import {
  inspectPowerPointTemplate,
  populatePowerPointTemplate,
} from "../src/index.js";

const temporaryDirectories: string[] = [];

interface TemplateShape {
  id: number;
  name: string;
  runs: Array<{
    color?: string;
    properties?: string;
    text: string;
  }>;
}

function encodeXml(value: string): string {
  return value
    .replace(/&/gu, "&amp;")
    .replace(/</gu, "&lt;")
    .replace(/>/gu, "&gt;");
}

async function createTemporaryDirectory(): Promise<string> {
  const directory = await mkdtemp(
    path.join(tmpdir(), "consultchimps-pptx-test-"),
  );
  temporaryDirectories.push(directory);
  return directory;
}

function shapeXml(shape: TemplateShape): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${shape.id}" name="${shape.name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="60000"><a:off x="1000" y="2000"/><a:ext cx="3000" cy="4000"/></a:xfrm><a:solidFill><a:srgbClr val="F4F4F4"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr lIns="100" rIns="200"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>${shape.runs
    .map(
      (run) =>
        `<a:r><a:rPr lang="en-US"${run.properties ?? ""}>${
          run.color
            ? `<a:solidFill><a:srgbClr val="${run.color}"/></a:solidFill>`
            : ""
        }</a:rPr><a:t>${encodeXml(run.text)}</a:t></a:r>`,
    )
    .join("")}</a:p></p:txBody></p:sp>`;
}

function slideXml(shapes: TemplateShape[], unsupportedText?: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes
    .map(shapeXml)
    .join(
      "",
    )}<p:pic><p:nvPicPr><p:cNvPr id="90" name="Unrelated picture"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImage"/></p:blipFill></p:pic>${
    unsupportedText
      ? `<p:graphicFrame><a:graphic><a:graphicData><a:t>${encodeXml(unsupportedText)}</a:t></a:graphicData></a:graphicFrame>`
      : ""
  }</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

async function writeTemplate(
  filePath: string,
  slides: string[],
): Promise<void> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${slides
      .map(
        (_, index) =>
          `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
      )
      .join("")}</Types>`,
  );
  zip.file(
    "_rels/.rels",
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
  );
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst>${slides
      .map(
        (_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 1}"/>`,
      )
      .join(
        "",
      )}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${slides
      .map(
        (_, index) =>
          `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index + 1}.xml"/>`,
      )
      .join("")}</Relationships>`,
  );
  slides.forEach((slide, index) => {
    zip.file(`ppt/slides/slide${index + 1}.xml`, slide);
  });
  await writeFile(
    filePath,
    await zip.generateAsync({ compression: "DEFLATE", type: "nodebuffer" }),
  );
}

async function writeWorkbook(
  filePath: string,
  rows: Array<Array<boolean | Date | null | number | string>>,
  options: {
    currencyColumn?: number;
    dateColumn?: number;
    percentColumn?: number;
    worksheetName?: string;
  } = {},
): Promise<void> {
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  if (options.currencyColumn !== undefined) {
    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const address = XLSX.utils.encode_cell({
        c: options.currencyColumn,
        r: rowIndex,
      });
      if (worksheet[address]) {
        worksheet[address].z = '$0.0,,"M"';
      }
    }
  }
  if (options.dateColumn !== undefined) {
    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const address = XLSX.utils.encode_cell({
        c: options.dateColumn,
        r: rowIndex,
      });
      if (worksheet[address]) {
        worksheet[address].z = "yyyy-mm-dd";
      }
    }
  }
  if (options.percentColumn !== undefined) {
    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const address = XLSX.utils.encode_cell({
        c: options.percentColumn,
        r: rowIndex,
      });
      if (worksheet[address]) {
        worksheet[address].z = "0.0%";
      }
    }
  }
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    options.worksheetName ?? "Companies",
  );
  await writeFile(
    filePath,
    XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
  );
}

async function outputSlideXmls(outputPath: string): Promise<string[]> {
  const zip = await JSZip.loadAsync(await readFile(outputPath));
  const presentationXml = await zip
    .file("ppt/presentation.xml")!
    .async("string");
  const relationshipsXml = await zip
    .file("ppt/_rels/presentation.xml.rels")!
    .async("string");
  const relationshipIds = [
    ...presentationXml.matchAll(/\br:id="(?<id>[^"]+)"/gu),
  ].map((match) => match.groups?.id ?? "");
  const targets = relationshipIds.map((relationshipId) => {
    const relationship = [
      ...relationshipsXml.matchAll(/<Relationship\b[^>]*\/>/gu),
    ].find((match) => match[0].includes(`Id="${relationshipId}"`));
    return /Target="(?<target>[^"]+)"/u.exec(relationship?.[0] ?? "")?.groups
      ?.target;
  });
  return Promise.all(
    targets.map((target) => zip.file(`ppt/${target}`)!.async("string")),
  );
}

async function expectErrorCode(
  operation: Promise<unknown>,
  code: string,
): Promise<void> {
  await expect(operation).rejects.toMatchObject({ code });
}

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { force: true, recursive: true })),
  );
});

describe("PowerPoint template population", () => {
  it("creates slides in row order while preserving text-run formatting, geometry, and unrelated objects", async () => {
    const directory = await createTemporaryDirectory();
    const templatePath = path.join(directory, "template.pptx");
    const workbookPath = path.join(directory, "companies.xlsx");
    const outputPath = path.join(directory, "profiles.pptx");
    const repeatedOutputPath = path.join(directory, "profiles-repeat.pptx");
    const templateSlide = slideXml([
      {
        id: 2,
        name: "Client",
        runs: [
          {
            color: "C00000",
            properties: ' b="1" i="1" u="sng"',
            text: "{{client_name}}",
          },
        ],
      },
      {
        id: 3,
        name: "Summary",
        runs: [
          {
            text: "Client: {{client_name}}",
          },
        ],
      },
      {
        id: 4,
        name: "Details",
        runs: [
          {
            text: "{{client_name}} generated {{revenue}}. Active: {{active}}. Date: {{report_date}}.",
          },
        ],
      },
      {
        id: 5,
        name: "Revenue",
        runs: [{ text: "Revenue: {{ revenue }}" }],
      },
      {
        id: 7,
        name: "Growth",
        runs: [{ text: "Growth: {{growth_rate}}" }],
      },
      {
        id: 6,
        name: "Static",
        runs: [{ text: "Static text stays unchanged." }],
      },
    ]);
    await writeTemplate(templatePath, [
      slideXml([{ id: 2, name: "Other", runs: [{ text: "Other slide" }] }]),
      templateSlide,
    ]);
    await writeWorkbook(
      workbookPath,
      [
        ["client_name", "revenue", "growth_rate", "report_date", "active"],
        [
          "Company A",
          12_400_000,
          0.082,
          new Date("2024-01-02T00:00:00Z"),
          true,
        ],
        [null, null, null, null, null],
        ["Company B", null, -0.021, new Date("2024-03-04T00:00:00Z"), false],
      ],
      { currencyColumn: 1, dateColumn: 3, percentColumn: 2 },
    );
    const originalTemplate = await readFile(templatePath);
    const originalWorkbook = await readFile(workbookPath);

    const result = await populatePowerPointTemplate({
      outputPath,
      templatePath,
      templateSlide: 2,
      workbookPath,
      worksheet: "Companies",
    });

    expect(result).toEqual({
      operation: "pptx.populate",
      artifacts: [
        {
          kind: "file",
          mediaType:
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
          path: outputPath,
        },
      ],
      warnings: ["Skipped 1 empty worksheet row."],
      metrics: {
        generatedSlides: 2,
        inputRows: 2,
        outputFiles: 1,
        placeholderFields: 5,
        placeholderOccurrences: 8,
        replacements: 16,
        skippedRows: 1,
        warnings: 1,
      },
    });
    expect(await readFile(templatePath)).toEqual(originalTemplate);
    expect(await readFile(workbookPath)).toEqual(originalWorkbook);
    await populatePowerPointTemplate({
      outputPath: repeatedOutputPath,
      templatePath,
      templateSlide: 2,
      workbookPath,
      worksheet: "Companies",
    });
    expect(await readFile(repeatedOutputPath)).toEqual(
      await readFile(outputPath),
    );

    const slides = await outputSlideXmls(outputPath);
    expect(slides).toHaveLength(2);
    const outputZip = await JSZip.loadAsync(await readFile(outputPath));
    expect(outputZip.file("ppt/slides/slide1.xml")).toBeNull();
    expect(outputZip.file("ppt/slides/slide2.xml")).toBeNull();
    expect(slides[0]).toContain("Company A");
    expect(slides[1]).toContain("Company B");
    expect(slides[0]).toContain("Revenue: $12.4M");
    expect(slides[1]).toContain("Revenue: </a:t>");
    expect(slides[0]).toContain("Growth: 8.2%");
    expect(slides[1]).toContain("Growth: -2.1%");
    expect(slides[0]).toContain("2024-01-02");
    expect(slides[0]).toContain("Active: TRUE");
    expect(slides[1]).toContain("Active: FALSE");
    expect(slides[0]).toContain(
      '<a:xfrm rot="60000"><a:off x="1000" y="2000"/><a:ext cx="3000" cy="4000"/></a:xfrm>',
    );
    expect(slides[0]).toContain('b="1" i="1" u="sng"');
    expect(slides[0]).toContain('<a:srgbClr val="C00000"/>');
    expect(slides[0]).toContain("Static text stays unchanged.");
    expect(slides[0]).toContain('name="Unrelated picture"');
    expect(slides.join("")).not.toContain("Other slide");
    expect(slides.join("")).not.toContain("{{");
    expect(
      (await readdir(directory)).some((entry) =>
        entry.startsWith(".consultchimps-pptx-"),
      ),
    ).toBe(false);
  });

  it("inspects valid, malformed, split-run, and unsupported placeholders without writing output", async () => {
    const directory = await createTemporaryDirectory();
    const templatePath = path.join(directory, "inspect.pptx");
    await writeTemplate(templatePath, [
      slideXml(
        [
          {
            id: 2,
            name: "Valid",
            runs: [{ text: "{{client_name}} / {{ client_name }}" }],
          },
          {
            id: 3,
            name: "Split",
            runs: [{ text: "{{reve" }, { text: "nue}}" }],
          },
          {
            id: 4,
            name: "Malformed",
            runs: [{ text: "{{growth_rate}" }],
          },
        ],
        "{{table_value}}",
      ),
    ]);

    await expect(
      inspectPowerPointTemplate(templatePath, { templateSlide: 1 }),
    ).resolves.toEqual({
      malformedPlaceholderCount: 1,
      placeholderOccurrences: 3,
      placeholders: [
        { name: "client_name", occurrences: 2 },
        { name: "revenue", occurrences: 1 },
      ],
      slideNumber: 1,
      unsupportedPlacementPlaceholders: ["table_value"],
      unsupportedSplitRunPlaceholders: [],
    });
  });

  it("populates split-run placeholders and defaults to the first worksheet and slide", async () => {
    const directory = await createTemporaryDirectory();
    const templatePath = path.join(directory, "template.pptx");
    const workbookPath = path.join(directory, "data.xlsx");
    const outputPath = path.join(directory, "output.pptx");
    await writeTemplate(templatePath, [
      slideXml([
        {
          id: 2,
          name: "Favorite number",
          runs: [
            { color: "C00000", text: "Favorite: {{fav_" },
            { color: "0000FF", text: "number}}" },
          ],
        },
      ]),
      slideXml([{ id: 3, name: "Unused", runs: [{ text: "Unused" }] }]),
    ]);
    await writeWorkbook(workbookPath, [["fav_number"], [42]], {
      worksheetName: "Records",
    });

    await expect(
      inspectPowerPointTemplate(templatePath, {}),
    ).resolves.toMatchObject({
      placeholderOccurrences: 1,
      placeholders: [{ name: "fav_number", occurrences: 1 }],
      slideNumber: 1,
      unsupportedSplitRunPlaceholders: [],
    });
    const result = await populatePowerPointTemplate({
      outputPath,
      templatePath,
      workbookPath,
    });

    expect(result).toMatchObject({
      metrics: {
        generatedSlides: 1,
        placeholderOccurrences: 1,
        replacements: 1,
      },
    });
    const slides = await outputSlideXmls(outputPath);
    expect(slides).toHaveLength(1);
    expect(slides[0]).toContain("Favorite: 42");
    expect(slides[0]).toContain('<a:srgbClr val="C00000"/>');
    expect(slides[0]).toContain('<a:srgbClr val="0000FF"/>');
    expect(slides[0]).not.toContain("{{fav_");
    expect(slides[0]).not.toContain("number}}");
  });

  it.each([
    {
      code: "PPTX_NO_PLACEHOLDERS",
      slide: slideXml([
        { id: 2, name: "Static", runs: [{ text: "No fields here" }] },
      ]),
    },
    {
      code: "PPTX_MALFORMED_PLACEHOLDER",
      slide: slideXml([
        { id: 2, name: "Malformed", runs: [{ text: "{{client_name}" }] },
      ]),
    },
    {
      code: "PPTX_UNSUPPORTED_PLACEHOLDER_PLACEMENT",
      slide: slideXml([], "{{client_name}}"),
    },
  ])(
    "rejects invalid template text with $code before writing",
    async ({ code, slide }) => {
      const directory = await createTemporaryDirectory();
      const templatePath = path.join(directory, "template.pptx");
      const workbookPath = path.join(directory, "data.xlsx");
      const outputPath = path.join(directory, "output.pptx");
      await writeTemplate(templatePath, [slide]);
      await writeWorkbook(workbookPath, [["client_name"], ["Company A"]]);

      await expectErrorCode(
        populatePowerPointTemplate({
          outputPath,
          templatePath,
          templateSlide: 1,
          workbookPath,
          worksheet: "Companies",
        }),
        code,
      );
      await expect(readFile(outputPath)).rejects.toMatchObject({
        code: "ENOENT",
      });
    },
  );

  it("rejects workbook selection and column validation failures before writing", async () => {
    const directory = await createTemporaryDirectory();
    const templatePath = path.join(directory, "template.pptx");
    const workbookPath = path.join(directory, "data.xlsx");
    const outputPath = path.join(directory, "output.pptx");
    await writeTemplate(templatePath, [
      slideXml([
        {
          id: 2,
          name: "Fields",
          runs: [{ text: "{{client_name}} {{revenue}}" }],
        },
      ]),
    ]);
    await writeWorkbook(workbookPath, [
      ["client_name", "growth"],
      ["Company A", "8%"],
    ]);

    await expectErrorCode(
      populatePowerPointTemplate({
        outputPath,
        templatePath,
        templateSlide: 1,
        workbookPath,
        worksheet: "Missing",
      }),
      "XLSX_WORKSHEET_NOT_FOUND",
    );
    await expectErrorCode(
      populatePowerPointTemplate({
        headerRow: 20,
        outputPath,
        templatePath,
        templateSlide: 1,
        workbookPath,
        worksheet: "Companies",
      }),
      "XLSX_INVALID_HEADER_ROW",
    );
    await expectErrorCode(
      populatePowerPointTemplate({
        outputPath,
        templatePath,
        templateSlide: 1,
        workbookPath,
        worksheet: "Companies",
      }),
      "PPTX_MISSING_EXCEL_COLUMN",
    );
    await expect(readFile(outputPath)).rejects.toMatchObject({
      code: "ENOENT",
    });
  });

  it("rejects duplicate or empty headers and a worksheet with no populated records", async () => {
    const directory = await createTemporaryDirectory();
    const templatePath = path.join(directory, "template.pptx");
    const duplicateWorkbook = path.join(directory, "duplicate.xlsx");
    const emptyHeaderWorkbook = path.join(directory, "empty-header.xlsx");
    const emptyDataWorkbook = path.join(directory, "empty-data.xlsx");
    const outputPath = path.join(directory, "output.pptx");
    await writeTemplate(templatePath, [
      slideXml([{ id: 2, name: "Field", runs: [{ text: "{{client_name}}" }] }]),
    ]);
    await writeWorkbook(duplicateWorkbook, [
      ["client_name", "client_name"],
      ["A", "B"],
    ]);
    await writeWorkbook(emptyHeaderWorkbook, [
      ["client_name", null],
      ["A", "value"],
    ]);
    await writeWorkbook(emptyDataWorkbook, [["client_name"]]);

    const baseOptions = {
      outputPath,
      templatePath,
      templateSlide: 1,
      worksheet: "Companies",
    };
    await expectErrorCode(
      populatePowerPointTemplate({
        ...baseOptions,
        workbookPath: duplicateWorkbook,
      }),
      "XLSX_DUPLICATE_HEADER",
    );
    await expectErrorCode(
      populatePowerPointTemplate({
        ...baseOptions,
        workbookPath: emptyHeaderWorkbook,
      }),
      "XLSX_EMPTY_HEADER",
    );
    await expectErrorCode(
      populatePowerPointTemplate({
        ...baseOptions,
        workbookPath: emptyDataWorkbook,
      }),
      "PPTX_NO_DATA_ROWS",
    );
    await expect(readFile(outputPath)).rejects.toMatchObject({
      code: "ENOENT",
    });
  });

  it("validates input types, missing files, slide number, and input collisions", async () => {
    const directory = await createTemporaryDirectory();
    const templatePath = path.join(directory, "template.pptx");
    const workbookPath = path.join(directory, "data.xlsx");
    const outputPath = path.join(directory, "output.pptx");
    await writeTemplate(templatePath, [
      slideXml([{ id: 2, name: "Field", runs: [{ text: "{{client_name}}" }] }]),
    ]);
    await writeWorkbook(workbookPath, [["client_name"], ["Company A"]]);

    const valid = {
      outputPath,
      templatePath,
      templateSlide: 1,
      workbookPath,
      worksheet: "Companies",
    };
    await expectErrorCode(
      populatePowerPointTemplate({
        ...valid,
        templatePath: path.join(directory, "missing.pptx"),
      }),
      "PPTX_TEMPLATE_NOT_FOUND",
    );
    await expectErrorCode(
      populatePowerPointTemplate({
        ...valid,
        workbookPath: path.join(directory, "missing.xlsx"),
      }),
      "XLSX_WORKBOOK_NOT_FOUND",
    );
    await expectErrorCode(
      populatePowerPointTemplate({ ...valid, templateSlide: 2 }),
      "PPTX_TEMPLATE_SLIDE_NOT_FOUND",
    );
    await expectErrorCode(
      populatePowerPointTemplate({
        ...valid,
        outputPath: templatePath,
      }),
      "FILES_INPUT_OVERWRITE",
    );
    await expectErrorCode(
      populatePowerPointTemplate({
        ...valid,
        outputPath: path.join(directory, "output.pdf"),
      }),
      "PPTX_UNSUPPORTED_OUTPUT_TYPE",
    );
  });

  it("protects an existing output unless overwrite is enabled and commits atomically", async () => {
    const directory = await createTemporaryDirectory();
    const templatePath = path.join(directory, "template.pptx");
    const workbookPath = path.join(directory, "data.xlsx");
    const outputPath = path.join(directory, "output.pptx");
    const previousOutput = Buffer.from("previous output");
    await writeTemplate(templatePath, [
      slideXml([{ id: 2, name: "Field", runs: [{ text: "{{client_name}}" }] }]),
    ]);
    await writeWorkbook(workbookPath, [["client_name"], ["Company A"]]);
    await writeFile(outputPath, previousOutput);
    const options = {
      outputPath,
      templatePath,
      templateSlide: 1,
      workbookPath,
      worksheet: "Companies",
    };

    await expectErrorCode(
      populatePowerPointTemplate(options),
      "FILES_OUTPUT_EXISTS",
    );
    expect(await readFile(outputPath)).toEqual(previousOutput);

    await writeWorkbook(workbookPath, [["different_header"], ["Company A"]]);
    await expectErrorCode(
      populatePowerPointTemplate({ ...options, overwrite: true }),
      "PPTX_MISSING_EXCEL_COLUMN",
    );
    expect(await readFile(outputPath)).toEqual(previousOutput);

    await writeWorkbook(workbookPath, [["client_name"], ["Company A"]]);
    await populatePowerPointTemplate({ ...options, overwrite: true });
    expect(await readFile(outputPath)).not.toEqual(previousOutput);
    expect(await outputSlideXmls(outputPath)).toHaveLength(1);
    expect(
      (await readdir(directory)).filter((entry) =>
        entry.startsWith(".consultchimps-pptx-"),
      ),
    ).toEqual([]);
  });
});
