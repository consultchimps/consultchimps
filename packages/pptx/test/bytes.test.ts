import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

import {
  isConsultChimpsError,
  OPERATION_ABORTED,
  type OperationProgress,
} from "@consultchimps/core";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import {
  inspectPresentationBytes,
  inspectPresentationOutcomeBytes,
  planPopulatePresentationBytes,
  populatePresentationBytes,
} from "../src/bytes.js";

// A DOS timestamp has two-second resolution, so a shorter pause could hide a
// clock-dependent byte difference inside one tick.
const CLOCK_TICK = 2_100;

function encodeXml(value: string): string {
  return value
    .replace(/&/gu, "&amp;")
    .replace(/</gu, "&lt;")
    .replace(/>/gu, "&gt;");
}

function slideXml(runs: string[]): string {
  const shape = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${runs
    .map(
      (text) =>
        `<a:r><a:rPr lang="en-US" b="1"/><a:t>${encodeXml(text)}</a:t></a:r>`,
    )
    .join("")}</a:p></p:txBody></p:sp>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shape}</p:spTree></p:cSld></p:sld>`;
}

/**
 * A template package built the way PowerPoint writes one: parts only, with no
 * directory entries.
 */
async function templateBytes(slides: string[]): Promise<Uint8Array> {
  const zip = new JSZip();
  const write = (partPath: string, content: string): void => {
    zip.file(partPath, content, { createFolders: false });
  };

  write(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${slides
      .map(
        (_, index) =>
          `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
      )
      .join("")}</Types>`,
  );
  write(
    "_rels/.rels",
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
  );
  write(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst>${slides
      .map(
        (_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 1}"/>`,
      )
      .join(
        "",
      )}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`,
  );
  write(
    "ppt/_rels/presentation.xml.rels",
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${slides
      .map(
        (_, index) =>
          `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index + 1}.xml"/>`,
      )
      .join("")}</Relationships>`,
  );
  slides.forEach((slide, index) => {
    write(`ppt/slides/slide${index + 1}.xml`, slide);
  });

  return zip.generateAsync({ compression: "DEFLATE", type: "uint8array" });
}

function workbookBytes(
  rows: Array<Array<null | number | string>>,
  worksheetName = "Companies",
): Uint8Array {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet(rows),
    worksheetName,
  );
  return new Uint8Array(
    XLSX.write(workbook, {
      bookType: "xlsx",
      compression: true,
      type: "array",
    }) as ArrayBuffer,
  );
}

async function outputSlides(bytes: Uint8Array): Promise<string[]> {
  const zip = await JSZip.loadAsync(bytes);
  const presentationXml = await zip
    .file("ppt/presentation.xml")!
    .async("string");
  const relationshipsXml = await zip
    .file("ppt/_rels/presentation.xml.rels")!
    .async("string");
  const relationshipIds = [
    ...presentationXml.matchAll(/\br:id="(?<id>[^"]+)"/gu),
  ].map((match) => match.groups?.id ?? "");
  return Promise.all(
    relationshipIds.map((relationshipId) => {
      const relationship = [
        ...relationshipsXml.matchAll(/<Relationship\b[^>]*\/>/gu),
      ].find((match) => match[0].includes(`Id="${relationshipId}"`));
      const target = /Target="(?<target>[^"]+)"/u.exec(relationship?.[0] ?? "")
        ?.groups?.target;
      return zip.file(`ppt/${target}`)!.async("string");
    }),
  );
}

const singlePlaceholderTemplate = (): Promise<Uint8Array> =>
  templateBytes([slideXml(["{{client}}", " reports ", "{{amount}}"])]);

describe("byte-level presentation population", () => {
  it("generates one slide per supplied record and keeps run formatting", async () => {
    const events: OperationProgress[] = [];
    const { result, outputs } = await populatePresentationBytes({
      template: {
        name: "client template.pptx",
        bytes: await singlePlaceholderTemplate(),
      },
      records: [
        { amount: "10", client: "North" },
        { amount: "20", client: "South" },
      ],
      onProgress: (progress) => events.push(progress),
    });

    expect(result.operation).toBe("pptx.populate");
    expect(result.metrics).toEqual({
      generatedSlides: 2,
      inputRows: 2,
      outputFiles: 1,
      placeholderFields: 2,
      placeholderOccurrences: 2,
      replacements: 4,
      skippedRows: 0,
      warnings: 0,
    });
    expect(result.warnings).toEqual([]);
    expect(outputs).toHaveLength(1);
    expect(outputs[0]?.name).toBe("client template-populated.pptx");
    expect(outputs[0]?.mediaType).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    );
    expect(result.artifacts.map((artifact) => artifact.path)).toEqual([
      "client template-populated.pptx",
    ]);
    expect(events.map((event) => [event.stage, event.completed])).toEqual([
      ["generating-slides", 1],
      ["generating-slides", 2],
    ]);

    const slides = await outputSlides(outputs[0]!.bytes);
    expect(slides).toHaveLength(2);
    expect(slides[0]).toContain("<a:t>North</a:t>");
    expect(slides[0]).toContain("<a:t>10</a:t>");
    expect(slides[0]).toContain('<a:rPr lang="en-US" b="1"/>');
    expect(slides[1]).toContain("<a:t>South</a:t>");
    expect(slides[1]).toContain("<a:t>20</a:t>");
  });

  it("reads records from workbook bytes and reports skipped rows", async () => {
    const { result, outputs } = await populatePresentationBytes({
      template: {
        name: "template.pptx",
        bytes: await singlePlaceholderTemplate(),
      },
      workbook: {
        name: "companies.xlsx",
        bytes: workbookBytes([
          ["client", "amount"],
          ["North", 10],
          [null, null],
          ["South", 20],
        ]),
      },
      outputName: "client profiles.pptx",
    });

    expect(outputs[0]?.name).toBe("client profiles.pptx");
    expect(result.warnings).toEqual(["Skipped 1 empty worksheet row."]);
    expect(result.metrics).toMatchObject({
      generatedSlides: 2,
      inputRows: 2,
      skippedRows: 1,
      warnings: 1,
    });

    const slides = await outputSlides(outputs[0]!.bytes);
    expect(slides[0]).toContain("<a:t>North</a:t>");
    expect(slides[1]).toContain("<a:t>South</a:t>");
  });

  it("plans a population without producing any bytes", async () => {
    const template = {
      name: "template.pptx",
      bytes: await singlePlaceholderTemplate(),
    };
    const workbook = {
      name: "companies.xlsx",
      bytes: workbookBytes([
        ["client", "amount"],
        ["North", 10],
        [null, null],
      ]),
    };

    const plan = await planPopulatePresentationBytes({ template, workbook });
    expect(plan.operation).toBe("pptx.populate");
    expect(plan.inputs).toEqual(["template.pptx", "companies.xlsx"]);
    expect(plan.outputs).toEqual([
      {
        kind: "file",
        mediaType:
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        path: "template-populated.pptx",
        exists: false,
      },
    ]);
    expect(plan.warnings).toEqual(["Skipped 1 empty worksheet row."]);
    expect(plan.metrics).toEqual({
      generatedSlides: 1,
      inputRows: 1,
      outputFiles: 1,
      placeholderFields: 2,
      placeholderOccurrences: 2,
      skippedRows: 1,
    });

    const recordsPlan = await planPopulatePresentationBytes({
      template,
      records: [{ amount: "1", client: "A" }],
    });
    expect(recordsPlan.inputs).toEqual(["template.pptx"]);
    expect(recordsPlan.warnings).toEqual([]);
  });

  it("inspects a template slide", async () => {
    const template = {
      name: "template.pptx",
      bytes: await templateBytes([
        slideXml(["{{client}}", "{{client}}"]),
        slideXml(["{{region}}"]),
      ]),
    };

    expect(await inspectPresentationBytes(template)).toEqual({
      malformedPlaceholderCount: 0,
      placeholderOccurrences: 2,
      placeholders: [{ name: "client", occurrences: 2 }],
      slideNumber: 1,
      unsupportedPlacementPlaceholders: [],
      unsupportedSplitRunPlaceholders: [],
    });
    expect(
      await inspectPresentationBytes(template, { templateSlide: 2 }),
    ).toMatchObject({
      placeholders: [{ name: "region", occurrences: 1 }],
      slideNumber: 2,
    });
    expect(
      (
        await populatePresentationBytes({
          template,
          templateSlide: 2,
          records: [{ region: "North" }],
        })
      ).result.metrics.generatedSlides,
    ).toBe(1);
  });

  it("reports an inspection as a structured operation result", async () => {
    const template = {
      name: "template.pptx",
      bytes: await templateBytes([slideXml(["{{client}}", "{{client}}"])]),
    };

    const outcome = await inspectPresentationOutcomeBytes(template);
    expect(outcome.inspection).toEqual(
      await inspectPresentationBytes(template),
    );
    expect(outcome.result).toEqual({
      operation: "pptx.inspect-template",
      // An inspection reads one slide and writes nothing.
      artifacts: [],
      warnings: [],
      metrics: {
        malformedPlaceholderLocations: 0,
        placeholderFields: 1,
        placeholderOccurrences: 2,
        unsupportedPlacementPlaceholders: 0,
        unsupportedSplitRunPlaceholders: 0,
      },
    });

    // Identical inputs and options must produce an identical result, including
    // the order of the warnings.
    expect(await inspectPresentationOutcomeBytes(template)).toEqual(outcome);
  });

  it("warns about every condition that would make a populate refuse", async () => {
    const malformed = await inspectPresentationOutcomeBytes({
      name: "malformed.pptx",
      // One unbalanced brace, which leaves the slide with no usable
      // placeholder either, so both warnings apply.
      bytes: await templateBytes([slideXml(["{{client}"])]),
    });
    expect(malformed.result.metrics).toEqual({
      malformedPlaceholderLocations: 1,
      placeholderFields: 0,
      placeholderOccurrences: 0,
      unsupportedPlacementPlaceholders: 0,
      unsupportedSplitRunPlaceholders: 0,
    });
    expect(malformed.result.warnings).toEqual([
      "Slide 1 has 1 location with malformed placeholder braces. A populate would refuse this template; use the exact {{field_name}} syntax.",
      "Slide 1 does not contain any valid {{field_name}} placeholders. A populate would refuse this template.",
    ]);

    // A run outside any <p:sp> text shape: the placeholder is found, but a
    // populate cannot fill it.
    const outsideShape = await inspectPresentationOutcomeBytes({
      name: "outside-shape.pptx",
      bytes: await templateBytes([
        slideXml(["{{client}}"]).replace(
          "</p:spTree>",
          "<a:t>{{region}}</a:t></p:spTree>",
        ),
      ]),
    });
    expect(outsideShape.result.metrics.unsupportedPlacementPlaceholders).toBe(
      1,
    );
    expect(outsideShape.result.warnings).toEqual([
      'Placeholders outside a supported text shape are not populated: "region". A populate would refuse this template.',
    ]);

    // Every warned condition is one the populate operation refuses outright.
    await expect(
      populatePresentationBytes({
        template: {
          name: "malformed.pptx",
          bytes: await templateBytes([slideXml(["{{client}"])]),
        },
        records: [{ client: "A" }],
      }),
    ).rejects.toMatchObject({ code: "PPTX_MALFORMED_PLACEHOLDER" });
  });

  it("requires exactly one data source", async () => {
    const template = {
      name: "template.pptx",
      bytes: await singlePlaceholderTemplate(),
    };
    const workbook = {
      name: "companies.xlsx",
      bytes: workbookBytes([
        ["client", "amount"],
        ["North", 10],
      ]),
    };

    await expect(populatePresentationBytes({ template })).rejects.toMatchObject(
      {
        code: "PPTX_INVALID_DATA_SOURCE",
        details: { records: false, workbook: false },
      },
    );
    await expect(
      populatePresentationBytes({
        template,
        records: [{ amount: "1", client: "A" }],
        workbook,
      }),
    ).rejects.toMatchObject({
      code: "PPTX_INVALID_DATA_SOURCE",
      details: { records: true, workbook: true },
    });
  });

  it("reports templates and records that cannot be populated", async () => {
    const template = {
      name: "template.pptx",
      bytes: await singlePlaceholderTemplate(),
    };

    await expect(
      populatePresentationBytes({ template, records: [] }),
    ).rejects.toMatchObject({ code: "PPTX_NO_DATA_ROWS" });
    await expect(
      populatePresentationBytes({
        template,
        records: [{ client: "A" }],
      }),
    ).rejects.toMatchObject({
      code: "PPTX_MISSING_EXCEL_COLUMN",
      details: { missingColumns: ["amount"] },
    });
    await expect(
      populatePresentationBytes({
        template,
        records: [{ amount: "1", client: "A" }],
        templateSlide: 4,
      }),
    ).rejects.toMatchObject({ code: "PPTX_TEMPLATE_SLIDE_NOT_FOUND" });
    await expect(
      populatePresentationBytes({
        template,
        records: [{ amount: "1", client: "A" }],
        templateSlide: 0,
      }),
    ).rejects.toMatchObject({ code: "PPTX_INVALID_TEMPLATE_SLIDE" });
    await expect(
      populatePresentationBytes({
        template: { name: "junk.pptx", bytes: new Uint8Array([1, 2, 3]) },
        records: [{ amount: "1", client: "A" }],
      }),
    ).rejects.toMatchObject({ code: "PPTX_INVALID_TEMPLATE" });
    await expect(
      populatePresentationBytes({
        template: {
          name: "plain.pptx",
          bytes: await templateBytes([slideXml(["No placeholders here"])]),
        },
        records: [{ amount: "1", client: "A" }],
      }),
    ).rejects.toMatchObject({ code: "PPTX_NO_PLACEHOLDERS" });
    await expect(
      populatePresentationBytes({
        template: {
          name: "broken.pptx",
          bytes: await templateBytes([slideXml(["{{client}} {{ oops"])]),
        },
        records: [{ amount: "1", client: "A" }],
      }),
    ).rejects.toMatchObject({ code: "PPTX_MALFORMED_PLACEHOLDER" });
  });

  it("sanitizes the output name", async () => {
    const records = [{ amount: "1", client: "A" }];
    const bytes = await singlePlaceholderTemplate();

    const unsafe = await populatePresentationBytes({
      template: { name: "cli*ent<>deck.pptx", bytes },
      records,
    });
    expect(unsafe.outputs[0]?.name).toBe("cli-ent-deck-populated.pptx");

    const reserved = await populatePresentationBytes({
      template: { name: "template.pptx", bytes },
      outputName: "aux.pptx",
      records,
    });
    expect(reserved.outputs[0]?.name).toBe("_aux.pptx");

    const bounded = await populatePresentationBytes({
      template: { name: `${"\u4E2D".repeat(100)}.pptx`, bytes },
      records,
    });
    const name = bounded.outputs[0]!.name;
    expect(Buffer.byteLength(name, "utf8")).toBeLessThanOrEqual(255);
    expect(name.endsWith(".pptx")).toBe(true);
    expect([...name.slice(0, -".pptx".length)]).toHaveLength(26);

    const blank = await populatePresentationBytes({
      template: { name: "template.pptx", bytes },
      outputName: "  ",
      records,
    });
    expect(blank.outputs[0]?.name).toBe("presentation.pptx");
  });

  it("produces byte-identical presentations for identical inputs", async () => {
    const options = {
      template: {
        name: "template.pptx",
        bytes: await singlePlaceholderTemplate(),
      },
      records: [
        { amount: "10", client: "North" },
        { amount: "20", client: "South" },
      ],
    };

    const first = await populatePresentationBytes(options);
    await new Promise((resolve) => setTimeout(resolve, CLOCK_TICK));
    const second = await populatePresentationBytes(options);

    expect(
      Buffer.compare(
        Buffer.from(first.outputs[0]!.bytes),
        Buffer.from(second.outputs[0]!.bytes),
      ),
    ).toBe(0);
  });

  it("cancels without producing partial output", async () => {
    const template = {
      name: "template.pptx",
      bytes: await singlePlaceholderTemplate(),
    };
    const records = [
      { amount: "10", client: "North" },
      { amount: "20", client: "South" },
    ];

    const beforeStart = new AbortController();
    beforeStart.abort();
    let thrown: unknown;
    try {
      await populatePresentationBytes({
        template,
        records,
        signal: beforeStart.signal,
      });
    } catch (error) {
      thrown = error;
    }
    expect(isConsultChimpsError(thrown)).toBe(true);
    expect((thrown as { code: string }).code).toBe(OPERATION_ABORTED);
    expect((thrown as Error).message).toContain(
      "no partial output was produced",
    );
    expect((thrown as Error).message).not.toContain("output files");

    // Cancelling between slides stops the remaining slides.
    const midway = new AbortController();
    const completed: number[] = [];
    await expect(
      populatePresentationBytes({
        template,
        records,
        signal: midway.signal,
        onProgress: (progress) => {
          completed.push(progress.completed);
          midway.abort();
        },
      }),
    ).rejects.toMatchObject({ code: OPERATION_ABORTED });
    expect(completed).toEqual([1]);

    // Cancelling while the package is serialized still returns nothing.
    const atEnd = new AbortController();
    await expect(
      populatePresentationBytes({
        template,
        records: [records[0]!],
        signal: atEnd.signal,
        onProgress: () => atEnd.abort(),
      }),
    ).rejects.toMatchObject({ code: OPERATION_ABORTED });
  });

  it("cancels an inspection whose answer is already superseded", async () => {
    const template = {
      name: "template.pptx",
      bytes: await templateBytes([slideXml(["{{client}}"])]),
    };

    const cancelled = new AbortController();
    cancelled.abort();
    await expect(
      inspectPresentationBytes(template, { signal: cancelled.signal }),
    ).rejects.toMatchObject({ code: OPERATION_ABORTED });
    await expect(
      inspectPresentationOutcomeBytes(template, { signal: cancelled.signal }),
    ).rejects.toMatchObject({ code: OPERATION_ABORTED });

    // Cancelling while the slide part is being decompressed still rejects,
    // rather than returning a report nobody is waiting for.
    const duringSlideRead = new AbortController();
    await expect(
      inspectPresentationBytes(template, {
        signal: duringSlideRead.signal,
        onProgress: () => duringSlideRead.abort(),
      }),
    ).rejects.toMatchObject({ code: OPERATION_ABORTED });

    // A signal that never aborts leaves the report exactly as it was, and the
    // progress events are the same for every template.
    const live = new AbortController();
    const events: OperationProgress[] = [];
    await expect(
      inspectPresentationBytes(template, {
        signal: live.signal,
        onProgress: (progress) => events.push(progress),
      }),
    ).resolves.toMatchObject({ placeholders: [{ name: "client" }] });
    expect(events).toEqual([
      {
        operation: "pptx.inspect-template",
        stage: "reading-slide",
        completed: 1,
        total: 2,
      },
      {
        operation: "pptx.inspect-template",
        stage: "inspecting-placeholders",
        completed: 2,
        total: 2,
      },
    ]);
  });

  it("cancels a plan instead of parsing packages nobody is waiting for", async () => {
    const template = {
      name: "template.pptx",
      bytes: await singlePlaceholderTemplate(),
    };
    const records = [{ amount: "1", client: "A" }];

    // Planning reads both packages in full, so a caller that has moved on,
    // such as a page replanning after a keystroke, must be able to stop that
    // work rather than only discard its answer.
    const cancelled = new AbortController();
    cancelled.abort();
    await expect(
      planPopulatePresentationBytes({
        template,
        records,
        signal: cancelled.signal,
      }),
    ).rejects.toMatchObject({ code: OPERATION_ABORTED });

    // A signal that never aborts leaves the plan exactly as it was.
    const live = new AbortController();
    await expect(
      planPopulatePresentationBytes({ template, records, signal: live.signal }),
    ).resolves.toMatchObject({
      operation: "pptx.populate",
      outputs: [{ path: "template-populated.pptx" }],
    });
  });
});

describe("byte entry point packaging", () => {
  it("keeps the built bytes entry free of node imports", async () => {
    const distDirectory = new URL("../dist/", import.meta.url);
    const visited = new Set<string>();
    const queue = ["bytes.js"];

    while (queue.length > 0) {
      const entry = queue.pop()!;
      if (visited.has(entry)) {
        continue;
      }
      visited.add(entry);
      const source = await readFile(
        fileURLToPath(new URL(entry, distDirectory)),
        "utf8",
      );
      expect(source, `${entry} must not import node builtins`).not.toMatch(
        /["']node:/u,
      );
      for (const match of source.matchAll(/from\s+["']\.\/([^"']+)["']/gu)) {
        queue.push(match[1]!);
      }
    }

    expect(visited.size).toBeGreaterThan(0);
  });
});
