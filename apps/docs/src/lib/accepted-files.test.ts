import { describe, expect, it } from "vitest";

import {
  PDF_FILES,
  PRESENTATION_FILES,
  WORKBOOK_FILES,
  type AcceptedFileKind,
} from "./accepted-files";

const KINDS: ReadonlyArray<readonly [string, AcceptedFileKind]> = [
  ["workbooks", WORKBOOK_FILES],
  ["presentations", PRESENTATION_FILES],
  ["PDFs", PDF_FILES],
];

/**
 * A file the way a picker hands one over: a name, and whatever media type the
 * browser decided to report, including nothing at all.
 */
function pickedFile(name: string, type = ""): File {
  return new File([new Uint8Array([0x50, 0x4b])], name, { type });
}

describe.each(KINDS)("accepted %s", (_label, kind) => {
  it("accepts every extension it lists, whatever the browser reports", () => {
    for (const extension of kind.extensions) {
      expect(kind.accepts(pickedFile(`report${extension}`))).toBe(true);
      // Extensions are matched without regard to case, because a file picked
      // on Windows can arrive shouting.
      expect(kind.accepts(pickedFile(`report${extension.toUpperCase()}`))).toBe(
        true,
      );
    }
  });

  it("accepts every media type it lists, whatever the name says", () => {
    for (const mediaType of kind.mediaTypes) {
      expect(kind.accepts(pickedFile("report", mediaType))).toBe(true);
    }
  });

  it("offers every extension and media type through the accept attribute", () => {
    for (const value of [...kind.extensions, ...kind.mediaTypes]) {
      expect(kind.accept.split(",")).toContain(value);
    }
  });

  it("names every extension it accepts in both descriptions", () => {
    for (const extension of kind.extensions) {
      expect(kind.description).toContain(extension);
      expect(kind.pluralDescription).toContain(extension);
    }
  });

  it("strips every extension it lists and leaves other names alone", () => {
    for (const extension of kind.extensions) {
      expect(kind.stripExtension(`quarter review${extension}`)).toBe(
        "quarter review",
      );
      // Only the trailing extension goes; a name that merely contains one
      // keeps it.
      expect(kind.stripExtension(`quarter${extension}.zip`)).toBe(
        `quarter${extension}.zip`,
      );
    }
  });

  it("falls back to a media type it lists", () => {
    expect(kind.mediaTypes).toContain(kind.fallbackMediaType);
  });

  it("refuses a file of another shape", () => {
    expect(kind.accepts(pickedFile("notes.txt", "text/plain"))).toBe(false);
    for (const extension of kind.extensions) {
      // The separating dot is part of the extension: a name that merely ends
      // in the same letters is a different file.
      const letters = extension.slice(1);
      expect(kind.accepts(pickedFile(`report${letters}`))).toBe(false);
      expect(kind.stripExtension(`report${letters}`)).toBe(`report${letters}`);
    }
  });
});

describe("accepted workbooks", () => {
  it("accepts a macro-enabled workbook by name and by media type", () => {
    // The regression this module exists for: a genuine .xlsm arrives with its
    // own media type, which the pages' single-extension check rejected.
    expect(
      WORKBOOK_FILES.accepts(
        pickedFile(
          "clients.xlsm",
          "application/vnd.ms-excel.sheet.macroEnabled.12",
        ),
      ),
    ).toBe(true);
    // A browser lower-cases what it reports, and the registered media type for
    // a macro-enabled workbook is not all lower case, so a verbatim comparison
    // would match nothing a real picker ever hands over.
    expect(
      WORKBOOK_FILES.accepts(
        pickedFile(
          "clients.xlsm",
          "application/vnd.ms-excel.sheet.macroenabled.12",
        ),
      ),
    ).toBe(true);
    expect(
      WORKBOOK_FILES.accepts(
        pickedFile(
          "renamed-by-a-download",
          "application/vnd.ms-excel.sheet.macroenabled.12",
        ),
      ),
    ).toBe(true);
  });

  it("does not accept a presentation or a PDF", () => {
    expect(WORKBOOK_FILES.accepts(pickedFile("deck.pptx"))).toBe(false);
    expect(WORKBOOK_FILES.accepts(pickedFile("report.pdf"))).toBe(false);
  });
});
