import { isConsultChimpsError } from "@consultchimps/core";
import { describe, expect, it } from "vitest";

import { parseCsvTable } from "../src/index.js";

function codeOf(run: () => unknown): string | undefined {
  try {
    run();
  } catch (error) {
    return isConsultChimpsError(error) ? error.code : undefined;
  }
  return undefined;
}

describe("parseCsvTable", () => {
  it("reads a header row and typed-as-text values", () => {
    const table = parseCsvTable("Customer,Region\nAcme,North\nBeta,South\n");

    expect(table.columns).toEqual(["Customer", "Region"]);
    expect(table.rows).toEqual([
      { Customer: "Acme", Region: "North" },
      { Customer: "Beta", Region: "South" },
    ]);
  });

  it("keeps quoted fields, doubled quotes, and embedded separators intact", () => {
    const table = parseCsvTable(
      'Customer,Note\n"Acme, Ltd","He said ""yes"""\n"Line one\nLine two",plain\n',
    );

    expect(table.rows).toEqual([
      { Customer: "Acme, Ltd", Note: 'He said "yes"' },
      { Customer: "Line one\nLine two", Note: "plain" },
    ]);
  });

  it("accepts CRLF, a lone CR, a byte-order mark, and no trailing newline", () => {
    const table = parseCsvTable("﻿Customer,Region\r\nAcme,North\rBeta,South");

    expect(table.columns).toEqual(["Customer", "Region"]);
    expect(table.rows).toHaveLength(2);
    expect(table.rows[1]).toEqual({ Customer: "Beta", Region: "South" });
  });

  it("fills a blank header, numbers a repeated one, and skips blank rows", () => {
    const table = parseCsvTable(
      "Customer,,Region,Region\nAcme,x,North,N\n\n,,,\nBeta,y,South,S\n",
    );

    expect(table.columns).toEqual([
      "Customer",
      "column_2",
      "Region",
      "Region_2",
    ]);
    expect(table.rows).toHaveLength(2);
  });

  it("treats an empty field as no value and a short row as trailing blanks", () => {
    const table = parseCsvTable("Customer,Region,Score\nAcme,,\nBeta\n");

    expect(table.rows).toEqual([
      { Customer: "Acme", Region: null, Score: null },
      { Customer: "Beta", Region: null, Score: null },
    ]);
  });

  it("drops a trailing separator but refuses a value with no column", () => {
    expect(parseCsvTable("Customer\nAcme,\n").rows).toEqual([
      { Customer: "Acme" },
    ]);
    expect(codeOf(() => parseCsvTable("Customer\nAcme,North\n"))).toBe(
      "DB_CSV_EXTRA_VALUES",
    );
  });

  it("names the file's own row number when a row has a value with no column", () => {
    let message = "";
    try {
      parseCsvTable("\n\nCustomer\nAcme\nBeta,North\n");
    } catch (error) {
      message = error instanceof Error ? error.message : "";
    }
    expect(message).toContain("Row 5");
  });

  it("refuses a file with no rows and one that ends inside a quoted value", () => {
    expect(codeOf(() => parseCsvTable("\n \n"))).toBe("DB_CSV_NO_HEADER_ROW");
    expect(codeOf(() => parseCsvTable('Customer\n"Acme\n'))).toBe(
      "DB_CSV_UNCLOSED_QUOTE",
    );
  });

  it("names where an unclosed quoted value was opened", () => {
    let message = "";
    try {
      parseCsvTable('Customer,Region\nAcme,North\nBeta,"South\nmore\n');
    } catch (error) {
      message = error instanceof Error ? error.message : "";
    }
    expect(message).toContain("row 3, column 2");
  });

  it("refuses text after a closing quote rather than joining it on", () => {
    // Before the parser had a "quote just closed" state this imported as
    // "Acmex", which is neither what the file says nor a refusal.
    expect(codeOf(() => parseCsvTable('Customer\n"Acme"x\n'))).toBe(
      "DB_CSV_TEXT_AFTER_QUOTE",
    );
    expect(
      codeOf(() => parseCsvTable('Customer,Region\n"Acme" ,North\n')),
    ).toBe("DB_CSV_TEXT_AFTER_QUOTE");
    let message = "";
    try {
      parseCsvTable('Customer,Region\nAcme,"North"junk\n');
    } catch (error) {
      message = error instanceof Error ? error.message : "";
    }
    expect(message).toContain("Row 2, column 2");
  });

  it("counts a line break inside a quoted value toward the file's lines", () => {
    let message = "";
    try {
      // The quoted value spans lines 2 and 3, so the bad row is line 4.
      parseCsvTable('Customer,Region\n"Two\nlines",North\nBeta,South,extra\n');
    } catch (error) {
      message = error instanceof Error ? error.message : "";
    }
    expect(message).toContain("Row 4");
  });

  it("reads a quote that does not start a field as an ordinary character", () => {
    // A deliberate tolerance: RFC 4180 leaves this undefined, and a bare quote
    // in an unquoted field is far more often a measurement than a mistake.
    expect(parseCsvTable('Part\n5" pipe\n').rows).toEqual([
      { Part: '5" pipe' },
    ]);
  });

  it("keeps an empty quoted field at the end of a file", () => {
    expect(parseCsvTable('Customer,Region\nAcme,""').rows).toEqual([
      { Customer: "Acme", Region: null },
    ]);
  });

  it("refuses a separator it could not tell from its own syntax", () => {
    expect(codeOf(() => parseCsvTable("a\nb", { delimiter: "" }))).toBe(
      "DB_CSV_INVALID_DELIMITER",
    );
    expect(codeOf(() => parseCsvTable("a\nb", { delimiter: '"' }))).toBe(
      "DB_CSV_INVALID_DELIMITER",
    );
    expect(codeOf(() => parseCsvTable("a\nb", { delimiter: "\n" }))).toBe(
      "DB_CSV_INVALID_DELIMITER",
    );
    expect(
      parseCsvTable("Customer;Region\nAcme;North", { delimiter: ";" }).rows,
    ).toEqual([{ Customer: "Acme", Region: "North" }]);
  });

  it("records the source file name when one is given", () => {
    expect(
      parseCsvTable("Customer\nAcme", { file: "customers.csv" }).source,
    ).toEqual({ file: "customers.csv" });
  });
});
