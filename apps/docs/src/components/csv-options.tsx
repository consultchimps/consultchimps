"use client";

/**
 * How the sheets tools read the CSV files a visitor chooses (ADR 0007): the
 * browser's twin of the command line's `--csv-*` flags. Every field stays text
 * unless the visitor asks for numbers or dates, and the encoding and delimiter
 * are detected unless one is chosen.
 */
import { useId, useState } from "react";

import { inputClass } from "@/components/tool-kit";
import type { CsvReadOptions } from "@consultchimps/xlsx/bytes";

/** What the visitor chose, as the form holds it; "" means detect or off. */
export interface CsvChoices {
  readonly dates: "" | "iso" | "dmy" | "mdy" | "ymd";
  readonly decimal: "." | ",";
  readonly delimiter: "" | "," | ";" | "\t" | "|";
  readonly encoding: "" | "utf-8" | "utf-16le" | "utf-16be" | "windows-1252";
  readonly numbers: boolean;
  readonly thousands: "default" | "none" | "," | "." | " " | "'";
}

export const DEFAULT_CSV_CHOICES: CsvChoices = {
  dates: "",
  decimal: ".",
  delimiter: "",
  encoding: "",
  numbers: false,
  thousands: "default",
};

/** Whether a chosen file is read as CSV, by the same rule the library uses. */
export function isCsvFileName(name: string): boolean {
  return /\.csv$/iu.test(name);
}

/** The library's options for these choices, or undefined for the defaults. */
export function csvReadOptionsFrom(
  choices: CsvChoices,
): CsvReadOptions | undefined {
  const options: CsvReadOptions = {};
  if (choices.delimiter !== "") options.delimiter = choices.delimiter;
  if (choices.encoding !== "") options.encoding = choices.encoding;
  if (choices.numbers) {
    options.numbers = true;
    options.decimalSeparator = choices.decimal;
    if (choices.thousands !== "default") {
      options.thousandsSeparator =
        choices.thousands === "none" ? "" : choices.thousands;
    }
  }
  if (choices.dates !== "") options.dates = choices.dates;
  return Object.keys(options).length === 0 ? undefined : options;
}

/** The choices and a setter for one of them. */
export function useCsvChoices(): [
  CsvChoices,
  <K extends keyof CsvChoices>(key: K, value: CsvChoices[K]) => void,
] {
  const [choices, setChoices] = useState<CsvChoices>(DEFAULT_CSV_CHOICES);
  return [
    choices,
    (key, value) => {
      setChoices((current) => ({ ...current, [key]: value }));
    },
  ];
}

const fieldLabelClass = "block text-sm font-semibold";
const fieldHintClass = "mt-1 text-sm text-fd-muted-foreground";
const checkboxLabelClass =
  "flex items-start gap-2.5 text-sm font-medium leading-6";
const checkboxClass =
  "mt-1 size-4 shrink-0 rounded border-fd-border accent-fd-primary";

interface SelectFieldProps<T extends string> {
  readonly disabled: boolean;
  readonly label: string;
  readonly onChange: (value: T) => void;
  readonly options: ReadonlyArray<readonly [T, string]>;
  readonly testId: string;
  readonly value: T;
}

function SelectField<T extends string>({
  disabled,
  label,
  onChange,
  options,
  testId,
  value,
}: SelectFieldProps<T>) {
  const id = useId();
  return (
    <div>
      <label className={fieldLabelClass} htmlFor={id}>
        {label}
      </label>
      <select
        className={`${inputClass} mt-2`}
        data-testid={testId}
        disabled={disabled}
        id={id}
        onChange={(event) => onChange(event.target.value as T)}
        value={value}
      >
        {options.map(([optionValue, text]) => (
          <option key={optionValue} value={optionValue}>
            {text}
          </option>
        ))}
      </select>
    </div>
  );
}

export interface CsvOptionsFieldsProps {
  readonly choices: CsvChoices;
  readonly disabled: boolean;
  readonly onChange: <K extends keyof CsvChoices>(
    key: K,
    value: CsvChoices[K],
  ) => void;
}

/** The CSV reading controls, for a page where a CSV file is chosen. */
export function CsvOptionsFields({
  choices,
  disabled,
  onChange,
}: CsvOptionsFieldsProps) {
  return (
    <fieldset className="mt-6 flex flex-col gap-4" data-testid="csv-options">
      <legend className={fieldLabelClass}>Reading CSV files</legend>
      <p className={fieldHintClass}>
        Each CSV file is one worksheet named after the file. Every field stays
        text unless you ask for numbers or dates, and a field that is not wholly
        a number or a date is never changed
      </p>
      <div className="grid gap-4 sm:grid-cols-2">
        <SelectField
          disabled={disabled}
          label="Delimiter"
          onChange={(value) => onChange("delimiter", value)}
          options={[
            ["", "Detect from the file"],
            [",", "Comma"],
            [";", "Semicolon"],
            ["\t", "Tab"],
            ["|", "Pipe"],
          ]}
          testId="csv-delimiter-select"
          value={choices.delimiter}
        />
        <SelectField
          disabled={disabled}
          label="Encoding"
          onChange={(value) => onChange("encoding", value)}
          options={[
            ["", "Detect from the file"],
            ["utf-8", "UTF-8"],
            ["utf-16le", "UTF-16 LE"],
            ["utf-16be", "UTF-16 BE"],
            ["windows-1252", "Windows-1252"],
          ]}
          testId="csv-encoding-select"
          value={choices.encoding}
        />
      </div>
      <label className={checkboxLabelClass}>
        <input
          checked={choices.numbers}
          className={checkboxClass}
          data-testid="csv-numbers-checkbox"
          disabled={disabled}
          onChange={(event) => onChange("numbers", event.target.checked)}
          type="checkbox"
        />
        <span>
          Read numbers
          <span className="block text-sm font-normal text-fd-muted-foreground">
            A field that holds only a number, such as 1,234.50, becomes a
            number. Leading zeros, currency signs, and more than 15 digits stay
            text
          </span>
        </span>
      </label>
      {choices.numbers ? (
        <div className="grid gap-4 sm:grid-cols-2">
          <SelectField
            disabled={disabled}
            label="Decimal separator"
            onChange={(value) => onChange("decimal", value)}
            options={[
              [".", "Point (1234.5)"],
              [",", "Comma (1234,5)"],
            ]}
            testId="csv-decimal-select"
            value={choices.decimal}
          />
          <SelectField
            disabled={disabled}
            label="Thousands separator"
            onChange={(value) => onChange("thousands", value)}
            options={[
              [
                "default",
                choices.decimal === "," ? "Point (default)" : "Comma (default)",
              ],
              ["none", "None"],
              [",", "Comma"],
              [".", "Point"],
              [" ", "Space"],
              ["'", "Apostrophe"],
            ]}
            testId="csv-thousands-select"
            value={choices.thousands}
          />
        </div>
      ) : null}
      <SelectField
        disabled={disabled}
        label="Dates"
        onChange={(value) => onChange("dates", value)}
        options={[
          ["", "Keep as text"],
          ["iso", "Read yyyy-mm-dd only"],
          ["dmy", "Day first, such as 31/01/2025"],
          ["mdy", "Month first, such as 01/31/2025"],
          ["ymd", "Year first, such as 2025/01/31"],
        ]}
        testId="csv-dates-select"
        value={choices.dates}
      />
    </fieldset>
  );
}
