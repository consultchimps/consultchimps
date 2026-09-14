import { databaseError } from "../errors.js";
import { assertSafeIdentifier, MAX_IDENTIFIER_LENGTH } from "../schema.js";
import type {
  ImportProfile,
  ImportRoute,
  ImportSource,
  ImportNaming,
  ImportDestination,
} from "./types.js";

function takeCodePoints(value: string, limit: number): string {
  return Array.from(value).slice(0, limit).join("");
}

function truncateUtf16(value: string, limit: number): string {
  if (value.length <= limit) return value;
  let result = "";
  for (const character of value) {
    if (result.length + character.length > limit) break;
    result += character;
  }
  return result;
}

function safeSuggestedName(value: string): string {
  const normalized = value
    .normalize("NFKC")
    .trim()
    .replace(/[^\p{L}\p{N}_]+/gu, "_")
    .replace(/^_+|_+$/gu, "");
  const truncated = truncateUtf16(normalized, MAX_IDENTIFIER_LENGTH);
  const result =
    truncated.length === 0 || truncated.startsWith("_consultchimps")
      ? "Imported_Data"
      : truncated;
  assertSafeIdentifier(result, "table");
  return result;
}

function initialsPrefixFor(name: string): string {
  const letters = name
    .split(/[^\p{L}\p{N}]+/gu)
    .filter((part) => part.length > 0)
    .map((part) => takeCodePoints(part, 1))
    .join("")
    .toUpperCase();
  return takeCodePoints((letters || takeCodePoints(name, 3)).toUpperCase(), 8);
}

type RecordIdPrefixPolicy = "initials" | "table-name";

function recordIdPrefixFor(name: string, policy: RecordIdPrefixPolicy): string {
  return policy === "table-name"
    ? takeCodePoints(name, 8).toUpperCase()
    : initialsPrefixFor(name);
}

function validateImportNaming(value: unknown): ImportNaming | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw databaseError(
      "DB_INVALID_IMPORT_NAMING",
      "The import naming policy is invalid. Choose source-or-selection, selection-label, or a single table name.",
      { field: "naming" },
    );
  }
  const fields = value as Record<string, unknown>;
  if (
    fields["kind"] === "source-or-selection" ||
    fields["kind"] === "selection-label"
  ) {
    return { kind: fields["kind"] };
  }
  if (fields["kind"] === "single-table" && typeof fields["name"] === "string") {
    return { kind: "single-table", name: fields["name"] };
  }
  throw databaseError(
    "DB_INVALID_IMPORT_NAMING",
    "The import naming policy is invalid. Choose source-or-selection, selection-label, or a single table name.",
    { field: "naming" },
  );
}

export function suggestImportDestination(
  label: string,
  prefixPolicy: RecordIdPrefixPolicy = "initials",
): Extract<ImportDestination, { readonly kind: "new-table-infer" }> {
  const name = safeSuggestedName(label);
  return {
    kind: "new-table-infer",
    name,
    recordId: { prefix: recordIdPrefixFor(name, prefixPolicy), padding: 6 },
  };
}

export async function draftImportProfile(options: {
  readonly sources: readonly ImportSource[];
  readonly naming?: ImportNaming | undefined;
}): Promise<ImportProfile> {
  if (options.sources.length === 0) {
    throw databaseError(
      "DB_IMPORT_NO_SOURCES",
      "Choose at least one source before drafting an import profile.",
    );
  }
  const naming = validateImportNaming(options.naming) ?? {
    kind: "source-or-selection",
  };
  const routes: ImportRoute[] = [];
  let selectionCount = 0;
  for (const source of options.sources) {
    for (const selection of source.selections) {
      selectionCount += 1;
      const destination = suggestImportDestination(
        naming.kind === "single-table"
          ? naming.name
          : naming.kind === "selection-label" || source.selections.length !== 1
            ? selection.label
            : source.key,
        naming.kind === "selection-label" ? "table-name" : "initials",
      );
      routes.push({
        source: source.key,
        selection: selection.key,
        destination,
        columns: [],
      });
    }
  }
  if (selectionCount === 0) {
    throw databaseError(
      "DB_IMPORT_NO_SELECTIONS",
      "Choose at least one source region before importing. If the workbook contains only hidden worksheets, include hidden sheets and try again.",
    );
  }
  if (naming.kind === "single-table" && selectionCount !== 1) {
    throw databaseError(
      "DB_IMPORT_INTO_AMBIGUOUS",
      "The destination override can be used only when exactly one source region is selected.",
    );
  }
  return { version: 1, routes };
}
