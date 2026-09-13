import { databaseError } from "../errors.js";
import type { ImportCell } from "./types.js";

const NUMBER_TEXT = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/u;

export function assertValidImportNumberCell(
  input: ImportCell,
  owner: "source" | "prepared",
): void {
  const cell =
    input.kind === "formula" && input.cached.kind !== "missing"
      ? input.cached
      : input;
  if (cell.kind !== "number" || NUMBER_TEXT.test(cell.raw)) return;
  if (owner === "prepared") {
    throw databaseError(
      "DB_INVALID_PREPARED_IMPORT",
      "The import plan contains an invalid captured number cell.",
    );
  }
  throw databaseError(
    "DB_INVALID_SOURCE_NUMBER",
    "A source number cell is invalid. Re-read the source and prepare the import again.",
  );
}
