import { InvalidArgumentError } from "commander";

/**
 * Parse an option that takes a whole number from 1. Only digits are accepted:
 * Number.parseInt would read "1.5" as 1 and "3rows" as 3, and the command would
 * then act on a row the reader never asked for. A refusal here is a usage
 * error, so it reports CLI_USAGE like any other.
 */
export function positiveInteger(value: string): number {
  const parsed = /^\d+$/u.test(value.trim()) ? Number(value) : Number.NaN;
  if (!Number.isSafeInteger(parsed) || parsed < 1) {
    throw new InvalidArgumentError("Expected a positive integer.");
  }
  return parsed;
}
