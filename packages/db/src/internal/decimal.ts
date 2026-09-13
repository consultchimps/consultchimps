const DECIMAL_PATTERN =
  /^[+-]?(?:(\d+)(?:\.(\d*))?|\.(\d+))(?:[eE]([+-]?\d+))?$/u;
const MAX_DECIMAL_PRECISION = 38n;
const MAX_EXPONENT_DIGITS = 8;

interface ParsedDecimal {
  readonly negative: boolean;
  readonly digits: string;
  readonly decimalPosition: bigint;
  readonly zero: boolean;
}

export interface ExactDecimal {
  readonly normalized: string;
  readonly integerDigits: number;
  readonly scale: number;
}

function parseDecimal(value: string): ParsedDecimal | undefined {
  const match = DECIMAL_PATTERN.exec(value);
  if (match === null) return undefined;
  const integer = match[1] ?? "";
  const fraction = match[2] ?? match[3] ?? "";
  const coefficient = `${integer}${fraction}`;
  const firstNonzero = coefficient.search(/[1-9]/u);
  if (firstNonzero === -1) {
    return {
      negative: false,
      digits: "0",
      decimalPosition: 0n,
      zero: true,
    };
  }
  const exponentText = match[4] ?? "0";
  if (exponentText.replace(/^[+-]?0*/u, "").length > MAX_EXPONENT_DIGITS) {
    return undefined;
  }
  const exponent = BigInt(exponentText);
  return {
    negative: value.startsWith("-"),
    digits: coefficient.slice(firstNonzero),
    decimalPosition: BigInt(integer.length - firstNonzero) + exponent,
    zero: false,
  };
}

function magnitude(parsed: ParsedDecimal, digits = parsed.digits): string {
  const position = Number(parsed.decimalPosition);
  if (position <= 0) return `0.${"0".repeat(-position)}${digits}`;
  if (position >= digits.length) {
    return `${digits}${"0".repeat(position - digits.length)}`;
  }
  return `${digits.slice(0, position)}.${digits.slice(position)}`;
}

export function exactDecimal(value: string): ExactDecimal | undefined {
  const parsed = parseDecimal(value);
  if (parsed === undefined) return undefined;
  if (parsed.zero) {
    return { normalized: "0", integerDigits: 1, scale: 0 };
  }
  const scale =
    BigInt(parsed.digits.length) > parsed.decimalPosition
      ? BigInt(parsed.digits.length) - parsed.decimalPosition
      : 0n;
  const integerDigits =
    parsed.decimalPosition > 1n ? parsed.decimalPosition : 1n;
  if (integerDigits + scale > MAX_DECIMAL_PRECISION) return undefined;
  const normalized = magnitude(parsed);
  return {
    normalized: parsed.negative ? `-${normalized}` : normalized,
    integerDigits: Number(integerDigits),
    scale: Number(scale),
  };
}

export function normalizeDecimal(
  value: string,
  precision: number,
  scale: number,
): string | undefined {
  const parsed = parseDecimal(value);
  if (parsed === undefined) return undefined;
  const fixedScale = (value: string): string => {
    if (scale === 0) return value;
    const [whole, fraction = ""] = value.split(".");
    return `${whole}.${fraction.padEnd(scale, "0")}`;
  };
  if (parsed.zero) return fixedScale("0");
  const significant = parsed.digits.replace(/0+$/u, "");
  const integerDigits =
    parsed.decimalPosition > 0n ? parsed.decimalPosition : 0n;
  const fractionalDigits =
    BigInt(significant.length) > parsed.decimalPosition
      ? BigInt(significant.length) - parsed.decimalPosition
      : 0n;
  if (
    integerDigits > BigInt(precision - scale) ||
    fractionalDigits > BigInt(scale)
  ) {
    return undefined;
  }
  const normalized = fixedScale(magnitude(parsed, significant));
  return parsed.negative ? `-${normalized}` : normalized;
}
