export interface NumberFormatCondition {
  readonly operator: "<" | "<=" | "<>" | "=" | ">" | ">=";
  readonly threshold: number;
}

export interface NumberFormatSection {
  readonly code: string;
  readonly condition?: NumberFormatCondition | undefined;
}

const CONDITION =
  /^(<=|>=|<>|=|<|>)([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)$/u;

function splitSections(format: string): readonly string[] {
  const sections: string[] = [];
  let start = 0;
  let quoted = false;
  for (let index = 0; index < format.length; index += 1) {
    const character = format[index];
    if (character === '"') {
      quoted = !quoted;
      continue;
    }
    if (quoted) continue;
    if (character === "\\" || character === "_" || character === "*") {
      index += 1;
      continue;
    }
    if (character === "[") {
      const end = format.indexOf("]", index + 1);
      if (end < 0) {
        throw new Error("A custom number format has an unterminated bracket.");
      }
      index = end;
      continue;
    }
    if (character === ";") {
      sections.push(format.slice(start, index));
      start = index + 1;
    }
  }
  if (quoted) {
    throw new Error(
      "A custom number format has an unterminated quoted string.",
    );
  }
  sections.push(format.slice(start));
  if (sections.length > 4) {
    throw new Error("A custom number format has more than four sections.");
  }
  return sections;
}

function sectionCondition(code: string): NumberFormatCondition | undefined {
  let quoted = false;
  let found: NumberFormatCondition | undefined;
  for (let index = 0; index < code.length; index += 1) {
    const character = code[index];
    if (character === '"') {
      quoted = !quoted;
      continue;
    }
    if (quoted) continue;
    if (character === "\\" || character === "_" || character === "*") {
      index += 1;
      continue;
    }
    if (character !== "[") continue;
    const end = code.indexOf("]", index + 1);
    if (end < 0) {
      throw new Error("A custom number format has an unterminated bracket.");
    }
    const token = code.slice(index + 1, end).trim();
    const match = CONDITION.exec(token);
    if (match) {
      if (found !== undefined) {
        throw new Error(
          "A custom number-format section has more than one condition.",
        );
      }
      const threshold = Number(match[2]);
      if (!Number.isFinite(threshold)) {
        throw new Error("A custom number format has an invalid condition.");
      }
      found = {
        operator: match[1] as NumberFormatCondition["operator"],
        threshold,
      };
    } else if (/^(?:<|=|>)/u.test(token)) {
      throw new Error("A custom number format has an invalid condition.");
    }
    index = end;
  }
  return found;
}

export function parseNumberFormatSections(
  format: string,
): readonly NumberFormatSection[] {
  const sections = splitSections(format).map((code) => {
    const condition = sectionCondition(code);
    return { code, ...(condition === undefined ? {} : { condition }) };
  });
  if (sections.slice(2).some((section) => section.condition !== undefined)) {
    throw new Error(
      "A custom number format has a condition outside its first two sections.",
    );
  }
  if (
    sections[0]?.condition === undefined &&
    sections[1]?.condition !== undefined
  ) {
    throw new Error(
      "A custom number format has a second-section condition without a first-section condition.",
    );
  }
  return sections;
}

function conditionMatches(
  condition: NumberFormatCondition,
  value: number,
): boolean {
  switch (condition.operator) {
    case "<":
      return value < condition.threshold;
    case "<=":
      return value <= condition.threshold;
    case "<>":
      return value !== condition.threshold;
    case "=":
      return value === condition.threshold;
    case ">":
      return value > condition.threshold;
    case ">=":
      return value >= condition.threshold;
  }
}

export function activeNumberFormatSection(
  sections: readonly NumberFormatSection[],
  value: number,
): NumberFormatSection | undefined {
  const firstCondition = sections[0]?.condition;
  const secondCondition = sections[1]?.condition;
  if (firstCondition !== undefined) {
    if (conditionMatches(firstCondition, value)) return sections[0];
    if (
      secondCondition !== undefined &&
      conditionMatches(secondCondition, value)
    ) {
      return sections[1];
    }
    return secondCondition === undefined ? sections[1] : sections[2];
  }
  if (value > 0) return sections[0];
  if (value < 0) return sections[1] ?? sections[0];
  return sections[2] ?? sections[0];
}
