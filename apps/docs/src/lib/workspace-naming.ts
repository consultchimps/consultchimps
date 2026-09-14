function truncateUtf16(value: string, maximumUnits: number): string {
  let units = 0;
  let result = "";
  for (const character of value) {
    if (units + character.length > maximumUnits) break;
    result += character;
    units += character.length;
  }
  return result;
}

export function workspaceWorkingCopyName(
  fileName: string,
  uniqueId = globalThis.crypto.randomUUID(),
): string {
  const extension = fileName.toLowerCase().endsWith(".duckdb")
    ? ".duckdb"
    : ".sqlite";
  const sanitized = fileName
    .replace(/\.[^.]+$/u, "")
    .normalize("NFKC")
    .replace(/[^\p{L}\p{N}._-]+/gu, "_")
    .replace(/^\.+/u, "");
  const stem = truncateUtf16(sanitized, 80);
  return `${stem || "database"}-${uniqueId}${extension}`;
}
