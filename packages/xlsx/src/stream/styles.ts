import { attribute, localName, parseXml } from "./xml.js";
import { readMetadataText, type ZipPackage } from "./zip.js";
import type { FileEntry } from "@zip.js/zip.js";

export interface WorkbookStyles {
  readonly date1904: boolean;
  isDateStyle(styleIndex: number): boolean;
  dateValue(raw: string, styleIndex: number): string | undefined;
}

const BUILT_IN_DATE_FORMATS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
  45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58,
]);

function customDateFormat(format: string): boolean {
  const cleaned = format
    .replace(/"[^"]*"/gu, "")
    .replace(/\\./gu, "")
    .replace(/\[(?!h+\]|m+\]|s+\])[^\]]*\]/giu, "")
    .replace(/_.|\*./gu, "")
    .toLowerCase();
  return /(?:^|[^a-z])[ymdhs]+(?:[^a-z]|$)/u.test(cleaned);
}

function hasTime(format: string | undefined): boolean {
  if (!format) return false;
  const cleaned = format
    .replace(/"[^"]*"/gu, "")
    .replace(/\\./gu, "")
    .toLowerCase();
  return /[hs]/u.test(cleaned) || /\[[hms]+\]/u.test(cleaned);
}

function pad(value: number, length = 2): string {
  return String(value).padStart(length, "0");
}

function serialDate(
  raw: string,
  date1904: boolean,
  includeTime: boolean,
): string | undefined {
  if (!/^-?(?:\d+(?:\.\d*)?|\.\d+)$/u.test(raw)) return undefined;
  const serial = Number(raw);
  if (!Number.isFinite(serial) || serial < 0) return undefined;
  const wholeDays = Math.floor(serial);
  let fractionalMilliseconds = Math.round((serial - wholeDays) * 86_400_000);
  let dayAdjustment = 0;
  if (fractionalMilliseconds === 86_400_000) {
    fractionalMilliseconds = 0;
    dayAdjustment = 1;
  }
  if (!date1904 && wholeDays === 60) {
    const time = new Date(fractionalMilliseconds).toISOString().slice(11, 23);
    return includeTime || fractionalMilliseconds !== 0
      ? `1900-02-29T${time}`
      : "1900-02-29";
  }
  const adjustedDays = date1904
    ? wholeDays + dayAdjustment
    : wholeDays + dayAdjustment - (wholeDays >= 60 ? 1 : 0);
  const epoch = Date.UTC(
    date1904 ? 1904 : 1899,
    date1904 ? 0 : 11,
    date1904 ? 1 : 31,
  );
  const date = new Date(
    epoch + adjustedDays * 86_400_000 + fractionalMilliseconds,
  );
  if (!Number.isFinite(date.getTime())) return undefined;
  const datePart = `${pad(date.getUTCFullYear(), 4)}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}`;
  if (!includeTime && fractionalMilliseconds === 0) return datePart;
  return `${datePart}T${pad(date.getUTCHours())}:${pad(date.getUTCMinutes())}:${pad(date.getUTCSeconds())}.${pad(date.getUTCMilliseconds(), 3)}`;
}

export async function loadWorkbookStyles(
  archive: ZipPackage,
  entry: FileEntry | undefined,
  date1904: boolean,
  signal: AbortSignal | undefined,
): Promise<WorkbookStyles> {
  const customFormats = new Map<number, string>();
  const cellFormats: {
    readonly id: number;
    readonly code?: string | undefined;
  }[] = [];
  if (entry) {
    const xml = await readMetadataText(entry, archive.limits, signal);
    let insideCellFormats = false;
    parseXml(xml, (parser) => {
      parser.on("opentag", (tag) => {
        const name = localName(tag.name);
        if (name === "numFmt") {
          const id = Number(attribute(tag, "numFmtId"));
          const code = attribute(tag, "formatCode");
          if (Number.isSafeInteger(id) && id >= 0 && code !== undefined) {
            customFormats.set(id, code);
          }
        } else if (name === "cellXfs") {
          insideCellFormats = true;
        } else if (insideCellFormats && name === "xf") {
          const id = Number(attribute(tag, "numFmtId") ?? "0");
          if (!Number.isSafeInteger(id) || id < 0) {
            throw new Error("A cell style has an invalid number format ID.");
          }
          cellFormats.push({ id, code: customFormats.get(id) });
        }
      });
      parser.on("closetag", (tag) => {
        if (localName(tag.name) === "cellXfs") insideCellFormats = false;
      });
    });
  }
  const dateStyles = new Set<number>();
  for (const [index, format] of cellFormats.entries()) {
    if (
      BUILT_IN_DATE_FORMATS.has(format.id) ||
      (format.code !== undefined && customDateFormat(format.code))
    ) {
      dateStyles.add(index);
    }
  }
  return {
    date1904,
    isDateStyle(styleIndex) {
      return dateStyles.has(styleIndex);
    },
    dateValue(raw, styleIndex) {
      const format = cellFormats[styleIndex];
      if (!dateStyles.has(styleIndex)) return undefined;
      return serialDate(raw, date1904, hasTime(format?.code));
    },
  };
}
