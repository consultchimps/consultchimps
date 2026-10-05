/**
 * L3: a table value as the table writer stores it. A value the readers hand on
 * as a workbook date, the one spelling `calendarIsoText` writes, goes back to
 * Excel as a date serial with a date format rather than as text (#244). Every
 * caller that turns table values into written rows goes through here.
 */
import { calendarIsoParts, serial1900 } from "../model/calendar.js";
import { CellDate } from "../package/cell-date.js";
import type { WritableCellValue } from "../package/table-writer.js";

/**
 * The value to write for a table value: a date for the workbook date spelling,
 * the value itself otherwise. A date before 1900, which the 1900 date system
 * cannot count, stays text.
 */
export function writableCellValue(value: WritableCellValue): WritableCellValue {
  if (typeof value !== "string") return value;
  const parts = calendarIsoParts(value);
  if (parts === undefined) return value;
  const serial = serial1900(parts);
  if (serial === undefined) return value;
  const time =
    parts.hour !== 0 ||
    parts.minute !== 0 ||
    parts.second !== 0 ||
    parts.millisecond !== 0;
  return new CellDate(serial, time, value);
}
