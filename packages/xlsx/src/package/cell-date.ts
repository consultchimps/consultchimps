/**
 * A date a cell holds, as the serial the writer stores and whether it carries
 * a time of day, so it can be written as a number with a date format. Its
 * string form is the table text it was made from, which is also how wide it
 * is.
 */
export class CellDate {
  /** Days since the 1900 date system's epoch, with the time as a fraction. */
  readonly serial: number;
  /** Whether the moment is anything other than midnight. */
  readonly time: boolean;
  readonly text: string;

  constructor(serial: number, time: boolean, text: string) {
    this.serial = serial;
    this.time = time;
    this.text = text;
  }

  toString(): string {
    return this.text;
  }
}
