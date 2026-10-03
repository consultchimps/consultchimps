/**
 * An error value a cell holds, such as `#DIV/0!` or `#N/A`, kept apart from
 * text that happens to read the same, so it can be written back as an error
 * cell. Its string form is the error's text, which is also how wide it is.
 */
export class CellError {
  readonly text: string;

  constructor(text: string) {
    this.text = text;
  }

  toString(): string {
    return this.text;
  }
}
