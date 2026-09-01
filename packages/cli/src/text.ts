/**
 * Terminal-safe rendering of text this process did not author.
 *
 * Worksheet names, headers, cell values, and filenames are input, and input is
 * untrusted. Everything here is written to a terminal, where a control
 * character is an instruction rather than a character: text carrying an escape
 * can move the cursor, recolor or erase the lines already printed, or reach a
 * terminal feature, so a crafted workbook could make the output say something
 * its source does not contain. Excel's own writer stores a control character as
 * the literal text `_x001b_`, but a numeric XML character reference decodes to
 * the character itself, so a hand-built package arrives here with the real
 * thing, and a filename may simply contain one.
 *
 * Only the rendering changes. The values themselves are untouched, so `--json`
 * still reports what the workbook holds, where JSON's own escaping already
 * makes a control character inert.
 */

/** C0 (including DEL) and C1: the ranges a terminal reads as commands. */
function isControlCode(code: number): boolean {
  return code < 0x20 || (code >= 0x7f && code <= 0x9f);
}

/**
 * Text with every control character shown as a visible `\\uXXXX` escape. That
 * also keeps a newline inside a value from forging a line of its own.
 *
 * The test is a comparison rather than a regular expression character class
 * because a pattern carrying these characters is itself what lint warns about,
 * and the bounds above say which ranges are meant.
 */
export function withoutTerminalControls(text: string): string {
  return escapeControls(text, false);
}

/**
 * The same, for text that arrives already laid out in lines: a library's own
 * usage prose, for instance, which is written as one block and would be
 * unreadable with its line breaks escaped. Only the line feed is spared, and
 * only because the text this is used on brings its own; a carriage return is
 * not, because on its own it rewrites the line already printed.
 */
export function withoutTerminalControlsInProse(text: string): string {
  return escapeControls(text, true);
}

function escapeControls(text: string, keepLineFeeds: boolean): string {
  let rendered = "";
  for (const character of text) {
    const code = character.codePointAt(0) ?? 0;
    const escaped = isControlCode(code) && !(keepLineFeeds && code === 0x0a);
    rendered += escaped
      ? `\\u${code.toString(16).toUpperCase().padStart(4, "0")}`
      : character;
  }
  return rendered;
}

/**
 * The same, for text a reader is meant to read back as an exact value.
 *
 * A backslash the text holds is doubled first, so `\\u001B` in the output is
 * always an escape this module wrote and never six characters the value
 * happened to contain.
 */
export function printable(text: string): string {
  return withoutTerminalControls(text.replaceAll("\\", "\\\\"));
}

/**
 * A value with its boundary shown, for a report that lists several on one line.
 *
 * The quotes are the boundary, so a quote inside the text is escaped too:
 * without that, the single stored string `North", 999` would read as the text
 * "North" followed by the number 999, a wrong answer about both the values and
 * their types, and a header spelled `City, State` would read as two columns.
 */
export function quotedValue(text: string): string {
  return `"${printable(text).replaceAll('"', '\\"')}"`;
}
