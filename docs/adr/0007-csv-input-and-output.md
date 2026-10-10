# CSV input and output

Status: Accepted (2026-10-09).

Consultants receive data as CSV as often as workbooks, and nothing in the
repository reads or writes it. Consolidate, inspect, split and merge should take
CSV beside `.xlsx` and `.xlsm`, and give CSV back where one table is the natural
result, on the command line and in the browser alike. A CSV parser is a core
path, so it has to be proven software, not new code of our own.

## Decision

- **Papa Parse** parses, pinned exactly in `@consultchimps/xlsx`. It is MIT, has
  no dependencies, is 20 KB minified, and runs unchanged in Node and a Web
  Worker. We decode the bytes ourselves with `TextDecoder` in streaming mode and
  hand Papa the text, 64 KB at a time, through its readable-stream input. Papa
  documents that input for Node, but it is plain JavaScript that reads any
  object with `read`, `on` and `removeListener`, so the worker hands it one too;
  the pinned version and a browser test hold that in place. The bytes come from
  the same random-access sources as workbooks: a file handle on the command
  line, `Blob.slice` in the worker. Nothing reads a whole file into memory.
- **Opening a CSV reads it once**, in pieces and keeping nothing, before any row
  is used, so the encoding is settled before the first row is decoded:
  - the encoding is the byte order mark's (UTF-8, UTF-16 LE or BE), else UTF-8
    when every byte is valid UTF-8, else Windows-1252, with a warning naming it;
    `--csv-encoding` overrides, except that a byte order mark wins over a choice
    that contradicts it, with a warning. A file with a zero byte and no byte
    order mark is refused, since it is most likely UTF-16 without one;
  - Papa guesses the delimiter (comma, semicolon, tab or pipe) from the first
    ten non-blank rows, or, when those hold no candidate, such as title lines,
    from the first ten that hold one; `--csv-delimiter` overrides it.
- **Line endings.** Papa splits rows on one kind of line ending per file, and
  joined the rows of a file mixing CRLF and LF silently. So rows are split on
  LF, or on CR when the first line ending in the file is a CR not followed by an
  LF, and a carriage return ending a row's last field is read as part of its
  line ending. CRLF, LF and files mixing the two read correctly; the one value
  this changes is a quoted last field whose text itself ends in a carriage
  return. A carriage return on its own anywhere else in a row split on LF could
  be a line ending or text, so it is refused rather than risk joining two rows.
- **Malformed quoting is refused**, naming the row: any quote error Papa
  reports. A row is also limited to 16 MB, which an opening quote never closed
  reaches first; Papa would otherwise carry the rest of the file into one field.
- **A CSV is one worksheet** named after the file stem, so `_source_file` and
  `_source_sheet` stay meaningful. The stem is made a legal sheet name the same
  way everywhere: a character Excel forbids becomes `_`, it is cut to 31
  characters, apostrophes then left at either end are dropped, `History` becomes
  `History_`, and an empty name becomes `Sheet1`. Merge then numbers repeated
  names as it does for workbooks. It has no Excel Tables, named ranges or merged
  cells. A row is a record, so a line break inside quotes stays in its cell; a
  blank line is an empty row; an empty field is a blank cell. The shared header
  rule applies unchanged. Excel's grid (1,048,576 rows, 16,384 columns) and its
  32,767 characters a cell limit only `.xlsx` output, which is refused beyond
  them before anything is written.
- **Typing is text, with opt-ins** that never change a value they do not fully
  match:
  - `--csv-numbers` reads a number from a field that is only an optional sign,
    digits with optional thousands groups, and an optional decimal part. A
    leading zero before another digit, spaces, an exponent, a currency sign or
    more than 15 significant digits keeps the field text, so codes and long
    identifiers survive. `--csv-decimal` (default `.`) and `--csv-thousands`
    (default `,`, or `.` when the decimal is `,`, or `none`) set the separators;
    the two cannot be the same.
  - `--csv-dates iso` reads `yyyy-mm-dd`, with an optional `T` or space and
    `hh:mm`, seconds and up to three decimals, and no time zone. `dmy`, `mdy`
    and `ymd` name the order for dates written with `/`, `-` or `.` and a
    four-digit year, and read ISO too. The day must exist and fall in Excel's
    range, from 1900. Nothing guesses day and month order.
  - A converted date takes the workbook date spelling, so an `.xlsx` output
    stores a real Excel date by the rule of #248, and a CSV output writes ISO.
  - One accepted exception: that rule also makes a date of a text field already
    in the exact spelling `2025-01-31T00:00:00.000Z`, as it does for a
    workbook's text cell.
- **CSV output** goes, a batch of rows at a time, to the same sinks as workbooks
  (a file stream, the browser's OPFS file). Papa's `unparse` quotes each batch;
  we format the values first and join batches ourselves:
  - RFC 4180 quoting, comma delimiters, `\r\n` after every row including the
    last, UTF-8 with a byte order mark so Excel opens it correctly;
    `--csv-bom false` leaves it out;
  - text starting with `=`, `+`, `-` or `@`, after any leading spaces, or with a
    tab or a carriage return, gets a leading `'`, so a spreadsheet does not run
    it as a formula. Text that is only a signed number such as `-12.5` is left
    alone, because it cannot be a formula and CSV to CSV should not change it,
    unless a spreadsheet would change it on reading: a leading zero such as
    `+00123`, or more than 15 significant digits, keeps the apostrophe. Papa's
    own guard would quote those numbers, so we pass it a pattern of our own;
  - dates as ISO text we format: `yyyy-mm-dd`, with `Thh:mm:ss` and any
    milliseconds when there is a time; numbers in JavaScript's shortest
    round-trip form; booleans as `TRUE` and `FALSE`; error cells as their text.
- **Per operation:**
  - consolidate writes CSV when `-o` ends in `.csv`, or with
    `--output-format csv`; one combined table is a natural CSV. An `-o`
    extension that contradicts `--output-format` is refused;
  - split writes one CSV per value with `--output-format csv`, and does so by
    default for a CSV input. A CSV output holds the data only, so it uses the
    compact split; keeping the workbook is refused with CSV output. A workbook
    split to CSV with no source named writes every worksheet that carries the
    column, one file per worksheet and value, named
    `<prefix>-<value> - <sheet>.csv` when more than one worksheet does;
  - merge takes CSV inputs, each becoming a tab named after its stem, but has no
    CSV output: its result is many tabs, which a CSV cannot hold;
  - inspect describes a CSV as its one worksheet and reports the encoding and
    delimiter it used.
- The command line and the browser run the same library code with the same piece
  size, so their outputs are byte for byte identical.

## Evidence

A spike ran Papa Parse 5.7.0 on neutral fixtures through the reader above:
quoted fields with line breaks, doubled quotes and delimiters, one quoted field
spanning two pieces, a UTF-8 character split across pieces, UTF-8 with and
without a byte order mark, UTF-16 LE, Windows-1252 bytes from an Excel "CSV"
save, comma, semicolon and tab delimiters, ragged rows, blank lines, CRLF, mixed
CRLF and LF, and files with and without a final line break. All read correctly
once the delimiter was guessed with blank lines skipped (Papa's guess on the raw
text failed on a two-row tab file, whose final line break counts as a row) and
rows were split on LF as above. An unclosed quote made Papa take the rest of the
file as one field, which is why it is refused.

| 150,000 rows, 12 columns, 13.5 MB | Node, file handle | Chromium worker, `Blob.slice` |
| --------------------------------- | ----------------- | ----------------------------- |
| Time, including the encoding scan | 0.3 to 0.8 s CPU  | 1.6 to 2.1 s                  |
| Peak memory above idle            | 19 MB             | about 24 MB (sampled)         |

## Rejected

- **csv-parse 7.0.3.** It reads mixed line endings when given a list of them and
  guesses delimiters well, but it works on Node's `Buffer`. Its browser build
  bundles `Buffer` and stream shims at 215 KB, ten times Papa's size, for a
  worker that has neither.
- **Papa's own file and encoding handling.** Its `File` input decodes each slice
  separately, which breaks a character split between slices, so decoding stays
  ours.
- **Refusing mixed line endings.** Telling a stray LF from one inside quotes
  needs a quote-aware scan, which would be a parser of our own.
- **Guessing types or date order.** A value read wrongly is worse than one left
  as text, which a mapping or the user can still convert.

## Build order

1. The reader, with encoding and delimiter detection and the typing opt-ins.
2. Consolidate and inspect take CSV.
3. Split and merge take CSV.
4. CSV output.
5. Documentation pages and the command line reference.
