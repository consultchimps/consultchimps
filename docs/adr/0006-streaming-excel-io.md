# Streaming Excel reading and writing

Status: Accepted (2026-10-03).

Consolidation holds every input's cells, the stacked table, and the output
workbook in memory at once, through SheetJS. It needs roughly 130 to 230 MB of
memory per MB of input, so a large consolidation reaches Node's default heap
limit of about 4 GB long before a laptop runs out of memory, and then runs at
100% CPU in garbage collection without finishing (#225). Each input is also
parsed twice, once by SheetJS and once by the document model for dates (#169),
and the output is built three times before a byte is written (#224).
Consolidation is a core capability, so its foundation has to be proven software,
not new code of our own.

## Decision

Excel reading and writing for consolidation moves onto two widely used libraries
with a thin layer of our own on top:

- **fflate** for zip: streaming inflate and deflate in Node and the browser.
- **saxes** for XML: a streaming parser, already pinned in `packages/xlsx`.
- **Our layer** maps XML events to cells and rows (shared strings, styles,
  dates, booleans, errors) and writes worksheet XML row by row. This is the part
  the existing conformance corpus and tests cover.

Consolidation reads in two passes. The first settles the header row, spacer
columns, the column union, and the mapping without keeping rows. The second
reads each worksheet's rows, writes them out, and drops them. Memory is bounded
by the largest worksheet and its shared strings, not by the sum of the inputs.

SheetJS stays for the other operations until each moves over. Removing it
entirely waits on rendering Excel number formats as display text, which the
records reader and PowerPoint populate need.

## Evidence

A spike ran both candidates against today's SheetJS path on neutral generated
data: ten workbooks of 15,000 rows each, and one of 150,000 rows. It also used a
hand-written fixture covering 1904 dates, rich text, sparse rows, cached formula
values, booleans, errors, and a title row.

|                                       | SheetJS today | ExcelJS 4.4.0 streaming | fflate + saxes       |
| ------------------------------------- | ------------- | ----------------------- | -------------------- |
| Peak memory, ten files                | 2,014 MB      | about 250 to 280 MB     | about 240 to 275 MB  |
| Peak memory, 150,000 rows             | 2,755 MB      | 360 MB                  | 169 MB               |
| CPU time, ten files                   | 116 s         | 38 to 52 s              | 33 to 36 s           |
| Cell values match today               | n/a           | yes                     | yes                  |
| Hand-written fixture                  | n/a           | two data losses         | correct              |
| Same bytes across runs and time zones | yes           | no                      | yes                  |
| Runs in a browser Web Worker          | n/a           | no                      | yes, 18.9 KB gzipped |
| Transitive dependencies               | n/a           | 96, some deprecated     | 1                    |

Wall-clock times on the test machine varied up to five times between identical
runs, so CPU time and memory are the measures to trust.

## Rejected

- **ExcelJS streaming.** It kept only the last run of a multi-run inline string.
  When a worksheet came before the shared strings and styles in the zip, the
  order Excel normally writes, it dropped every shared string and date style.
  Its browser build has no streaming API. Its zip timestamps follow the clock
  with no setting to fix them. Its last release was October 2023.
- **The repository's own streaming reader** (`packages/xlsx/src/stream`). It was
  written for database imports, has not been released, and has only been
  exercised by paused work.
- **SheetJS dense mode** (#223). It lowers memory per cell but keeps whole
  workbooks in memory, and the code it tunes is being replaced.

## Consequences

- **One deliberate change to output bytes.** The writer stores text as inline
  strings rather than a shared-strings table, and an error value is written as
  an error cell (`#DIV/0!`) rather than SheetJS's internal code. The CLI and the
  browser still produce byte-identical workbooks. Zip timestamps are fixed,
  because fflate writes local time and ignores its constructor option, so each
  entry's `mtime` is set explicitly.
- **fflate becomes a dependency**, pinned exactly.
- **Reading needs random access.** The reader starts from the zip's central
  directory, so the CLI reads through a file handle and the browser through
  `Blob.slice`.
- **The browser consolidates as the command line does**, reading in pieces and
  writing to disk; see Browser streaming below. Its other operations still read
  whole inputs, in the worker.
- **Behaviour must carry over unchanged.** The current suite already tests most
  of what the spike left untested, and those tests gate the change: merged title
  banners, spacer columns, `uniqueHeaders` naming, malformed packages, and
  non-standard part paths. Encrypted files have no test today and need one
  first.

## Number-format display text

Added 2026-10-04. The records reader, which PowerPoint population reads through,
shows each cell as Excel displays it. That text now comes from **numfmt** (MIT,
no dependencies, pinned exactly), with a thin adapter of our own: the built-in
format table SheetJS used for ids 0 to 81, the 1904 date system (numfmt has
none, so a date format reads the serial plus 1,462, except elapsed-time tokens),
and currency tags such as `[$€-407]` keeping the reader's separators.

A spike compared each candidate with SheetJS's display text, cell by cell.

|                                   | numfmt 3.2.6   | excel-style-dataformatter 2.0.1 |
| --------------------------------- | -------------- | ------------------------------- |
| Conformance corpus, numeric cells | 162 of 162     | 23 of 162                       |
| Generated fixture, 1900 and 1904  | 7,106 of 8,712 | 3,589 of 8,712                  |
| Last release                      | April 2026     | 2017                            |

No other maintained formatter exists on npm apart from SheetJS's own. The
generated fixture covers every built-in id and the hard cases. Most of its
differences are SheetJS departing from Excel: scientific notation, leading
zeros, rounding `-0.5`, values far out of range. Neither matches Excel for dates
before 1900, which Excel shows as `####`.

Review found two numfmt 3.2.6 defects the fixture missed, where it departs from
both Excel and SheetJS. General cuts off a number with 10 or 11 integer digits
and a fraction (`4403928373.5` shows `4403928373`), and a time that rounds up to
midnight keeps the old date. `src/operations/numfmt-guards.ts` corrects exactly
those inputs until upstream fixes them; a test pins numfmt's own answer for
each, so it fails when an upgrade makes the guard unnecessary.

Three smaller differences favour SheetJS and are accepted: a negative number
that rounds to zero shows `0` rather than `-0`, the `A/P` marker shows `AM` or
`PM`, and a fraction format rounding up to a whole number shows `1 1/1`.

## Build order

1. The streaming writer, used for consolidation's output (#224).
2. The two-pass consolidation on the streaming reader, which also ends the
   double parse for this path (#225, #169).
3. The remaining readers (inspection, tables, split input), one at a time. Done.
4. Number-format display text, after which SheetJS can be removed (#241). Done,
   and SheetJS is removed from the repository.

## Browser streaming

Added 2026-10-06. The browser tools held each whole input in the page, copied it
into the worker, and returned each output as one buffer that the page copied
again to download. They now follow the command line: inputs are read in pieces
and outputs are written to disk as they are produced.

- **Inputs.** A page hands the worker each chosen `File`, a reference to the
  file on disk, and never reads it. Consolidation reads through `blobSource`,
  which is `Blob.slice` random access. Operations that still need a whole
  workbook read it once, in the worker.
- **Outputs.** Consolidation writes into the Origin Private File System through
  a sync access handle in the worker, and the page downloads the OPFS `File`,
  which is backed by disk. Where OPFS is missing or refuses (a private window in
  some browsers), the output is collected as `Blob` parts instead, and the page
  says so once it passes 50 MB. A run's OPFS file is deleted when the run after
  next starts, so a download still reading it is not cut short, or as soon as
  the page stops offering it: the results are cleared, the tool is left, or the
  page is closed. The worker holds a Web Lock for each output it offers, and
  whenever a tool page starts its worker it deletes every output whose lock no
  tab holds. An output can stay in the browser's site storage between a tab
  closing before cleanup and the next visit; clearing the site's data removes
  it. Drafting a mapping no longer builds the workbook at all.
- **Not chosen.** The File System Access save picker writes straight to a file
  the visitor names, but only Chromium has it and it must be opened by a click
  before the run starts. OPFS with sync access handles is in Chromium, Firefox
  and Safari by their documentation, with no extra step.

Peak private memory of the whole headless Chromium process tree, above the idle
page, on neutral generated workbooks of 12 columns. A spike first wrote 512 MB
from a worker, 64 KB at a time:

|                                 | Normal profile | Private context |
| ------------------------------- | -------------- | --------------- |
| OPFS sync access handle         | 13 MB          | 781 MB          |
| `Blob` parts                    | 1,164 MB       | 1,163 MB        |
| One `Uint8Array`, then a `Blob` | not measured   | 1,546 MB        |

OPFS stays on disk in a normal profile and is held in memory in a private one,
where it is still no worse than `Blob` parts; the page cannot tell the two
apart, so a private window shows no notice. Firefox wrote and downloaded through
OPFS correctly. Playwright's WebKit build for Windows refused OPFS in both kinds
of profile, so Safari was not checked here; that refusal is the case the `Blob`
fallback covers.

Consolidation in a normal profile, before and after:

| Input                           | Before | After  |
| ------------------------------- | ------ | ------ |
| 10 x 15,000 rows (9.8 MB)       | 158 MB | 123 MB |
| 150,000 rows (9.7 MB)           | 178 MB | 139 MB |
| 600,000 rows (39 MB, 47 MB out) | 345 MB | 182 MB |

What remains is the reader's working set and the engine, which do not grow with
the file. The other operations, before, in a private context:

| Operation, input        | Peak above idle |
| ----------------------- | --------------- |
| Merge, 10 x 15,000 rows | 263 MB          |
| Unprotect, 150,000 rows | 301 MB          |
| Inspect, 150,000 rows   | 950 MB          |

Merge, unprotect, split and PowerPoint population still build or read whole
packages in the library, on the command line too, so they gain only the input
and output transport. Inspection moved; see below. Moving each onto the
streaming reader and writer is separate work.

## Inspection

Added 2026-10-08. Inspection read through the document model, which held the
expanded package, every cell as an object, and three copies of every row: one
for the description and one for each header profile. It now reads each worksheet
once through the streaming reader, as consolidation's first pass does. It keeps
the rows the header rule may still need, which on an ordinary worksheet is a
dozen, then per column the last row holding a value and up to five samples, and
per row only whether it holds content. A row of cells that hold no value, such
as formulas with no cached value, is remembered by its columns, since whether it
counts depends on which columns are kept. A worksheet whose rows are out of
order is gathered whole, as for consolidation. The command line reads the file
through a file handle.

Peak memory of the command line, on neutral generated workbooks:

| Input                                  | Before   | After  |
| -------------------------------------- | -------- | ------ |
| 150,000 rows (9.7 MB)                  | 2,349 MB | 185 MB |
| 10 sheets x 15,000 rows (9.7 MB)       | 1,311 MB | 236 MB |
| 60,000 rows, 400 styles, 12,000 merges | 800 MB   | 170 MB |

CPU time for 150,000 rows fell from 70 to 11 seconds. The descriptions are
identical. Over the 439 distinct workbooks the package's tests open, with seven
option sets, 20 workbooks describe differently, each now as the readers read it:
escapes such as `_x000D_` decoded, phonetic runs left out, samples in row order,
and damaged packages the readers refuse refused here too.

The browser inspector reads the chosen `File` through `Blob.slice` with
`describeWorkbookSource`, as consolidation does. Peak private memory of the
Chromium process tree above the idle page: before, then with the whole file read
in the worker, then read in pieces.

| Input                                  | Before   | Whole file | In pieces |
| -------------------------------------- | -------- | ---------- | --------- |
| 150,000 rows (9.7 MB)                  | 1,468 MB | 121 MB     | 89 MB     |
| 10 sheets x 15,000 rows (9.7 MB)       | 723 MB   | 137 MB     | 129 MB    |
| 60,000 rows, 400 styles, 12,000 merges | 456 MB   | 92 MB      | 95 MB     |
