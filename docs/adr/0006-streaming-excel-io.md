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
- **The browser is not yet streaming end to end.** It still holds its inputs and
  output in memory. Its peak memory was not measured.
- **Behaviour must carry over unchanged.** The current suite already tests most
  of what the spike left untested, and those tests gate the change: merged title
  banners, spacer columns, `uniqueHeaders` naming, malformed packages, and
  non-standard part paths. Encrypted files have no test today and need one
  first.

## Build order

1. The streaming writer, used for consolidation's output (#224).
2. The two-pass consolidation on the streaming reader, which also ends the
   double parse for this path (#225, #169).
3. The remaining readers (inspection, tables, split input), one at a time.
4. Number-format display text, after which SheetJS can be removed.
