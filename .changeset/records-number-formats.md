---
"@consultchimps/xlsx": minor
---

`readWorksheetRecords` and its byte twin, which PowerPoint population reads
through, now read through the streaming reader (ADR 0006) and render each cell's
number format with numfmt, a maintained formatter, instead of SheetJS. Two
numfmt defects are corrected locally until upstream fixes them: a General number
with 10 or 11 integer digits and a fraction rounds as Excel shows it, and a time
that rounds up to midnight shows the next day.

A few rare formats now show differently, mostly closer to Excel: scientific
notation such as `0E+0`, leading zeros such as `00.000`, and `-0.5` under `0`,
now `-1`. A negative number that rounds to zero under a format without decimals
shows `0` rather than `-0`, the `A/P` marker shows `AM` or `PM`, and a fraction
that rounds to a whole number shows `1 1/1`. A cell holding text where a number
belongs, or a date cell holding only spaces, now shows as empty instead of
`#NUM!` or `1900-01-00`, and a date before 1900 shows its serial number rather
than nothing. As with the table readers, a carriage return before a line feed is
kept, a formula's cached text is unescaped once, and a worksheet holding an
empty declared date is read.
