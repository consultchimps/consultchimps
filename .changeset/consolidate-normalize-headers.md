---
"@consultchimps/tabular": minor
"@consultchimps/xlsx": minor
"consultchimps": minor
---

Consolidation can now match columns whose headers differ only in case, spacing,
or punctuation. Different systems often export the same schema with different
header conventions - "Failed Checks" in one file, "Failed_Checks" in another,
"Reviewer: Lead Contact" versus "Reviewer_Lead_Contact" - and until now each
spelling became its own mostly-empty output column.

Opt in with `--normalize-headers` on `consultchimps sheets consolidate`, or
`normalizeHeaders: true` on `consolidateWorkbooks` and `unionTables`. Matching
ignores case and treats any run of spaces or punctuation as one separator; the
first spelling seen names the output column. The default behaviour is unchanged.
The tabular package also exports the new `normalizedColumnKey` helper behind
this matching.
