# Column mapping document

A versioned JSON file that folds differently named source columns onto one
schema before rows are stacked. Passed as `--map <file>`. A change to its shape
gets a new `version`, so a document that parses today keeps parsing.

```json
{
  "version": 1,
  "columns": [
    { "name": "Case_ID", "aliases": ["Reference", "Case Number"] },
    {
      "name": "Opened_On",
      "aliases": ["Run Time"],
      "coercion": { "type": "date", "format": "DD/MM/YYYY" }
    },
    {
      "name": "Amount",
      "aliases": ["Total"],
      "coercion": {
        "type": "number",
        "decimalSeparator": ",",
        "thousandsSeparator": "."
      }
    }
  ],
  "constants": { "Dataset": "quarterly" }
}
```

## Rules

**Aliases match by normalized key.** Case, spacing and punctuation are ignored,
so one alias catches `Case ID`, `case_id` and `CASE-ID`. A canonical column
matches its own name without being repeated in its alias list.

**Canonical names are written verbatim.** The `name` you give is the output
header, exactly as spelled.

**An unclaimed column keeps its own name** and is reported as a warning naming
it. Loud, but nothing is lost. Read those warnings: they are usually a spelling
you did not know existed.

**Two columns of one worksheet folding into one canonical column stop the run**,
naming the file, the worksheet, both columns and the target. Combining them
would silently drop one of the two values.

**Constant columns come last**, after the mapped and unmapped columns and before
the `_source_*` provenance columns.

**`--normalize-headers` does not compete with a mapping.** Aliases always match
by normalized key; the flag governs only the columns no entry claimed.

## Coercions

Two types, both deterministic and both applied per column.

`date` reads the column's **text**, written in the format you declare, and
writes an ISO 8601 date such as `2024-03-09`. Format tokens are the day, month
and year placeholders; every other character is a literal separator.

`number` parses text into a number. `decimalSeparator` and `thousandsSeparator`
are optional; an empty `thousandsSeparator` says the source writes none.

Declare a `date` coercion only for columns Excel holds as **text**. If the
column holds a number, or a value Excel already stores as a date, the run stops
rather than guessing: a bare number is indistinguishable from a case reference
or a quantity, and a stored date needs no coercion at all. Drop the coercion, or
format the column as text in the source.

## Failure codes

| Code                               | Cause                                                                         |
| ---------------------------------- | ----------------------------------------------------------------------------- |
| `TABLE_MAPPING_INVALID`            | the document breaks its own rules, for example two entries claiming one alias |
| `TABLE_MAPPING_COLUMN_COLLISION`   | two columns of one worksheet fold into one canonical column                   |
| `TABLE_MAPPING_CONSTANT_COLLISION` | a constant column has the same name as a mapped column                        |
| `TABLE_MAPPING_COERCION_FAILED`    | a value does not parse under the declared coercion                            |
| `XLSX_MAPPING_FILE_INVALID`        | the file is not valid JSON                                                    |

## Drafting one

`--suggest-map draft.json` writes a draft from the headers the run actually
read, grouping only the spellings that already normalize to the same key. It
proposes no synonyms, applies nothing, and still writes the consolidated
workbook. Review it, add the synonyms it cannot know about, then apply it with
`--map` on a second run.
