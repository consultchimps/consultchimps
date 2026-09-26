<!-- Generated from the built CLI by scripts/generate-cli-skill-reference.ts. Do not edit; run `pnpm skills:reference`. -->

# ConsultChimps CLI reference

The document commands of `consultchimps` 0.12.0, as the CLI itself
prints them. A flag absent here does not exist in that version.

The `db` commands are left out: they need
native database bindings and no skill carries recipes for them. Run
`consultchimps db --help` against an install to see them.

## consultchimps

```text
Usage: consultchimps [options] [command]

Local-first tools for spreadsheets, persistent databases, presentations, and
PDFs.

Options:
  -V, --version   output the version number
  --json          print one line of machine-readable JSON for automation instead
                  of the detailed explanation
  -h, --help      display help for command

Commands:
  sheets          inspect, combine, or divide Excel workbooks without changing
                  the original files
  pptx            inspect or populate PowerPoint templates without changing the
                  source files
  pdf             split or combine PDF documents without changing the originals
  db              Create persistent local databases, manage schemas, and import
                  workbook submissions
  help [command]  display help for command

Quick start:
  consultchimps db create -o inventory.duckdb
  consultchimps db import prepare inventory.duckdb --input inventory.xlsx -o review.ccplan
  consultchimps sheets inspect clients.xlsx
  consultchimps sheets consolidate "inputs/*.xlsx" -o combined.xlsx
  consultchimps sheets merge "inputs/*.xlsx" -o all-sheets.xlsx
  consultchimps sheets split clients.xlsx -c Region -o by-region
  consultchimps pptx populate --template profile.pptx --data clients.xlsx --sheet Clients --template-slide 1 -o profiles.pptx
  consultchimps pdf split report.pdf -o pages
  consultchimps pdf merge "inputs/*.pdf" -o combined.pdf

Automation with --json:
  Place --json before the command to replace the explanation with one line of
  JSON on stdout. Success prints {"ok":true,"result":...} and failure prints
  {"ok":false,"error":{"message":...,"code":...}} while keeping the nonzero
  exit code. Nothing else is written to stdout, so the output can be piped
  straight into a JSON parser.

  consultchimps --json pdf split report.pdf -o pages

Run consultchimps help <command> or append --help to a command for all options.
```

## consultchimps sheets

```text
Usage: consultchimps sheets [options] [command]

inspect, combine, or divide Excel workbooks without changing the original files

Options:
  -h, --help                         display help for command

Commands:
  unprotect [options] <input>        remove ordinary worksheet and
                                     workbook-structure protection
  merge [options] <inputs...>        copy every worksheet from multiple Excel
                                     workbooks into one workbook, keeping each
                                     sheet separate
  consolidate [options] <inputs...>  stack the rows from every worksheet into
                                     one combined sheet, matching columns by
                                     header
  split [options] <input>            create one new Excel workbook for each
                                     distinct value in a selected column
  inspect [options] <input>          describe what is in an Excel workbook,
                                     creating and changing nothing
  help [command]                     display help for command

Examples:
  consultchimps sheets inspect clients.xlsx
  consultchimps sheets consolidate "inputs/*.xlsx" -o combined.xlsx
  consultchimps sheets merge "inputs/*.xlsx" -o all-sheets.xlsx
  consultchimps sheets split clients.xlsx -c Region -o by-region

Safety:
  Your original Excel workbooks are not changed. ConsultChimps creates new
  output files and refuses to replace existing outputs unless you use --force.

Run consultchimps sheets help <command> for all command options.
```

## consultchimps sheets unprotect

```text
Usage: consultchimps sheets unprotect [options] <input>

remove ordinary worksheet and workbook-structure protection

Arguments:
  input                the source .xlsx or .xlsm workbook

Options:
  -o, --output <path>  where to save the unprotected workbook
  -f, --force          replace the output file if it already exists
  -h, --help           display help for command

Examples:
  consultchimps sheets unprotect protected.xlsx -o unprotected.xlsx

This removes worksheet and workbook-structure protection without changing the source file. Office files encrypted to require a password to open are not supported.
```

## consultchimps sheets merge

```text
Usage: consultchimps sheets merge [options] <inputs...>

copy every worksheet from multiple Excel workbooks into one workbook, keeping
each sheet separate

Arguments:
  inputs               Excel files, folders, or quoted patterns such as
                       "inputs/*.xlsx"

Options:
  -o, --output <path>  where to save the new workbook
  --no-index           do not add the visible Sheet Index worksheet
  --values             replace formulas with their stored values while
                       preserving formatting
  -f, --force          replace the output file if it already exists; use with
                       care
  -h, --help           display help for command

Examples:
  consultchimps sheets merge "inputs/*.xlsx" --values -o all-sheets.xlsx
  consultchimps sheets merge north.xlsx south.xlsx --output all-sheets.xlsx

Every source worksheet remains a separate tab. Sheet Index records source names
and hidden/visible status. Duplicate tab names receive a suffix. --values
removes formulas but always retains cell and workbook formatting.

When you want one combined sheet instead of separate tabs:
  Use consultchimps sheets consolidate to stack the rows from every worksheet
  into a single sheet, matching columns by header.
```

## consultchimps sheets consolidate

```text
Usage: consultchimps sheets consolidate [options] <inputs...>

stack the rows from every worksheet into one combined sheet, matching columns by
header

Arguments:
  inputs                 Excel files, folders, or quoted patterns such as
                         "inputs/*.xlsx"

Options:
  -o, --output <path>    where to save the new consolidated .xlsx workbook
  --sheet <names...>     include only worksheets with these exact names
  --header-row <number>  row containing column names, counted from 1
  --hidden               include hidden worksheets as well as visible ones
  --normalize-headers    match columns whose headers differ only in case,
                         spacing, or punctuation, such as "Failed Checks" and
                         "Failed_Checks"
  --map <file>           JSON column mapping that folds differently named
                         columns into one column each
  --suggest-map <file>   write a draft column mapping built from the headers
                         found, for you to review
  --no-source            leave out columns that identify each row's source file,
                         worksheet, and row
  --output-sheet <name>  name of the worksheet created in the new workbook
                         (default: "Consolidated")
  --values               write stored values instead of formulas while
                         preserving output formatting
  -f, --force            replace the output file if it already exists; use with
                         care
  -h, --help             display help for command

Examples:
  consultchimps sheets consolidate "inputs/*.xlsx" -o combined.xlsx
  consultchimps sheets consolidate north.xlsx south.xlsx --output combined.xlsx

What happens:
  1. ConsultChimps finds the matching Excel files.
  2. It reads every selected, non-empty worksheet.
  3. It matches columns by header name and combines all data rows.
  4. It writes one new workbook and explains exactly what was created.

Your original workbooks are never changed.
Consolidation already writes stored values rather than copying formulas;
--values makes that requirement explicit.

Matching columns that are named differently:
  consultchimps sheets consolidate inputs/ --suggest-map draft.json -o combined.xlsx
  consultchimps sheets consolidate inputs/ --map mapping.json -o combined.xlsx

  --suggest-map writes a draft mapping grouping the headers that differ only in
  case, spacing, or punctuation, for you to read and edit; nothing is applied
  for you. --map then folds every listed spelling into the one column you named.
  A column no mapping entry covers keeps its own name and is reported as a
  warning. Two columns of one worksheet folding into one column stop the run
  rather than quietly losing a value. Use one option or the other, not both.

When you want each worksheet kept as its own tab instead:
  Use consultchimps sheets merge to copy every worksheet into one workbook
  without combining any rows.
```

## consultchimps sheets split

```text
Usage: consultchimps sheets split [options] <input>

create one new Excel workbook for each distinct value in a selected column

Arguments:
  input                     the source .xlsx or .xlsm workbook to divide

Options:
  -c, --column <name>       column whose values decide which rows go into each
                            new workbook
  -o, --output <directory>  folder where the new workbooks will be saved
  --output-dir <directory>  folder where the new workbooks will be saved (alias
                            for --output)
  --sheet <name>            exact name of the worksheet to divide
  --table <name>            use this named Excel Table instead of the
                            worksheet's full used range (preferred)
  --range <name>            use this named range instead of the worksheet's full
                            used range
  --header-row <number>     row containing column names, counted from 1
  --hidden                  allow the selected worksheet to be hidden
  --preserve-workbook       keep the full workbook layout (default without a
                            selector and for --table)
  --no-preserve-workbook    write plain data-only workbooks using the
                            single-source split mode
  --values                  replace formulas with their stored values while
                            preserving formatting
  --strict                  match split values exactly, including case,
                            whitespace, and value type
  --skip-blank              do not create an output group for rows with a blank
                            split-column value
  --prefix <name>           text to place at the start of each output filename
  -f, --force               replace matching output files that already exist;
                            use with care
  -h, --help                display help for command

Examples:
  consultchimps sheets split clients.xlsx -c Region --output-dir by-region
  consultchimps sheets split clients.xlsx --table ClientData --column Region --values
  consultchimps sheets split clients.xlsx --range ClientRange --column Region

What happens:
  1. By default, ConsultChimps finds --column in every worksheet.
  2. It collects distinct non-blank values across all matching worksheets.
  3. It copies the whole workbook once per value and removes other rows.
  4. Worksheets without --column are copied unchanged.
  5. Pivot tables and their caches are removed and reported as a warning: a
     cache holds a private copy of every source row, so it would carry other
     values into each file.

Matching trims surrounding whitespace, ignores case, and treats ordinary
numeric text like the equivalent number. Use --strict for exact matching.
Use --sheet, --table, or --range for the legacy single-source split mode.
Use --no-preserve-workbook only when a compact, data-only result is wanted.

--values removes formulas while retaining their stored results and all
formatting in a preserved workbook. A formula without a stored result becomes
a formatted blank cell and is reported as a warning.

Pro tip: before a table split, prepare the workbook exactly as you want to
deliver it - set each sheet's zoom, place the cursor on cell A1 so every
file opens consistently, add any cover sheet, and save.

Your original workbook is never changed.
```

## consultchimps sheets inspect

```text
Usage: consultchimps sheets inspect [options] <input>

describe what is in an Excel workbook, creating and changing nothing

Arguments:
  input                  the .xlsx or .xlsm workbook to describe

Options:
  --sheet <name>         describe only the worksheet with this exact name;
                         repeat for several
  --header-row <number>  row containing column names, counted from 1
  --hidden               describe hidden worksheets as well as visible ones
  --samples <number>     distinct sample values to report per column, from 0 to
                         5 (default: 5)
  -h, --help             display help for command

Examples:
  consultchimps sheets inspect clients.xlsx
  consultchimps sheets inspect clients.xlsx --hidden --samples 2
  consultchimps sheets inspect --sheet North --sheet South clients.xlsx

What you get:
  1. Each described worksheet, with its visibility, the size of its used range,
     and how many data rows sit below its header row.
  2. The header row an operation would actually use, and every column on it
     with a few of the values stored beneath it.
  3. The Excel Tables and named ranges the described worksheets contain.

Sample values are the first few distinct non-empty values a column stores, at
most five, reported exactly as the workbook holds them: text is quoted, so the
number 1 and the text "1" stay apart. Use --samples 0 for headers only.

No file is created and nothing in the workbook is changed. Run this before
consolidating, merging, or splitting to confirm the worksheet names, header
rows, and column spellings those commands will match on.
```

## consultchimps pptx

```text
Usage: consultchimps pptx [options] [command]

inspect or populate PowerPoint templates without changing the source files

Options:
  -h, --help                             display help for command

Commands:
  inspect-template [options] <template>  list text placeholders on one PowerPoint template slide without creating a file
  populate [options]                     create one populated PowerPoint slide for every nonempty Excel data row
  help [command]                         display help for command

Examples:
  consultchimps pptx inspect-template profile.pptx
  consultchimps pptx populate --template profile.pptx --data clients.xlsx -o profiles.pptx

Safety:
  Your source PowerPoint template and Excel workbook are not changed.
  ConsultChimps creates one new presentation and refuses to replace an existing
  output unless you use --force.
```

## consultchimps pptx inspect-template

```text
Usage: consultchimps pptx inspect-template [options] <template>

list text placeholders on one PowerPoint template slide without creating a file

Arguments:
  template                   the source .pptx template to inspect

Options:
  --template-slide <number>  template slide number, counted from 1 (default: 1)
  -h, --help                 display help for command

Example:
  consultchimps pptx inspect-template profile.pptx

The report identifies valid {{field_name}} placeholders, malformed placeholder
braces, and unsupported placeholder placements. Split-run placeholders are
supported.
```

## consultchimps pptx populate

```text
Usage: consultchimps pptx populate [options]

create one populated PowerPoint slide for every nonempty Excel data row

Options:
  --template <path>          source .pptx file containing {{field_name}}
                             placeholders
  --data <path>              source .xlsx workbook containing the data
  --sheet <name>             exact worksheet name containing the data (default:
                             first worksheet)
  --template-slide <number>  template slide number, counted from 1 (default: 1)
  -o, --output <path>        where to save the new populated .pptx presentation
  --header-row <number>      row containing field names, counted from 1
  -f, --force                replace the output presentation if it already
                             exists; use with care
  -h, --help                 display help for command

Example:
  consultchimps pptx populate --template profile.pptx --data clients.xlsx --output profiles.pptx

Put placeholders such as {{client_name}} or Revenue: {{revenue}} in ordinary
text shapes on the template slide. Each nonempty row below the Excel header
creates one slide, in worksheet order. The first worksheet and first slide are
used unless you select them. Empty cells become empty text.

The output contains only the generated slides. Source files are never changed.
```

## consultchimps pdf

```text
Usage: consultchimps pdf [options] [command]

split or combine PDF documents without changing the originals

Options:
  -h, --help                   display help for command

Commands:
  split [options] <input>      create one new PDF file for every page in a
                               source PDF
  merge [options] <inputs...>  combine several PDF documents into one new PDF
  help [command]               display help for command

Examples:
  consultchimps pdf split report.pdf -o pages
  consultchimps pdf merge "inputs/*.pdf" -o combined.pdf

Safety:
  Your original PDF documents are not changed. ConsultChimps creates new output
  files and refuses to replace existing outputs unless you use --force.

Run consultchimps pdf help <command> for all command options.
```

## consultchimps pdf split

```text
Usage: consultchimps pdf split [options] <input>

create one new PDF file for every page in a source PDF

Arguments:
  input                     the source PDF document to divide

Options:
  -o, --output <directory>  folder where the separate page files will be saved
  --prefix <name>           text to place at the start of each output filename
  -f, --force               replace matching output files that already exist;
                            use with care
  -h, --help                display help for command

Example:
  consultchimps pdf split report.pdf -o pages

What happens:
  ConsultChimps creates one clearly numbered PDF for every page, lists every
  new file, and leaves the source PDF unchanged.
```

## consultchimps pdf merge

```text
Usage: consultchimps pdf merge [options] <inputs...>

combine several PDF documents into one new PDF

Arguments:
  inputs               PDF files, folders, or quoted patterns such as
                       "inputs/*.pdf"

Options:
  -o, --output <path>  where to save the new combined PDF document
  -f, --force          replace the output PDF if it already exists; use with
                       care
  -h, --help           display help for command

Examples:
  consultchimps pdf merge "inputs/*.pdf" -o combined.pdf
  consultchimps pdf merge first.pdf second.pdf --output combined.pdf

What happens:
  ConsultChimps reads the matching PDFs in their resolved order, copies every
  page into one new document, reports the final page count, and leaves every
  source PDF unchanged.
```
