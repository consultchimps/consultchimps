# The conformance corpus

Phase 0 of the migration described in
[`../../ARCHITECTURE.md`](../../ARCHITECTURE.md). The corpus pins what
`@consultchimps/xlsx` does to every tracked workbook structure **today**, so the
layered rewrite (L0-L2) can prove it changed only what it meant to change.
Nothing here touches `src/`.

## Layout

| File                               | Contents                                                                                          |
| ---------------------------------- | ------------------------------------------------------------------------------------------------- |
| `fixtures.ts`                      | The generator. Builds complete OOXML packages by hand, plus the readers the tests assert with.    |
| `split.corpus.test.ts`             | Every split mode (all-worksheet, table, named range, worksheet, bytes) against both bindings.     |
| `merge-consolidate.corpus.test.ts` | Merge and consolidate: what the Phase-1b transplant preserves, and what it strips with a warning. |
| `invariants.corpus.test.ts`        | Preservation, determinism and output-safety guarantees the engine already makes.                  |
| `tier1-gaps.corpus.test.ts`        | The Tier-1 gaps, now all closed: a pin plus a `Tier-1 fix:` test per gap.                         |

## The pairing convention

Fixtures are authored in **pairs**. `buildCorpusWorkbook({ shape: "table" })`
and `buildCorpusWorkbook({ shape: "range" })` place the _same six records in the
same cells_; they differ only in whether an Excel Table part claims the region.

```
A            B          C        D        E         F        (H, I outside the region)
 1  Corpus allocation report                                   <- merged A1:F1
 3  Record       Client     Group    Amount   Doubled   Ratio  <- header row
 4  1            Client A   Alpha    10       20        5
 5  2            Client B   Beta     20       40        10
 6  3            Client C   Alpha    30       60        15      H6 "Alpha side note", merged H6:I6
 7  4            Client D   Beta     40       80        20
 8  5            Client E   Gamma    50       100       25
 9  6            Client F   Alpha    60       120       30
10  Total                            210                       <- table shape only (totals row)
12  Footer note  210                                           <- second data block
```

Worksheets: `Data` (the region), `Summary` (cross-sheet cached aggregates and
the pivot table), `Hidden` (`state="hidden"`, also carries the split column) and
`VeryHidden` (`state="veryHidden"`, does not).

Splitting by `Group` keeps records 1, 3 and 6 in the Alpha output, so rows 4, 6
and 9 compact onto rows 4, 5 and 6. Every dependent-reference assertion in the
corpus is built on that specific movement.

Tests that must hold for both bindings use `it.each(SHAPES)`. Where the two
bindings genuinely differ - the range binding treats every row below the header
as data, so it deletes the totals row and the footer block that the table
binding keeps - the asymmetry gets its own named pin rather than being hidden
behind a branch.

Two structures are table-only by nature and have no range counterpart: the
**totals row** and **structured-reference formulas**. `totalsRow` defaults to
`true` for `shape: "table"` and is ignored for `shape: "range"`; requesting
`formulas: "structured"` on a range fixture throws.

## Composable options

Every structure is an independent flag on `CorpusWorkbookOptions`, so a test can
build the smallest workbook that exhibits the thing it pins:

| Option            | Default             | Adds                                                                  |
| ----------------- | ------------------- | --------------------------------------------------------------------- |
| `shape`           | (required)          | `"table"` adds `xl/tables/table1.xml` and the worksheet `tableParts`. |
| `formulas`        | `"a1"`              | `"none"`, `"a1"` (`D4*2`), or `"structured"` (`DataTable[...]*2`).    |
| `sharedFormula`   | `false`             | The `Ratio` column becomes a shared formula spanning `F4:F9`.         |
| `arrayFormula`    | `false`             | An array formula at `Data!G4`.                                        |
| `uncachedFormula` | `false`             | `Summary!B4`: a formula with no cached `<v>`.                         |
| `totalsRow`       | `shape === "table"` | Row 10 with a `SUBTOTAL`/`SUM` aggregate.                             |
| `dependents`      | `true`              | Merged ranges, conditional formatting, data validation, a hyperlink.  |
| `comments`        | `true`              | `xl/comments1.xml` and its legacy VML drawing, anchored on `B6`.      |
| `definedNames`    | `true`              | Workbook-scoped `CorpusRange`, sheet-scoped `LocalNote`.              |
| `summarySheet`    | `true`              | Cross-sheet formulas with cached values covering every group.         |
| `footerBlock`     | `true`              | The second data block on row 12.                                      |
| `hiddenSheets`    | `true`              | The `Hidden` and `VeryHidden` worksheets.                             |
| `calcChain`       | formulas present    | `xl/calcChain.xml`.                                                   |
| `pivot`           | `false`             | Pivot table, cache definition and cache records, with their rels.     |
| `macro`           | `false`             | A stub `xl/vbaProject.bin` and the macro-enabled content types.       |
| `numberFormat`    | `false`             | A workbook-defined number format on `Amount`, as a third `cellXfs`.   |

The pivot cache records deliberately carry **literal per-row values** rather
than shared-item indexes, so a confidentiality assertion can grep one group's
output for another group's data.

## Why the packages are hand-authored

SheetJS cannot round-trip shared formulas, array formulas, conditional
formatting, data validation, comments, sheet-scoped defined names, very-hidden
sheets, pivot caches or a macro project. Writing the parts directly also keeps
the byte layout stable, so a corpus test fails when the library changes rather
than when SheetJS does. JSZip assembles the package; every entry carries a fixed
date so identical options produce identical bytes.

## Adding a fixture

Follow the ARCHITECTURE.md cookbook ("Fix a discovered edge case"):

1. **Extend the generator, do not fork it.** Add a flag to
   `CorpusWorkbookOptions`, default it `false` (or to whatever keeps existing
   fixtures byte-identical), resolve it in `resolveOptions`, and emit the parts.
   If the structure is a package part, register its content-type override and
   its relationship in `buildCorpusWorkbook` beside the others.
2. **Author it in both shapes** whenever the structure is not inherently a Table
   feature. If it is Table-only, say so in the option's doc comment and throw
   from `resolveOptions` when the range shape asks for it.
3. **Place it so it moves.** A dependent reference that never shifts pins
   nothing. Anchor new structures on row 6 or row 9, which the Alpha split
   relocates, or across `D4:D9`, which the Alpha split shrinks.
4. **Add the assertion, then decide the tier.** If the current behaviour is
   acceptable, write a passing test named `pins: ...` or `invariant: ...`. If it
   is a gap, write the _pair_: a passing `pins:` test recording today's output
   and an adjacent `it.fails` test asserting the required behaviour, with a
   comment naming the gap. Never leave a gap as a skipped test.
5. **Never fix the gap in an operation.** Per the cookbook, invariant fixes
   belong in the model (L1) and Table/range-specific fixes in the binding (L2).

## Flipping a pin

A `pins:` test records behaviour without endorsing it, so replacing one is the
normal way an engine improves. Phase 1b flipped every "merge currently drops X"
pin in `merge-consolidate.corpus.test.ts` at once, and the convention that made
that reviewable is worth repeating:

1. **One assertion per flip.** Keep the test that pinned the loss, rename it
   from `pins:` to `invariant:`, and rewrite its body to assert the structure
   survives. The diff then shows the loss and its replacement side by side.
2. **Say why on the assertion.** Each flipped test carries a `FLIPPED (was ...)`
   comment naming the pin it replaces and the mechanism that makes the new
   behaviour true. A flip with no stated mechanism is a test that will be
   re-flipped by the next regression.
3. **A loss that stays is still pinned, with its warning.** Pivot caches,
   external links and an untravellable macro project remain removed; their tests
   assert the removal AND the warning text, so "silently dropped" can never come
   back disguised as "declared behaviour".
4. **Declare the cell in the same change.** `src/contract.ts` and
   `test/contract.test.ts` move with the corpus; `test/contract.test.ts` fails
   when a structure gains or loses a declared cell without the expected-missing
   list being edited on purpose.

Naming convention for test titles:

- `invariant: ...` - a guarantee the engine already makes and must keep.
- `pins: ...` - current behaviour recorded without endorsing it, lossy outcomes
  included.
- `Tier-1 gap: ...` - an `it.fails` test describing required behaviour that does
  not exist yet.

## Known gaps documented here

All four are listed under Phase 0 in ARCHITECTURE.md, and **all four are now
closed**. The first three were closed by the Tier-1 utilities in `src/tier1/`;
the fourth by Phase 1's move onto the layered engine. Every `it.fails` twin has
become a `Tier-1 fix: ...` test and every pin records the new output, so the
corpus contains no expected-failure tests.

1. ~~Pivot-cache records of one group reach another group's split output.~~
   Fixed: `stripPivotParts` removes every pivot table and cache from each
   output, with a warning, because a cache cannot yet be filtered alongside the
   rows it caches.
2. ~~A values-only split bakes an aggregate computed over every group's rows.~~
   Fixed: `blankStaleCachedFormulas` clears the cached result of any formula
   whose references reach into deleted rows before the values conversion runs,
   so the output shows a reported blank instead of a wrong number.
3. ~~The calculation chain keeps referencing cells the split deleted.~~ Fixed:
   `pruneCalcChain` drops the entries for deleted cells and renumbers the
   survivors, removing the part, its relationship and its content-type override
   when nothing is left.
4. ~~Dependent references (merged ranges, conditional-formatting and
   data-validation `sqref`, hyperlink `ref`, shared-formula `ref`, and A1
   formula text) are not adjusted when rows compact - on **both** bindings.~~
   Fixed in Phase 1: the split edits through L1, whose row-relocation pass
   rewrites all of them in one traversal, together with two references the
   corpus did not previously assert - the `<comment ref>` and the `<x:Row>`
   anchor inside `xl/drawings/vmlDrawing1.vml` that pin a cell comment to a row,
   and the worksheet's declared `dimension/@ref`.

A fifth, smaller finding was pinned in `invariants.corpus.test.ts` rather than
listed as a Tier-1 gap: the all-worksheet split wrote replaced parts with
JSZip's defaults instead of through the deterministic package writer, so its
outputs took the current wall-clock time on every rewritten part and gained
folder entries the source package never had. Their contents were reproducible;
their bytes only within one DOS timestamp tick. That pin is now an invariant:
every write goes through L0, and both split paths are byte-reproducible.
